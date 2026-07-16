function requestLoginOtp(email) {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const usersSheet = ss.getSheetByName(SHEET_USERS);
  const data = usersSheet.getDataRange().getValues();
  const targetEmail = String(email).trim().toLowerCase();

  // ユーザー存在確認
  const userExists = data.some((r, i) => i > 0 && String(r[COL.EMAIL]).trim().toLowerCase() === targetEmail);
  if (!userExists) return { success: false, message: "未登録のメールアドレスです。" };

  // Slack Botトークン取得
  const botToken = getScriptProperty('SLACK_BOT_TOKEN');
  if (!botToken) return { success: false, message: "システムエラー: Bot Token未設定" };

  try {
    // EmailからSlack IDを特定
    const lookupRes = UrlFetchApp.fetch(`https://slack.com/api/users.lookupByEmail?email=${encodeURIComponent(targetEmail)}`, {
      headers: { "Authorization": "Bearer " + botToken },
      muteHttpExceptions: true
    });
    const lookupJson = JSON.parse(lookupRes.getContentText());
    if (!lookupJson.ok) {
      // よくあるエラーを人間向けに説明
      let userMessage = "Slackアカウントが見つかりません。(Botがワークスペースにいない可能性があります)";
      const err = lookupJson.error || '';
      if (err === 'users_not_found' || err === 'user_not_found') userMessage = '指定したメールアドレスのSlackユーザーが見つかりません。メールアドレスをご確認ください。';
      else if (err === 'not_authed' || err === 'invalid_auth' || err === 'account_inactive') userMessage = 'Botトークンが無効です。Script Properties の SLACK_BOT_TOKEN を確認してください。';
      else if (err === 'missing_scope') userMessage = 'Botに必要な権限がありません。users:read.email 権限を付与してください。';
      else if (err === 'rate_limited') userMessage = 'Slack API の利用制限に達しました。しばらくしてから再試行してください。';
      console.warn('users.lookupByEmail failed:', err, lookupJson);
      return { success: false, message: userMessage, needBotSetup: true };
    }

    const slackUserId = lookupJson.user.id;

    // OTP生成 (6桁数字)
    const otp = Math.floor(100000 + Math.random() * 900000).toString();

    // ScriptPropertiesに一時保存 (有効期限10分想定)
    const otpPayload = JSON.stringify({ code: otp, created: new Date().getTime() });
    PropertiesService.getScriptProperties().setProperty(`OTP_${targetEmail}`, otpPayload);

    // DM送信（認証コードのみ）
    const plainText = `【${APP_NAME}】認証コード: *${otp}*\nこのコードを画面に入力してください。(有効期限10分)`;

    const msgRes = UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": "Bearer " + botToken },
      payload: JSON.stringify({ channel: slackUserId, text: plainText }),
      muteHttpExceptions: true
    });
    const msgJson = JSON.parse(msgRes.getContentText());
    if (!msgJson.ok) {
      console.warn('chat.postMessage failed:', msgJson.error, msgJson);
      throw new Error("Slack DM送信失敗: " + (msgJson.error || 'unknown'));
    }

    return { success: true };
  } catch (e) {
    return { success: false, message: "OTP送信エラー: " + e.message };
  }
}
function verifyLoginOtp(email, code) {
  const targetEmail = String(email).trim().toLowerCase();
  const propKey = `OTP_${targetEmail}`;
  const stored = PropertiesService.getScriptProperties().getProperty(propKey);

  if (!stored) return { success: false, message: "認証コードが無効か期限切れです。" };

  const { code: correctCode, created } = JSON.parse(stored);
  const now = new Date().getTime();

  // 10分有効
  if (now - created > 10 * 60 * 1000) {
    PropertiesService.getScriptProperties().deleteProperty(propKey);
    return { success: false, message: "認証コードの期限が切れています。" };
  }

  if (String(code).trim() !== String(correctCode)) {
    return { success: false, message: "認証コードが間違っています。" };
  }

  // 認証成功: プロパティ削除
  PropertiesService.getScriptProperties().deleteProperty(propKey);

  // セッション発行
  const newSessionToken = Utilities.getUuid();
  const nowTs = Date.now();
  const sessions = _loadStore(SESSIONS_PROP_KEY);
  sessions[newSessionToken] = { email: targetEmail, created: nowTs };
  _saveStore(SESSIONS_PROP_KEY, sessions);
  return { success: true, token: newSessionToken };
}
function getAuthUrl() {
  const clientId = _normalizeSlackCredential(getScriptProperty('SLACK_CLIENT_ID'));
  const redirectUri = _getFrontendOAuthCallbackUrl();
  const userScopes = ["chat:write", "users:read", "users:read.email", "channels:read", "groups:read", "channels:write", "groups:write"].join(",");
  return `https://slack.com/oauth/v2/authorize?client_id=${clientId}&user_scope=${userScopes}&redirect_uri=${encodeURIComponent(redirectUri)}`;
}
function handleSlackOAuthCode(code, redirectUri) {
  const clientId = _normalizeSlackCredential(getScriptProperty('SLACK_CLIENT_ID'));
  const clientSecret = _normalizeSlackCredential(getScriptProperty('SLACK_CLIENT_SECRET'));
  const expectedRedirectUri = _getFrontendOAuthCallbackUrl();
  const actualRedirectUri = String(redirectUri || expectedRedirectUri).trim() || expectedRedirectUri;

  if (!clientId || !clientSecret) {
    return { success: false, message: 'システムエラー: Slack API設定不足' };
  }
  if (actualRedirectUri !== expectedRedirectUri) {
    return { success: false, message: '不正なredirect_uriです' };
  }

  const options = {
    method: "post",
    payload: { client_id: clientId, client_secret: clientSecret, code: code, redirect_uri: expectedRedirectUri }
  };

  try {
    const res = UrlFetchApp.fetch("https://slack.com/api/oauth.v2.access", options);
    const json = JSON.parse(res.getContentText());
    if (!json.ok) return { success: false, message: `Slack認証エラー: ${json.error}` };

    const userSlackToken = json.authed_user.access_token;
    const slackUserId = json.authed_user.id;

    const infoRes = UrlFetchApp.fetch(`https://slack.com/api/users.info?user=${slackUserId}`, {
      headers: { "Authorization": "Bearer " + userSlackToken }
    });
    const infoJson = JSON.parse(infoRes.getContentText());
    if (!infoJson.ok) return { success: false, message: `ユーザー情報取得エラー: ${infoJson.error}` };

    const userEmailRaw = infoJson.user.profile.email;
    const userEmail = String(userEmailRaw || '').trim().toLowerCase();
    const ss = SpreadsheetApp.openById(getSpreadsheetId());

    // ユーザー登録チェック
    const usersSheet = ss.getSheetByName(SHEET_USERS);
    const userData = usersSheet.getDataRange().getValues();
    const userExists = userData.some((r, i) => i > 0 && String(r[COL.EMAIL] || '').trim().toLowerCase() === userEmail);

    if (!userExists) {
      return { success: false, message: `未登録ユーザー (${userEmail})` };
    }

    // セッションとトークンを別々に保存
    const newSessionToken = Utilities.getUuid();
    const now = Date.now();
    const sessions = _loadStore(SESSIONS_PROP_KEY);
    sessions[newSessionToken] = { email: userEmail, created: now };
    _saveStore(SESSIONS_PROP_KEY, sessions);

    const tokensByEmail = _loadStore(TOKENS_BY_EMAIL_PROP_KEY);
    tokensByEmail[userEmail] = { slackToken: userSlackToken, created: now };
    _saveStore(TOKENS_BY_EMAIL_PROP_KEY, tokensByEmail);

    return { success: true, sessionToken: newSessionToken, email: userEmail };
  } catch (e) {
    return { success: false, message: `システムエラー: ${e.message}` };
  }
}
function handleSlackCallback(code) {
  try {
    const callbackUrl = _getFrontendOAuthCallbackUrl();
    const target = callbackUrl + '?code=' + encodeURIComponent(String(code || ''));
    return _buildRedirectHtml(target);
  } catch (e) {
    return HtmlService.createHtmlOutput(`システムエラー: ${e.message}`);
  }
}
function getUserToken(sessionToken) {
  const sessions = _loadStore(SESSIONS_PROP_KEY);
  const entry = sessions[sessionToken];
  if (!entry) throw new Error("セッションが無効です");

  const now = Date.now();
  const created = entry.created || 0;
  if ((now - created) / (1000 * 60 * 60 * 24) > SESSION_DURATION_DAYS) {
    // expired: remove and save
    delete sessions[sessionToken];
    _saveStore(SESSIONS_PROP_KEY, sessions);
    throw new Error("セッション期限切れ");
  }

  const email = entry.email;
  const tokensByEmail = _loadStore(TOKENS_BY_EMAIL_PROP_KEY);
  const tokenEntry = tokensByEmail[email];
  if (!tokenEntry || !tokenEntry.slackToken) throw new Error("Slack連携(Token)がありません。PCからSlackログインを行うか、管理者に連絡してください。");

  return { token: tokenEntry.slackToken, email: email };
}
