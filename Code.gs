/* --------------------------------------------------------------------------
 * 設定 & 定数定義
 * -------------------------------------------------------------------------- */
const APP_NAME = 'Slack-tool';
const APP_HEADER_COLOR = '#1a237e'; // 紺色

const SPREADSHEET_ID = getScriptProperty('SPREADSHEET_ID');
const SHEET_USERS = 'Users';
const SHEET_LOGS = 'Logs';
const SHEET_OPTIONS = 'Options';
const SHEET_TOKENS = 'Tokens'; // 新規テーブル

const SESSION_DURATION_DAYS = 3; // セッション有効期限(日)

// Usersシートの列定義 (0-based index)
const COL = {
  NAME_JP: 0,
  NAME_EN: 1,
  STUDENT_ID: 2,
  GRADE: 3,
  FIELD: 4,
  EMAIL: 5,
  PHONE: 6,
  BIRTHDAY: 7,
  ALMA_MATER: 8,
  ORG_START: 9,
  CAR_OWNER: 24,
  ADMIN: 25
};

// Tokensシートの列定義 (新規)
const COL_TOKENS = {
  SESSION_ID: 0,
  EMAIL: 1,
  SLACK_TOKEN: 2, // User Token (xoxp-...)
  CREATED_AT: 3
};

const HEADER_USERS = [
  '氏名', 'Name', '学籍番号', '学年', '分野', 'メールアドレス', '電話番号', '生年月日', '出身校',
  '所属局1', '所属部門1', '役職1',
  '所属局2', '所属部門2', '役職2',
  '所属局3', '所属部門3', '役職3',
  '所属局4', '所属部門4', '役職4',
  '所属局5', '所属部門5', '役職5',
  '車所有', 'Admin'
];

const HEADER_TOKENS = ['Session ID', 'Email', 'Slack Token', 'Created At'];
const HEADER_OPTIONS = ['学年リスト', '分野リスト', '役職リスト', '所属局リスト', '部門マスタ(局)', '部門マスタ(部門)'];
const HEADER_LOGS = ['Time', 'Sender', 'Recipient', 'Status', 'Details'];

function getScriptProperty(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

/* --------------------------------------------------------------------------
 * 0. 初期セットアップ (マイグレーション & トリガー設定)
 * -------------------------------------------------------------------------- */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // 1. Optionsシート
  let sheetOptions = ss.getSheetByName(SHEET_OPTIONS);
  if (!sheetOptions) sheetOptions = ss.insertSheet(SHEET_OPTIONS);
  if (sheetOptions.getLastRow() === 0) sheetOptions.getRange(1, 1, 1, HEADER_OPTIONS.length).setValues([HEADER_OPTIONS]);

  // 2. Usersシート
  let sheetUsers = ss.getSheetByName(SHEET_USERS);
  if (!sheetUsers) sheetUsers = ss.insertSheet(SHEET_USERS);
  sheetUsers.getRange(1, 1, 1, HEADER_USERS.length).setValues([HEADER_USERS]);

  // 3. Logsシート
  let sheetLogs = ss.getSheetByName(SHEET_LOGS);
  if (!sheetLogs) sheetLogs = ss.insertSheet(SHEET_LOGS);
  if (sheetLogs.getLastRow() === 0) sheetLogs.getRange(1, 1, 1, HEADER_LOGS.length).setValues([HEADER_LOGS]);

  // 4. Tokensシート (新規作成 & 非表示)
  let sheetTokens = ss.getSheetByName(SHEET_TOKENS);
  if (!sheetTokens) {
    sheetTokens = ss.insertSheet(SHEET_TOKENS);
  }
  if (sheetTokens.getLastRow() === 0) {
    sheetTokens.getRange(1, 1, 1, HEADER_TOKENS.length).setValues([HEADER_TOKENS]);
  }
  // マイグレーション時に非表示にする
  sheetTokens.hideSheet();

  // --- A. 書式設定 ---
  const startRow = 2;
  const numRows = 999;
  sheetUsers.getRange(startRow, 3, numRows, 5).setNumberFormat('@');
  sheetUsers.getRange(startRow, 8, numRows, 1).setNumberFormat('yyyy/mm/dd');

  // --- B. 入力規則 ---
  const rangeGradeOpt = sheetOptions.getRange('A2:A');
  const rangeFieldOpt = sheetOptions.getRange('B2:B');
  const rangeRoleOpt  = sheetOptions.getRange('C2:C');
  const rangeOrgOpt   = sheetOptions.getRange('D2:D');
  const rangeDeptOpt  = sheetOptions.getRange('F2:F');

  const buildRule = (range) => SpreadsheetApp.newDataValidation().requireValueInRange(range).setAllowInvalid(true).build();
  const ruleGrade = buildRule(rangeGradeOpt);
  const ruleField = buildRule(rangeFieldOpt);
  const ruleRole  = buildRule(rangeRoleOpt);
  const ruleOrg   = buildRule(rangeOrgOpt);
  const ruleDept  = buildRule(rangeDeptOpt);

  sheetOptions.getRange('E2:E').setDataValidation(ruleOrg);
  sheetUsers.getRange(startRow, 4, numRows, 1).setDataValidation(ruleGrade);
  sheetUsers.getRange(startRow, 5, numRows, 1).setDataValidation(ruleField);

  for (let k = 0; k < 5; k++) {
    const baseCol = 10 + (k * 3);
    sheetUsers.getRange(startRow, baseCol, numRows, 1).setDataValidation(ruleOrg);
    sheetUsers.getRange(startRow, baseCol + 1, numRows, 1).setDataValidation(ruleDept);
    sheetUsers.getRange(startRow, baseCol + 2, numRows, 1).setDataValidation(ruleRole);
  }

  // 車所有 (Y列) と Admin列をチェックボックスに変更（フォールバックあり）
  try {
    sheetUsers.getRange(startRow, COL.CAR_OWNER + 1, numRows, 1).insertCheckboxes();
  } catch (e) {
    const ruleCarOwner = SpreadsheetApp.newDataValidation().requireValueInList(['TRUE', 'FALSE']).setAllowInvalid(true).build();
    sheetUsers.getRange(startRow, COL.CAR_OWNER + 1, numRows, 1).setDataValidation(ruleCarOwner);
  }
  try {
    sheetUsers.getRange(startRow, COL.ADMIN + 1, numRows, 1).insertCheckboxes();
  } catch (e) {
    // ignore
  }

  // --- C. 条件付き書式 ---
  sheetUsers.clearConditionalFormatRules();
  const rules = [];
  const getColLetter = (idx) => {
    let letter = "";
    while (idx > 0) {
      let temp = (idx - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      idx = (idx - temp - 1) / 26;
    }
    return letter;
  };

  const colStudentIdLet = getColLetter(3);
  const rangeStudentId = sheetUsers.getRange(`${colStudentIdLet}2:${colStudentIdLet}`);
  const formulaStudentId = `=AND(${colStudentIdLet}2<>"", NOT(REGEXMATCH(TO_TEXT(${colStudentIdLet}2), "^[0-9]{8}$")))`;
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formulaStudentId).setBackground("#FFFF00").setRanges([rangeStudentId]).build());

  for (let k = 0; k < 5; k++) {
    const colOrgIndex = 10 + (k * 3);
    const colDeptIndex = 11 + (k * 3);
    const colOrgLet = getColLetter(colOrgIndex);
    const colDeptLet = getColLetter(colDeptIndex);
    const range = sheetUsers.getRange(`${colDeptLet}2:${colDeptLet}`);
    const formula = `=AND(${colDeptLet}2<>"", COUNTIFS(INDIRECT("${SHEET_OPTIONS}!$E:$E"), ${colOrgLet}2, INDIRECT("${SHEET_OPTIONS}!$F:$F"), ${colDeptLet}2)=0)`;
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula).setBackground("#FFFF00").setRanges([range]).build());
  }
  sheetUsers.setConditionalFormatRules(rules);

  installTriggers();
  // マイグレーション: 既存の 'TRUE'/'FALSE' 文字列を boolean に変換
  try {
    const lastRowUsers = sheetUsers.getLastRow();
    if (lastRowUsers >= startRow) {
      const carRange = sheetUsers.getRange(startRow, COL.CAR_OWNER + 1, lastRowUsers - startRow + 1, 1);
      const carVals = carRange.getValues().map(r => [(r[0] === 'TRUE' || r[0] === true) ? true : false]);
      carRange.setValues(carVals);

      const adminRange = sheetUsers.getRange(startRow, COL.ADMIN + 1, lastRowUsers - startRow + 1, 1);
      const adminVals = adminRange.getValues().map(r => [(r[0] === 'TRUE' || r[0] === true) ? true : false]);
      adminRange.setValues(adminVals);
    }
  } catch (e) {
    console.warn('Checkbox migration failed:', e.toString());
  }
  console.log("セットアップ完了");
}

function installTriggers() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const triggerFuncName = 'handleSpreadsheetEdit';
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(t => t.getHandlerFunction() === triggerFuncName);
  if (!exists) ScriptApp.newTrigger(triggerFuncName).forSpreadsheet(ss).onEdit().create();
}

/* --------------------------------------------------------------------------
 * Webアプリ & OAuthエンドポイント
 * -------------------------------------------------------------------------- */
function doGet(e) {
  if (e.parameter.code) return handleSlackCallback(e.parameter.code);
  let html = HtmlService.createHtmlOutputFromFile('index').getContent();
  html = html.replace(/{{APP_NAME}}/g, APP_NAME);
  html = html.replace(/{{APP_HEADER_COLOR}}/g, APP_HEADER_COLOR);
  return HtmlService.createHtmlOutput(html)
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* --------------------------------------------------------------------------
 * 1. ログイン & セッション管理 (Tokensシート利用)
 * -------------------------------------------------------------------------- */
function getLoginUser(sessionToken) {
  try {
    if (!sessionToken) return { status: 'guest' };

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const tokensSheet = ss.getSheetByName(SHEET_TOKENS);
    const usersSheet = ss.getSheetByName(SHEET_USERS);

    if (!tokensSheet || !usersSheet) return { status: 'error', message: "DB構成エラー" };

    const tokenData = tokensSheet.getDataRange().getValues();
    let tokenRowIndex = -1;
    let userEmail = "";
    let slackToken = "";

    // セッション検索 & 有効期限チェック
    // 逆順でループして最新を探す＆古い無効なものを掃除するのも手だが、ここではシンプルに検索
    // *期限切れチェック時に行削除を行う*
    const rowsToDelete = [];

    for (let i = 1; i < tokenData.length; i++) {
      if (tokenData[i][COL_TOKENS.SESSION_ID] === sessionToken) {
        const createdAt = new Date(tokenData[i][COL_TOKENS.CREATED_AT]);
        const now = new Date();
        const diffDays = (now - createdAt) / (1000 * 60 * 60 * 24);

        if (diffDays > SESSION_DURATION_DAYS) {
          // 期限切れ: 行を削除してゲスト扱い
          rowsToDelete.push(i + 1);
          continue;
        }

        tokenRowIndex = i + 1;
        userEmail = tokenData[i][COL_TOKENS.EMAIL];
        slackToken = tokenData[i][COL_TOKENS.SLACK_TOKEN];
        break;
      }
    }

    // 期限切れ行の削除 (後ろから削除しないとインデックスがずれるため注意だが、ここでは1件のみ想定)
    if (rowsToDelete.length > 0) {
      rowsToDelete.reverse().forEach(row => tokensSheet.deleteRow(row));
      return { status: 'guest', message: 'セッション有効期限切れ' };
    }

    if (tokenRowIndex === -1) return { status: 'guest' };

    // ユーザー情報取得
    const userData = usersSheet.getDataRange().getValues();
    const userRow = userData.find(r => r[COL.EMAIL] === userEmail);

    if (!userRow) return { status: 'error', message: "ユーザー情報が見つかりません" };

    // Slackのトークン接頭辞は環境や種別により 'xoxp-', 'xoxb-', 'xoxa-', 'xoxc-' 等があるため
    // 'xox' で始まるものは連携済みとみなす
    const hasToken = !!slackToken && slackToken.toString().startsWith('xox');
    return {
      status: 'authorized',
      hasToken: hasToken,
      user: { name: userRow[COL.NAME_JP], email: userEmail }
    };

  } catch (e) {
    return { status: 'error', message: "認証エラー: " + e.toString() };
  }
}

// 1-A. OTPリクエスト (BotからDM送信)
function requestLoginOtp(email) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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

    // DM送信（plain text と blocks 両方でリンクを表示する）
    const plainText = `【${APP_NAME}】認証コード: *${otp}*\nこのコードを画面に入力してください。(有効期限10分)`;
    const shareUrl = getScriptProperty('SHAREABLE_URL');;
    const blocks = [
      { type: 'section', text: { type: 'mrkdwn', text: plainText } },
      { type: 'section', text: { type: 'mrkdwn', text: `または、以下のリンクを開いてください。\n<${shareUrl}>` } }
    ];

    const msgRes = UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": "Bearer " + botToken },
      payload: JSON.stringify({ channel: slackUserId, text: plainText, blocks: blocks }),
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

// 1-B. OTP検証 & セッション発行
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
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tokensSheet = ss.getSheetByName(SHEET_TOKENS);

  const newSessionToken = Utilities.getUuid();
  const timestamp = new Date();

  tokensSheet.appendRow([newSessionToken, targetEmail, "", timestamp]);

  return { success: true, token: newSessionToken };
}

/* --------------------------------------------------------------------------
 * 2. OAuth関連 (PCからのログイン用 - Tokensシートに対応)
 * -------------------------------------------------------------------------- */
function getAuthUrl() {
  const clientId = getScriptProperty('SLACK_CLIENT_ID');
  const scriptUrl = ScriptApp.getService().getUrl();
  const userScopes = ["chat:write", "users:read", "users:read.email", "channels:read", "groups:read", "channels:write", "groups:write"].join(",");
  return `https://slack.com/oauth/v2/authorize?client_id=${clientId}&user_scope=${userScopes}&redirect_uri=${encodeURIComponent(scriptUrl)}`;
}

function getScriptUrl() { return ScriptApp.getService().getUrl(); }

function handleSlackCallback(code) {
  const clientId = getScriptProperty('SLACK_CLIENT_ID');
  const clientSecret = getScriptProperty('SLACK_CLIENT_SECRET');
  const scriptUrl = ScriptApp.getService().getUrl();

  if (!clientId || !clientSecret) return HtmlService.createHtmlOutput("システムエラー: Slack API設定不足");

  const options = {
    method: "post",
    payload: { client_id: clientId, client_secret: clientSecret, code: code, redirect_uri: scriptUrl }
  };

  try {
    const res = UrlFetchApp.fetch("https://slack.com/api/oauth.v2.access", options);
    const json = JSON.parse(res.getContentText());
    if (!json.ok) return HtmlService.createHtmlOutput(`Slack認証エラー: ${json.error}`);

    const userSlackToken = json.authed_user.access_token;
    const slackUserId = json.authed_user.id;

    const infoRes = UrlFetchApp.fetch(`https://slack.com/api/users.info?user=${slackUserId}`, {
      headers: { "Authorization": "Bearer " + userSlackToken }
    });
    const infoJson = JSON.parse(infoRes.getContentText());
    if (!infoJson.ok) return HtmlService.createHtmlOutput(`ユーザー情報取得エラー: ${infoJson.error}`);

    const userEmail = infoJson.user.profile.email;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ユーザー登録チェック
    const usersSheet = ss.getSheetByName(SHEET_USERS);
    const userData = usersSheet.getDataRange().getValues();
    const userExists = userData.some(r => r[COL.EMAIL] === userEmail);

    if (!userExists) return HtmlService.createHtmlOutput(`<h2 style="color:red; text-align:center;">未登録ユーザー (${userEmail})</h2>`);

    // Tokensシートに保存
    const tokensSheet = ss.getSheetByName(SHEET_TOKENS);
    const newSessionToken = Utilities.getUuid();
    const timestamp = new Date();

    tokensSheet.appendRow([newSessionToken, userEmail, userSlackToken, timestamp]);

    return HtmlService.createHtmlOutput(`
    <html>
      <head>
        <meta charset="UTF-8">
        <style>
          body {
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            background: #f3f4f6;
            margin: 0;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Hiragino Kaku Gothic ProN', 'Hiragino Sans', Meiryo, sans-serif;
          }
          .container {
            background: white;
            padding: 40px;
            border-radius: 10px;
            text-align: center;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            max-width: 400px;
            width: 90%;
          }
          h2 {
            color: #059669;
            margin: 0 0 30px 0;
          }
          p {
            color: #6b7280;
            margin-bottom: 20px;
            font-size: 14px;
          }
          button {
            background-color: #2563eb;
            color: white;
            padding: 12px 32px;
            border: none;
            border-radius: 8px;
            font-weight: bold;
            font-size: 14px;
            cursor: pointer;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            transition: background-color 0.3s;
            display: none;
          }
          button:hover {
            background-color: #1d4ed8;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>連携完了</h2>
          <p>認証が完了しました。自動的にリダイレクトしています...</p>
          <button id="toolButton" onclick="redirectToTool()">ツールへ戻る</button>
          <script>
            const sessionToken = '${newSessionToken}';
            const scriptUrl = '${scriptUrl}';

            function redirectToTool() {
              localStorage.setItem('slack_app_session', sessionToken);
              if (scriptUrl) {
                window.top.location.href = scriptUrl;
              } else {
                window.location.reload();
              }
            }

            // 2秒後に自動リダイレクト
            setTimeout(function() {
              localStorage.setItem('slack_app_session', sessionToken);
              if (scriptUrl) {
                window.top.location.href = scriptUrl;
              } else {
                window.location.reload();
              }
            }, 2000);

            // 自動リダイレクト失敗時のためにボタンを表示
            setTimeout(function() {
              document.getElementById('toolButton').style.display = 'inline-block';
            }, 3000);
          </script>
        </div>
      </body>
    </html>
    `);
  } catch (e) {
    return HtmlService.createHtmlOutput(`システムエラー: ${e.message}`);
  }
}

/* --------------------------------------------------------------------------
 * 3. 共通機能 (Token取得ロジック - Tokensシートから取得)
 * -------------------------------------------------------------------------- */
function getUserToken(sessionToken) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tokensSheet = ss.getSheetByName(SHEET_TOKENS);
  const data = tokensSheet.getDataRange().getValues();

  // 最新の状態を確認するため再検索
  const tokenRow = data.find(r => r[COL_TOKENS.SESSION_ID] === sessionToken);

  if (!tokenRow) throw new Error("セッションが無効です");

  // 有効期限チェック
  const createdAt = new Date(tokenRow[COL_TOKENS.CREATED_AT]);
  const now = new Date();
  if ((now - createdAt) / (1000 * 60 * 60 * 24) > SESSION_DURATION_DAYS) {
    throw new Error("セッション期限切れ");
  }

  const slackToken = tokenRow[COL_TOKENS.SLACK_TOKEN];
  if (!slackToken) throw new Error("Slack連携(Token)がありません。PCからSlackログインを行うか、管理者に連絡してください。");

  return { token: slackToken, email: tokenRow[COL_TOKENS.EMAIL] };
}

function getSlackID(token, email) {
  try {
    const res = UrlFetchApp.fetch(`https://slack.com/api/users.lookupByEmail?email=${encodeURIComponent(email)}`, {
      headers: { "Authorization": "Bearer " + token }, muteHttpExceptions: true
    });
    const json = JSON.parse(res.getContentText());
    return json.ok ? json.user.id : null;
  } catch (e) {
    console.warn("getSlackID Error: " + e.message);
    return null;
  }
}

/* --------------------------------------------------------------------------
 * 4. マイページ機能 (プロフィール取得・更新)
 * -------------------------------------------------------------------------- */
function getUserProfile(sessionToken, targetEmail) {
  const login = getLoginUser(sessionToken);
  if (login.status !== 'authorized') throw new Error("認証されていません");

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();

  // 管理者判定はログインユーザーの Users 行の Z 列 (COL.ADMIN)
  const loginRow = data.find(r => r[COL.EMAIL] === login.user.email);
  const isAdmin = loginRow && (loginRow[COL.ADMIN] === 'TRUE' || loginRow[COL.ADMIN] === true);

  const emailToFetch = (targetEmail && isAdmin) ? targetEmail : login.user.email;
  const row = data.find(r => r[COL.EMAIL] === emailToFetch);
  if (!row) throw new Error("データが見つかりません");

  const birthdayVal = row[COL.BIRTHDAY] instanceof Date ? Utilities.formatDate(row[COL.BIRTHDAY], Session.getScriptTimeZone(), 'yyyy-MM-dd') : (row[COL.BIRTHDAY] || '');
  const viewedIsAdmin = row[COL.ADMIN] === 'TRUE' || row[COL.ADMIN] === true;

  return {
    name: row[COL.NAME_JP],
    nameEn: row[COL.NAME_EN],
    email: row[COL.EMAIL],
    studentId: row[COL.STUDENT_ID],
    grade: row[COL.GRADE],
    field: row[COL.FIELD],
    phone: row[COL.PHONE],
    birthday: birthdayVal,
    almaMater: row[COL.ALMA_MATER],
    carOwner: row[COL.CAR_OWNER] === 'TRUE' || row[COL.CAR_OWNER] === true,
    orgs: [
      { org: row[COL.ORG_START], dept: row[COL.ORG_START+1], role: row[COL.ORG_START+2] },
      { org: row[COL.ORG_START+3], dept: row[COL.ORG_START+4], role: row[COL.ORG_START+5] },
      { org: row[COL.ORG_START+6], dept: row[COL.ORG_START+7], role: row[COL.ORG_START+8] },
      { org: row[COL.ORG_START+9], dept: row[COL.ORG_START+10], role: row[COL.ORG_START+11] },
      { org: row[COL.ORG_START+12], dept: row[COL.ORG_START+13], role: row[COL.ORG_START+14] }
    ],
    canEditNameEmail: isAdmin,
    isAdmin: viewedIsAdmin
  };
}

function updateUserProfile(sessionToken, formData, targetEmail) {
  const login = getLoginUser(sessionToken);
  if (login.status !== 'authorized') throw new Error("認証されていません");

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();

  // 管理者判定 (ログインユーザーの ADMIN 列)
  const loginRow = data.find(r => r[COL.EMAIL] === login.user.email);
  const isAdmin = loginRow && (loginRow[COL.ADMIN] === 'TRUE' || loginRow[COL.ADMIN] === true);

  const emailToSave = (targetEmail && isAdmin) ? targetEmail : login.user.email;

  let rowIndex = -1;
  for(let i=1; i<data.length; i++) {
    if (String(data[i][COL.EMAIL]).trim().toLowerCase() === String(emailToSave).trim().toLowerCase()) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) throw new Error("ユーザーが見つかりません");

  // 更新処理
  const normalizeSpace = (s) => {
    if (s === null || s === undefined) return '';
    return String(s).replace(/\u3000/g, ' ').replace(/\s+/g, ' ').trim();
  };

  // 英語名は常に保存
  sheet.getRange(rowIndex, COL.NAME_EN + 1).setValue(normalizeSpace(formData.nameEn || ''));

  sheet.getRange(rowIndex, COL.STUDENT_ID + 1).setValue(formData.studentId || '');
  sheet.getRange(rowIndex, COL.GRADE + 1).setValue(formData.grade || '');
  sheet.getRange(rowIndex, COL.FIELD + 1).setValue(formData.field || '');
  // 電話番号は先頭0を保持するため、セルを文字列書式にしてから明示的に文字列で保存する
  try {
    const phoneRange = sheet.getRange(rowIndex, COL.PHONE + 1);
    phoneRange.setNumberFormat('@');
    phoneRange.setValue(String(formData.phone || ''));
  } catch (e) {
    sheet.getRange(rowIndex, COL.PHONE + 1).setValue(String(formData.phone || ''));
  }

  // 生年月日: フロント側は yyyy-MM-dd を渡す想定。空でなければ Date として保存。
  if (formData.birthday) {
    const d = new Date(formData.birthday);
    if (!isNaN(d.getTime())) sheet.getRange(rowIndex, COL.BIRTHDAY + 1).setValue(d);
  } else {
    sheet.getRange(rowIndex, COL.BIRTHDAY + 1).setValue('');
  }

  sheet.getRange(rowIndex, COL.ALMA_MATER + 1).setValue(formData.almaMater || '');
  sheet.getRange(rowIndex, COL.CAR_OWNER + 1).setValue(formData.carOwner ? true : false);

  // 所属情報 (5セット)
  if (formData.orgs && Array.isArray(formData.orgs)) {
    for (let k = 0; k < 5; k++) {
      if (k < formData.orgs.length) {
        // 所属局1 (k===0) の編集は管理者のみ許可
        if (k === 0 && !isAdmin) continue;
        const o = formData.orgs[k];
        const baseCol = COL.ORG_START + (k * 3) + 1;
        sheet.getRange(rowIndex, baseCol).setValue(o.org || "");
        sheet.getRange(rowIndex, baseCol + 1).setValue(o.dept || "");
        sheet.getRange(rowIndex, baseCol + 2).setValue(o.role || "");
      }
    }
  }

  // 管理者は氏名・メール・管理フラグの編集が可能
  if (isAdmin) {
    if (formData.name) sheet.getRange(rowIndex, COL.NAME_JP + 1).setValue(normalizeSpace(formData.name));
    if (formData.email) sheet.getRange(rowIndex, COL.EMAIL + 1).setValue(String(formData.email).trim().toLowerCase());
    if (typeof formData.isAdmin !== 'undefined') sheet.getRange(rowIndex, COL.ADMIN + 1).setValue(formData.isAdmin ? true : false);
  }

  return { success: true };
}

/* --------------------------------------------------------------------------
 * 5. DM送信 & チャンネル招待
 * -------------------------------------------------------------------------- */
/**
 * 管理者向け: 新規ユーザー作成
 * @param {string} sessionToken
 * @param {object} userObj
 */
function createUser(sessionToken, userObj) {
  const login = getLoginUser(sessionToken);
  if (login.status !== 'authorized') throw new Error('認証されていません');

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const usersSheet = ss.getSheetByName(SHEET_USERS);
  if (!usersSheet) throw new Error('Users シートが見つかりません');

  // 管理者判定
  const allUsers = usersSheet.getDataRange().getValues();
  const loginRow = allUsers.find(r => r[COL.EMAIL] === login.user.email);
  const isAdmin = loginRow && (loginRow[COL.ADMIN] === 'TRUE' || loginRow[COL.ADMIN] === true);
  if (!isAdmin) throw new Error('権限がありません');

  // 必須チェック
  const name = (userObj.name || '').toString().trim();
  const email = (userObj.email || '').toString().trim().toLowerCase();
  if (!name) throw new Error('氏名を入力してください');
  if (!email) throw new Error('メールアドレスを入力してください');

  // 重複チェック
  for (let i = 1; i < allUsers.length; i++) {
    if (String(allUsers[i][COL.EMAIL]).trim().toLowerCase() === email) {
      throw new Error('同じメールアドレスのユーザーが既に存在します');
    }
  }

  // 行データ作成
  const row = new Array(HEADER_USERS.length).fill('');
  row[COL.NAME_JP] = name;
  row[COL.NAME_EN] = userObj.nameEn || '';
  row[COL.STUDENT_ID] = userObj.studentId || '';
  row[COL.GRADE] = userObj.grade || '';
  row[COL.FIELD] = userObj.field || '';
  row[COL.EMAIL] = email;
  row[COL.PHONE] = userObj.phone || '';
  if (userObj.birthday) {
    const d = new Date(userObj.birthday);
    if (!isNaN(d.getTime())) row[COL.BIRTHDAY] = d;
  }
  row[COL.ALMA_MATER] = userObj.almaMater || '';

  // 所属 (5セット)
  if (userObj.orgs && Array.isArray(userObj.orgs)) {
    for (let k = 0; k < 5; k++) {
      const base = COL.ORG_START + (k * 3);
      if (k < userObj.orgs.length) {
        const o = userObj.orgs[k] || {};
        row[base] = o.org || '';
        row[base + 1] = o.dept || '';
        row[base + 2] = o.role || '';
      }
    }
  }

  // チェックボックス列
  row[COL.CAR_OWNER] = userObj.carOwner ? true : false;
  row[COL.ADMIN] = userObj.isAdmin ? true : false;

  usersSheet.appendRow(row);
  // appendRow の後に、電話番号セルを文字列書式で上書きして先頭0を確実に保持する
  try {
    const lastRow = usersSheet.getLastRow();
    const phoneRangeNew = usersSheet.getRange(lastRow, COL.PHONE + 1);
    phoneRangeNew.setNumberFormat('@');
    phoneRangeNew.setValue(String(userObj.phone || ''));
  } catch (e) {
    // フォールバック: 何もしない
  }
  return { success: true };
}

function sendDMs(sessionToken, message, recipients) {
  const { token, email: senderEmail } = getUserToken(sessionToken);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const logSheet = ss.getSheetByName(SHEET_LOGS);
  let successCount = 0;
  const failedList = [];
  const time = new Date();

  recipients.forEach((r) => {
    try {
      const uid = getSlackID(token, r.email);
      if (!uid) throw new Error("Slackアカウントなし");
      const text = message.replace(/{mention}/g, `<@${uid}>`);
      // 共有可能な Slack へのリンクを添付
      const shareUrl = `https://slack.com/app_redirect?channel=${uid}`;
      const fullText = `${text}\n\nまたは、以下のリンクを開いてください。\n${shareUrl}`;
      // blocks を使って確実にリンクが表示されるようにする
      const blocks = [
        { type: 'section', text: { type: 'mrkdwn', text: text } },
        { type: 'section', text: { type: 'mrkdwn', text: `または、以下のリンクを開いてください。\n<${shareUrl}>` } }
      ];
      const res = UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", {
        method: "post",
        contentType: "application/json",
        headers: { "Authorization": "Bearer " + token },
        payload: JSON.stringify({ channel: uid, text: fullText, blocks: blocks }),
        muteHttpExceptions: true
      });
      const json = JSON.parse(res.getContentText());
      if (!json.ok) throw new Error(json.error || "Unknown Error");
      successCount++;
      logSheet.appendRow([time, senderEmail, r.email, "DM Success", ""]);
    } catch (e) {
      failedList.push({ email: r.email, error: e.message });
      logSheet.appendRow([time, senderEmail, r.email, "DM Failed", e.message]);
    }
    Utilities.sleep(500);
  });
  return { success: successCount, failed: failedList };
}

function getChannels(sessionToken) {
  const { token } = getUserToken(sessionToken);
  const fetch = function(types) {
    let channels = [];
    let cursor = "";
    do {
      const url = `https://slack.com/api/conversations.list?types=${types}&exclude_archived=true&limit=200&cursor=${cursor}`;
      const res = UrlFetchApp.fetch(url, { headers: { "Authorization": "Bearer " + token }, muteHttpExceptions: true });
      const json = JSON.parse(res.getContentText());
      if (!json.ok) throw new Error(json.error);
      if (json.channels) channels = channels.concat(json.channels);
      cursor = (json.response_metadata && json.response_metadata.next_cursor) ? json.response_metadata.next_cursor : "";
    } while (cursor);
    return channels;
  };

  try {
    try {
      return fetch("public_channel,private_channel")
        .map(c => ({ id: c.id, name: c.name, is_private: c.is_private }))
        .sort((a, b) => a.name.localeCompare(b.name));
    } catch (e) {
      if (e.message === 'missing_scope') {
        return fetch("public_channel")
          .map(c => ({ id: c.id, name: c.name, is_private: c.is_private }))
          .sort((a, b) => a.name.localeCompare(b.name));
      }
      throw e;
    }
  } catch (e) {
    throw new Error("チャンネル一覧取得失敗: " + e.message);
  }
}

function inviteToChannel(sessionToken, channelId, recipients) {
  const { token, email: senderEmail } = getUserToken(sessionToken);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const logSheet = ss.getSheetByName(SHEET_LOGS);
  let successCount = 0;
  const failedList = [];
  const time = new Date();

  recipients.forEach((r) => {
    try {
      const uid = getSlackID(token, r.email);
      if (!uid) throw new Error("Slackアカウントなし");
      const res = UrlFetchApp.fetch("https://slack.com/api/conversations.invite", {
        method: "post", contentType: "application/json", headers: { "Authorization": "Bearer " + token },
        payload: JSON.stringify({ channel: channelId, users: uid }), muteHttpExceptions: true
      });
      const json = JSON.parse(res.getContentText());
      if (!json.ok) {
        if (json.error === 'already_in_channel') throw new Error("既に参加済みです");
        else if (json.error === 'not_in_channel') throw new Error("実行者(あなた)がこのチャンネルに参加していません");
        else throw new Error(json.error);
      }
      successCount++;
      logSheet.appendRow([time, senderEmail, r.email, "Invite Success (User)", channelId]);
    } catch (e) {
      failedList.push({ email: r.email, error: e.message });
      logSheet.appendRow([time, senderEmail, r.email, "Invite Failed", `${channelId}: ${e.message}`]);
    }
    Utilities.sleep(500);
  });
  return { success: successCount, failed: failedList };
}

/* --------------------------------------------------------------------------
 * 6. 検索 & ユーティリティ
 * -------------------------------------------------------------------------- */
function getSearchOptions() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const optSheet = ss.getSheetByName(SHEET_OPTIONS);
  if (!optSheet) return { grades: [], fields: [], roles: [], orgs: [], deptMaster: [] };
  const data = optSheet.getDataRange().getValues();
  const options = { grades: [], fields: [], roles: [], orgs: [], deptMaster: [] };
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) options.grades.push(data[i][0]);
    if (data[i][1]) options.fields.push(data[i][1]);
    if (data[i][2]) options.roles.push(data[i][2]);
    if (data[i][3]) options.orgs.push(data[i][3]);
    if (data[i][4] && data[i][5]) options.deptMaster.push({ org: data[i][4], dept: data[i][5] });
  }
  return options;
}

function searchRecipients(criteria) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  const results = [];
  const q = criteria.query ? criteria.query.toLowerCase() : "";
  const filterGrade = criteria.grade || "";
  const filterField = criteria.field || "";
  const filterOrg = criteria.org || "";
  const filterDept = criteria.dept || "";
  const filterRole = criteria.role || "";

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const nameJp = row[COL.NAME_JP];
    const nameEn = row[COL.NAME_EN];
    const studentId = row[COL.STUDENT_ID] || "";
    const grade = row[COL.GRADE];
    const field = row[COL.FIELD];
    const email = row[COL.EMAIL];
    const almaMater = row[COL.ALMA_MATER] || "";
    const searchString = `${nameJp} ${nameEn} ${email} ${almaMater} ${studentId}`.toLowerCase();

    if (q && !searchString.includes(q)) continue;
    if (filterGrade && grade !== filterGrade) continue;
    if (filterField && field !== filterField) continue;

    let isOrgMatch = !filterOrg;
    let isDeptMatch = !filterDept;
    let isRoleMatch = !filterRole;
    if (filterOrg || filterDept || filterRole) { isOrgMatch = false; isDeptMatch = false; isRoleMatch = false; }

    const depts = [];
    for (let k = 0; k < 5; k++) {
      const start = COL.ORG_START + (k * 3);
      if (start + 2 >= row.length) break;
      const org = row[start];
      const dept = row[start + 1];
      const role = row[start + 2];
      if (org || dept || role) depts.push([org, dept, role].filter(Boolean).join(" "));
      if (filterOrg && org === filterOrg) isOrgMatch = true;
      if (filterDept && dept === filterDept) isDeptMatch = true;
      if (filterRole && role === filterRole) isRoleMatch = true;
    }

    if (filterOrg && !isOrgMatch) continue;
    if (filterDept && !isDeptMatch) continue;
    if (filterRole && !isRoleMatch) continue;

    results.push({
      name: nameJp, email: email, department: depts.join(", ") || "所属なし", grade: grade, field: field
    });
  }
  return results.slice(0, 50);
}

function handleSpreadsheetEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_USERS) return;
  const range = e.range;
  const col = range.getColumn();
  const row = range.getRow();
  if (row < 2) return;

  // 所属局列（10, 13, 16, 19, 22）の編集を検出
  const startCol = 10;
  if (col < startCol || col > 22) return;
  if ((col - startCol) % 3 !== 0) return;

  const orgName = e.value;
  const deptRange = sheet.getRange(row, col + 1); // 所属部門の隣のセル

  // 所属局が空の場合、所属部門もクリア
  if (!orgName) {
    deptRange.clearContent();
    deptRange.clearDataValidation();
    return;
  }

  const ss = e.source;
  const optSheet = ss.getSheetByName(SHEET_OPTIONS);
  const lastRow = optSheet.getLastRow();
  if (lastRow < 2) return;

  // Optionsシートの部門マスタ（E列:局, F列:部門）から該当する部門を抽出
  const masterData = optSheet.getRange(2, 5, lastRow - 1, 2).getValues();
  const filteredDepts = masterData.filter(r => r[0] === orgName).map(r => r[1]).filter(String);

  // 該当する部門のみのドロップダウンリストを設定
  if (filteredDepts.length > 0) {
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(filteredDepts).setAllowInvalid(true).build();
    deptRange.setDataValidation(rule);
  } else {
    deptRange.clearDataValidation();
  }
}

/* --------------------------------------------------------------------------
 * 管理者列初期化バッチ
 * Users シートのヘッダに 'Admin' を追加し、Z列の空セルを FALSE に設定します
 * -------------------------------------------------------------------------- */
function initAdminColumnDefaults() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) throw new Error('Users シートが見つかりません');

  const lastCol = Math.max(sheet.getLastColumn(), COL.ADMIN + 1);
  const headerRange = sheet.getRange(1, 1, 1, lastCol);
  const headers = headerRange.getValues()[0];

  // ヘッダが短ければ拡張
  if (headers.length < COL.ADMIN + 1) {
    const newHeaders = headers.slice();
    for (let i = headers.length; i < COL.ADMIN; i++) newHeaders[i] = '';
    newHeaders[COL.ADMIN] = 'Admin';
    sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  } else {
    // 既存ヘッダに Admin をセット（上書きでも安全）
    sheet.getRange(1, COL.ADMIN + 1).setValue('Admin');
  }

  // データ行の Admin 列が空の行は FALSE に設定
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, message: 'ヘッダのみです' };

  const colRange = sheet.getRange(2, COL.ADMIN + 1, lastRow - 1, 1);
  const colVals = colRange.getValues();
  let updated = 0;
  for (let i = 0; i < colVals.length; i++) {
    const v = colVals[i][0];
    if (v === '' || v === null || typeof v === 'undefined') {
      colVals[i][0] = false;
      updated++;
    } else if (v === 'TRUE') {
      colVals[i][0] = true;
    } else if (v === 'FALSE') {
      colVals[i][0] = false;
    }
  }
  if (updated > 0) colRange.setValues(colVals);
  // Admin列をチェックボックス化（データ行）
  try {
    sheet.getRange(2, COL.ADMIN + 1, lastRow - 1, 1).insertCheckboxes();
  } catch (e) {
    // ignore
  }
  // CAR_OWNER列もチェックボックス化（データ行）
  try {
    sheet.getRange(2, COL.CAR_OWNER + 1, lastRow - 1, 1).insertCheckboxes();
  } catch (e) {
    // ignore
  }
  return { success: true, updated: updated };
}

/* --------------------------------------------------------------------------
 * 7. 名簿出力機能
 * - 共有フォルダIDはスクリプトプロパティ 'EXPORT_SHARED_FOLDER_ID' に保存する想定
 * - フロントからはフォルダ一覧取得と、選択した項目でスプレッドシートを生成するAPIを提供
 * -------------------------------------------------------------------------- */

function listExportFolders() {
  try {
    const rootId = getScriptProperty('EXPORT_SHARED_FOLDER_ID');
    if (!rootId) return { success: false, message: '共有フォルダが設定されていません' };
    // Use Advanced Drive service to list child folders
    const folders = [];
    const q = "'" + rootId + "' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false";
    const res = Drive.Files.list({ q: q, fields: 'items(id,title)' });
    if (res && res.items) {
      res.items.forEach(item => folders.push({ id: item.id, name: item.title }));
    }
    // get root title
    let rootName = '';
    try {
      const r = Drive.Files.get(rootId, { fields: 'id,title' });
      rootName = r && r.title ? r.title : '';
    } catch (e) {
      rootName = '';
    }
    return { success: true, folders: [{ id: rootId, name: rootName }].concat(folders) };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function createRosterSpreadsheet(sessionToken, selectedFields, folderId, filename) {
  // sessionToken: to validate user and permissions
  try {
    const login = getLoginUser(sessionToken);
    if (login.status !== 'authorized') throw new Error('認証されていません');

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const usersSheet = ss.getSheetByName(SHEET_USERS);
    if (!usersSheet) throw new Error('Users シートが見つかりません');

    // 管理者判定
    const usersData = usersSheet.getDataRange().getValues();
    const loginRow = usersData.find(r => r[COL.EMAIL] === login.user.email);
    const isAdmin = loginRow && (loginRow[COL.ADMIN] === 'TRUE' || loginRow[COL.ADMIN] === true);

    // no special "すべての項目" handling anymore

    // Build mapping from requested labels to column indexes in Users sheet
    const indices = [];
    const headersOut = [];

    const pushIf = (label, idx) => { headersOut.push(label); indices.push(idx); };

    // helper to add org/dept/role sequences
    const pushOrgSeq = (baseIdx, labelBase) => {
      for (let k = 0; k < 5; k++) {
        const idx = baseIdx + (k * 3);
        pushIf(labelBase + String(k + 1), idx);
      }
    };

    // map requested labels to columns
    selectedFields = selectedFields || [];
    for (let f of selectedFields) {
      if (f === '氏名') pushIf('氏名', COL.NAME_JP);
      else if (f === 'Name') pushIf('Name', COL.NAME_EN);
      else if (f === '学年') pushIf('学年', COL.GRADE);
      else if (f === '分野') pushIf('分野', COL.FIELD);
      else if (f === 'メールアドレス') pushIf('メールアドレス', COL.EMAIL);
      else if (f === '所属局1～5' || f === '所属局1') pushOrgSeq(COL.ORG_START, '所属局');
      else if (f === '所属部門1～5' || f === '所属部門1') pushOrgSeq(COL.ORG_START + 1, '所属部門');
      else if (f === '役職1～5' || f === '役職1') pushOrgSeq(COL.ORG_START + 2, '役職');
      else if (f === '車所有') pushIf('車所有', COL.CAR_OWNER);
      else if (f === 'Admin') pushIf('Admin', COL.ADMIN);
    }

    if (indices.length === 0) throw new Error('出力項目が選択されていません');

    // Read users data rows and build output rows
    const outRows = [];
    for (let i = 1; i < usersData.length; i++) {
      const row = usersData[i];
      // skip empty rows (メールアドレスが空なら無視)
      if (!row[COL.EMAIL]) continue;
      const outRow = indices.map(ci => row[ci] === undefined || row[ci] === null ? '' : row[ci]);
      outRows.push(outRow);
    }

    // Create new spreadsheet via Sheets API (Advanced Service)
    const title = filename || ('名簿_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm'));
    const resource = { properties: { title: title } };
    const created = Sheets.Spreadsheets.create(resource);
    const newId = created.spreadsheetId;
    const sheetName = (created.sheets && created.sheets[0] && created.sheets[0].properties && created.sheets[0].properties.title) ? created.sheets[0].properties.title : 'Sheet1';

    // write header and rows via Sheets API
    try {
      Sheets.Spreadsheets.Values.update({ values: [headersOut] }, newId, sheetName + '!A1', { valueInputOption: 'RAW' });
      if (outRows.length > 0) {
        Sheets.Spreadsheets.Values.update({ values: outRows }, newId, sheetName + '!A2', { valueInputOption: 'RAW' });
      }
    } catch (e) {
      console.warn('Sheets write failed:', e.toString());
    }

    // Move file into target folder using Drive API (Advanced Drive service)
    try {
      // add to target and remove from root
      Drive.Files.update({}, newId, { addParents: folderId, removeParents: 'root' });
    } catch (e) {
      console.warn('フォルダ移動失敗:', e.toString());
    }

    return { success: true, url: 'https://docs.google.com/spreadsheets/d/' + newId, id: newId };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * CSV 出力版: Drive に保存せず CSV 文字列を返す
 * フロントでダウンロード処理を行う
 */
function createRosterCsv(sessionToken, params) {
  try {
    const login = getLoginUser(sessionToken);
    if (login.status !== 'authorized') throw new Error('認証されていません');

    params = params || {};
    const selectedFields = params.selectedFields || [];
    const filter = params.filter || { type: 'all' };

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const usersSheet = ss.getSheetByName(SHEET_USERS);
    if (!usersSheet) throw new Error('Users シートが見つかりません');

    // 管理者判定
    const usersData = usersSheet.getDataRange().getValues();
    const loginRow = usersData.find(r => r[COL.EMAIL] === login.user.email);
    const isAdmin = loginRow && (loginRow[COL.ADMIN] === 'TRUE' || loginRow[COL.ADMIN] === true);

    // no special "すべての項目" handling

    // 出力項目マッピング
    const adminOnlyFields = ['学籍番号','電話番号','生年月日','出身校','車所有','Admin'];
    // 管理者権限のないユーザーが管理者専用項目を要求していないか確認
    if (!isAdmin) {
      for (const f of selectedFields || []) {
        if (adminOnlyFields.indexOf(f) !== -1) throw new Error('管理者権限が必要な項目が含まれています');
      }
    }

    const allowedIdxSet = new Set();
    const add = (idx) => { if (typeof idx === 'number' && idx >= 0) allowedIdxSet.add(idx); };

    for (let f of selectedFields || []) {
      if (f === '氏名') add(COL.NAME_JP);
      else if (f === 'Name') add(COL.NAME_EN);
      else if (f === '学籍番号') add(COL.STUDENT_ID);
      else if (f === '学年') add(COL.GRADE);
      else if (f === '分野') add(COL.FIELD);
      else if (f === 'メールアドレス') add(COL.EMAIL);
      else if (f === '電話番号') add(COL.PHONE);
      else if (f === '生年月日') add(COL.BIRTHDAY);
      else if (f === '出身校') add(COL.ALMA_MATER);
      else if (f === '所属局1') { add(COL.ORG_START + 0 * 3); }
      else if (f === '所属部門1') { add(COL.ORG_START + 0 * 3 + 1); }
      else if (f === '役職1') { add(COL.ORG_START + 0 * 3 + 2); }
      else if (f === '所属局2～5') {
        for (let k = 1; k <= 4; k++) add(COL.ORG_START + k * 3);
      } else if (f === '所属部門2～5') {
        for (let k = 1; k <= 4; k++) add(COL.ORG_START + k * 3 + 1);
      } else if (f === '役職2～5') {
        for (let k = 1; k <= 4; k++) add(COL.ORG_START + k * 3 + 2);
      } else if (f === '車所有') add(COL.CAR_OWNER);
      else if (f === 'Admin') add(COL.ADMIN);
    }

    const indices = Array.from(allowedIdxSet).sort((a,b)=>a-b);
    const headersOut = indices.map(i => HEADER_USERS[i] || '');

    if (indices.length === 0) throw new Error('出力項目が選択されていません');

    // フィルタ処理
    const outRows = [];
    for (let i = 1; i < usersData.length; i++) {
      const row = usersData[i];
      if (!row[COL.EMAIL]) continue;

      let include = false;
      if (!filter || filter.type === 'all') include = true;
      else if (filter.type === 'orgs' && Array.isArray(filter.selections) && filter.selections.length > 0) {
        // orgMatchMode: 'mainOnly' or 'allAffiliations'
        const mode = filter.orgMatchMode === 'mainOnly' ? 'mainOnly' : 'allAffiliations';
        for (const sel of filter.selections) {
          const targetOrg = sel.org;
          const targetDept = sel.dept || '';
          if (mode === 'mainOnly') {
            const org1 = row[COL.ORG_START];
            const dept1 = row[COL.ORG_START + 1];
            if (targetOrg && org1 === targetOrg) {
              if (!targetDept || dept1 === targetDept) { include = true; break; }
            }
          } else {
            // any affiliation match
            for (let k = 0; k < 5; k++) {
              const o = row[COL.ORG_START + k * 3];
              const d = row[COL.ORG_START + k * 3 + 1];
              if (o && o === targetOrg) {
                if (!targetDept || d === targetDept) { include = true; break; }
              }
            }
            if (include) break;
          }
        }
      }

      if (!include) continue;

      const outRow = indices.map(ci => row[ci] === undefined || row[ci] === null ? '' : row[ci]);
      outRows.push(outRow);
    }

    // CSV 生成（Excelの文字化け対策: UTF-8 BOM を先頭に付与）
    const escape = (v) => {
      if (v === null || typeof v === 'undefined') return '';
      const s = String(v);
      if (s.indexOf('"') !== -1) return '"' + s.replace(/"/g, '""') + '"';
      if (s.indexOf(',') !== -1 || s.indexOf('\n') !== -1 || s.indexOf('\r') !== -1) return '"' + s + '"';
      return s;
    };

    const rows = [];
    rows.push(headersOut.map(escape).join(','));
    outRows.forEach(r => rows.push(r.map(escape).join(',')));
    const body = rows.join('\r\n');

    // ファイル名: クライアントが指定すればそれを優先、なければサーバ側の Tokyo 時刻で生成
    const ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');
    const filename = (params && params.filename) ? String(params.filename) : ('list_' + ts + '.csv');

    // Shift_JIS 出力を試みる（Blob によるエンコード）。成功したら base64 バイナリを返す。
    try {
      const blob = Utilities.newBlob('');
      // setDataFromString が利用可能であれば Shift_JIS でセット
      if (typeof blob.setDataFromString === 'function') {
        blob.setDataFromString(body, 'Shift_JIS');
      } else {
        // 互換性フォールバック：直接 newBlob(body) を使う（UTF-8）
        blob.setBytes(Utilities.newBlob(body).getBytes());
      }
      const bytes = blob.getBytes();
      const b64 = Utilities.base64Encode(bytes);
      return { success: true, csvBase64: b64, filename: filename, encoding: 'shift_jis' };
    } catch (e) {
      // フォールバック: UTF-8 with BOM
      const bom = '\uFEFF';
      const csv = bom + body;
      return { success: true, csv: csv, filename: filename, encoding: 'utf-8' };
    }
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
