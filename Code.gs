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
  CAR_OWNER: 24
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
  '車所有'
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

  // 車所有 (Y列)
  const ruleCarOwner = SpreadsheetApp.newDataValidation().requireValueInList(['TRUE', 'FALSE']).setAllowInvalid(true).build();
  sheetUsers.getRange(startRow, 25, numRows, 1).setDataValidation(ruleCarOwner);

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

    const hasToken = (slackToken && slackToken.toString().startsWith('xoxp-'));
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
    if (!lookupJson.ok) return { success: false, message: "Slackアカウントが見つかりません。(Botがワークスペースにいない可能性があります)" };

    const slackUserId = lookupJson.user.id;

    // OTP生成 (6桁数字)
    const otp = Math.floor(100000 + Math.random() * 900000).toString();

    // ScriptPropertiesに一時保存 (有効期限10分想定)
    const otpPayload = JSON.stringify({ code: otp, created: new Date().getTime() });
    PropertiesService.getScriptProperties().setProperty(`OTP_${targetEmail}`, otpPayload);

    // DM送信
    const msgRes = UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": "Bearer " + botToken },
      payload: JSON.stringify({
        channel: slackUserId,
        text: `【${APP_NAME}】認証コード: *${otp}*\nこのコードを画面に入力してください。(有効期限10分)`
      }),
      muteHttpExceptions: true
    });
    if (!JSON.parse(msgRes.getContentText()).ok) throw new Error("Slack DM送信失敗");

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
function getUserProfile(sessionToken) {
  const login = getLoginUser(sessionToken);
  if (login.status !== 'authorized') throw new Error("認証されていません");

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  const row = data.find(r => r[COL.EMAIL] === login.user.email);

  if (!row) throw new Error("データが見つかりません");

  // 編集可能なフィールドを返す (日付は文字列変換)
  return {
    name: row[COL.NAME_JP], // 編集不可
    nameEn: row[COL.NAME_EN],
    email: row[COL.EMAIL],  // 編集不可
    studentId: row[COL.STUDENT_ID],
    grade: row[COL.GRADE],
    field: row[COL.FIELD],
    phone: row[COL.PHONE],
    birthday: row[COL.BIRTHDAY] instanceof Date ? Utilities.formatDate(row[COL.BIRTHDAY], Session.getScriptTimeZone(), 'yyyy/MM/dd') : row[COL.BIRTHDAY],
    almaMater: row[COL.ALMA_MATER],
    carOwner: row[COL.CAR_OWNER] === 'TRUE' || row[COL.CAR_OWNER] === true,
    orgs: [
      { org: row[COL.ORG_START], dept: row[COL.ORG_START+1], role: row[COL.ORG_START+2] },
      { org: row[COL.ORG_START+3], dept: row[COL.ORG_START+4], role: row[COL.ORG_START+5] },
      { org: row[COL.ORG_START+6], dept: row[COL.ORG_START+7], role: row[COL.ORG_START+8] },
      { org: row[COL.ORG_START+9], dept: row[COL.ORG_START+10], role: row[COL.ORG_START+11] },
      { org: row[COL.ORG_START+12], dept: row[COL.ORG_START+13], role: row[COL.ORG_START+14] }
    ]
  };
}

function updateUserProfile(sessionToken, formData) {
  const login = getLoginUser(sessionToken);
  if (login.status !== 'authorized') throw new Error("認証されていません");

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();

  let rowIndex = -1;
  for(let i=1; i<data.length; i++) {
    if (data[i][COL.EMAIL] === login.user.email) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) throw new Error("ユーザーが見つかりません");

  // 更新処理
  sheet.getRange(rowIndex, COL.NAME_EN + 1).setValue(formData.nameEn);
  sheet.getRange(rowIndex, COL.STUDENT_ID + 1).setValue(formData.studentId);
  sheet.getRange(rowIndex, COL.GRADE + 1).setValue(formData.grade);
  sheet.getRange(rowIndex, COL.FIELD + 1).setValue(formData.field);
  sheet.getRange(rowIndex, COL.PHONE + 1).setValue(formData.phone);
  sheet.getRange(rowIndex, COL.BIRTHDAY + 1).setValue(formData.birthday);
  sheet.getRange(rowIndex, COL.ALMA_MATER + 1).setValue(formData.almaMater);
  sheet.getRange(rowIndex, COL.CAR_OWNER + 1).setValue(formData.carOwner ? 'TRUE' : 'FALSE');

  // 所属情報 (5セット)
  if (formData.orgs && Array.isArray(formData.orgs)) {
    for (let k = 0; k < 5; k++) {
      if (k < formData.orgs.length) {
        const o = formData.orgs[k];
        const baseCol = COL.ORG_START + (k * 3) + 1;
        sheet.getRange(rowIndex, baseCol).setValue(o.org || "");
        sheet.getRange(rowIndex, baseCol + 1).setValue(o.dept || "");
        sheet.getRange(rowIndex, baseCol + 2).setValue(o.role || "");
      }
    }
  }

  return { success: true };
}

/* --------------------------------------------------------------------------
 * 5. DM送信 & チャンネル招待
 * -------------------------------------------------------------------------- */
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
      const res = UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", {
        method: "post", contentType: "application/json", headers: { "Authorization": "Bearer " + token },
        payload: JSON.stringify({ channel: uid, text: text }), muteHttpExceptions: true
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
