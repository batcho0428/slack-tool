/* --------------------------------------------------------------------------
 * 設定 & 定数定義
 * -------------------------------------------------------------------------- */
const APP_NAME = '45th NUTFES 実行委員マスタ';
const APP_HEADER_COLOR = '#1a237e'; // 紺色

// SPREADSHEET_IDは関数実行時に取得（定数化によるタイミング問題を回避）
const SHEET_USERS = 'Users';
const SHEET_LOGS = 'Logs';
const SHEET_OPTIONS = 'Options';
const SHEET_FORMS = 'Forms';
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
  RETIRED: 9,
  CONTINUE_NEXT: 10,
  ADMIN: 11,
  CAR_OWNER: 12,
  ORG_START: 13
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
  '退局', '次年度継続', 'Admin', '車所有',
  '所属局1', '所属部門1', '役職1', 'Admin (12)',
  '所属局2', '所属部門2', '役職2',
  '所属局3', '所属部門3', '役職3',
  '所属局4', '所属部門4', '役職4',
  '所属局5', '所属部門5', '役職5'
];

const HEADER_TOKENS = ['Session ID', 'Email', 'Slack Token', 'Created At'];
const HEADER_OPTIONS = ['学年リスト', '分野リスト', '役職リスト', '所属局リスト', '部門マスタ(局)', '部門マスタ(部門)'];
const HEADER_FORMS = ['アンケートシート', 'フォームURL', 'フォームタイトル', '担当局', '担当部門', '収集中', 'スコアの名前', 'スコアの単位'];
const HEADER_LOGS = ['Time', 'Sender', 'Recipient', 'Status', 'Details'];

function getScriptProperty(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

/**
 * Debug helper: probe Forms sheet rows and report whether target spreadsheet and form can be opened.
 * Returns array of { rowIndex, rawA, rawB, spreadsheetId, ssOpen:bool, ssError, formOpen:bool, formError }
 */
function debugProbeForms() {
  const out = [];
  try {
    const ssMain = SpreadsheetApp.openById(getSpreadsheetId());
    const fs = ssMain.getSheetByName(SHEET_FORMS);
    if (!fs) return { success: false, message: 'Forms シートが存在しません' };
      const lr = fs.getLastRow();
      if (lr < 2) return { success: true, probes: [] };
      const cols = Math.max(3, HEADER_FORMS.length);
      const rows = fs.getRange(2,1,Math.max(0, lr-1), cols).getValues();
    for (let i = 0; i < rows.length; i++) {
      const a = String(rows[i][0] || '').trim();
      const b = String(rows[i][1] || '').trim();
      const entry = { rowIndex: i+2, rawA: a, rawB: b, spreadsheetId: null, ssOpen: false, ssError: null, formOpen: false, formError: null };
      try {
        const sid = _extractSpreadsheetId(a) || a || null;
        entry.spreadsheetId = sid;
        if (sid) {
          try { const targetSs = SpreadsheetApp.openById(sid); entry.ssOpen = true; }
          catch (e) { entry.ssOpen = false; entry.ssError = String(e); }
        } else {
          entry.ssError = 'スプレッドシートIDが見つかりません';
        }
        if (b) {
          try {
            let formObj = null;
            try { formObj = FormApp.openByUrl(b); } catch (ee) {
              const fid = _extractFormId(b);
              if (fid) formObj = FormApp.openById(fid);
            }
            if (formObj) entry.formOpen = true; else entry.formError = 'FormAppで開けませんでした';
          } catch (e) { entry.formOpen = false; entry.formError = String(e); }
        }
      } catch (e) {
        entry.ssError = entry.ssError || String(e);
      }
      out.push(entry);
    }
    return { success: true, probes: out };
  } catch (e) {
    return { success: false, message: String(e) };
  }
}

// Stores in Script Properties (JSON maps)
// sessions: sessionId -> { email, created }
// tokensByEmail: email -> { slackToken, created }
const SESSIONS_PROP_KEY = 'SESSIONS_STORE';
const TOKENS_BY_EMAIL_PROP_KEY = 'TOKENS_BY_EMAIL_STORE';
const TOKENS_PROP_MAX = 200000; // safe threshold (characters). 古いものから削除して収める

function _loadStore(key) {
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  if (!raw) return {};
  try { return JSON.parse(raw); } catch (e) { return {}; }
}

function _saveStore(key, obj) {
  let serialized = JSON.stringify(obj);
  if (serialized.length <= TOKENS_PROP_MAX) {
    PropertiesService.getScriptProperties().setProperty(key, serialized);
    return;
  }
  // remove oldest entries until it fits
  const entries = Object.keys(obj).map(k => ({ k, created: obj[k] && obj[k].created ? obj[k].created : 0 }));
  entries.sort((a, b) => a.created - b.created);
  for (let i = 0; i < entries.length && serialized.length > TOKENS_PROP_MAX; i++) {
    delete obj[entries[i].k];
    serialized = JSON.stringify(obj);
  }
  PropertiesService.getScriptProperties().setProperty(key, serialized);
}

function getSpreadsheetId() {
  const id = getScriptProperty('SPREADSHEET_ID');
  if (!id) throw new Error('SPREADSHEET_IDがScript Propertiesに設定されていません');
  return id;
}

/* --------------------------------------------------------------------------
 * 0. 初期セットアップ (マイグレーション & トリガー設定)
 * -------------------------------------------------------------------------- */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());

  // 1. Optionsシート
  let sheetOptions = ss.getSheetByName(SHEET_OPTIONS);
  if (!sheetOptions) sheetOptions = ss.insertSheet(SHEET_OPTIONS);
  if (sheetOptions.getLastRow() === 0) sheetOptions.getRange(1, 1, 1, HEADER_OPTIONS.length).setValues([HEADER_OPTIONS]);

  // 1b. Formsシート (アンケート情報を専用シートに移行)
  let sheetForms = ss.getSheetByName(SHEET_FORMS);
  if (!sheetForms) sheetForms = ss.insertSheet(SHEET_FORMS);
  // Ensure header row contains our expected HEADER_FORMS columns. Preserve existing non-empty headers when possible.
  try {
    const existingCols = Math.max(1, sheetForms.getLastColumn());
    const cols = Math.max(existingCols, HEADER_FORMS.length);
    const cur = sheetForms.getRange(1, 1, 1, cols).getValues()[0] || [];
    const newHeaders = [];
    for (let i = 0; i < HEADER_FORMS.length; i++) {
      // prefer existing header if non-empty, else use standard
      newHeaders[i] = (cur[i] && String(cur[i]).trim()) ? String(cur[i]) : HEADER_FORMS[i];
    }
    sheetForms.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);

    // Ensure '収集中' column is formatted as checkboxes and normalize existing values
    const collectingIdx = newHeaders.findIndex(h => String(h || '').trim() === '収集中');
    if (collectingIdx >= 0) {
      try {
        const startRow = 2;
        const lastRow = Math.max(sheetForms.getLastRow(), startRow);
        const numRows = Math.max(1, lastRow - 1);
        // convert existing TRUE/FALSE strings to booleans
        try {
          const range = sheetForms.getRange(startRow, collectingIdx + 1, numRows, 1);
          const vals = range.getValues();
          const norm = vals.map(r => {
            const v = r[0];
            if (v === true || v === 'TRUE' || String(v).toLowerCase() === 'true') return [true];
            if (v === false || v === 'FALSE' || String(v).toLowerCase() === 'false') return [false];
            return [false];
          });
          range.setValues(norm);
        } catch (e) {
          // ignore conversion errors
        }
        // insert checkbox formatting
        try { sheetForms.getRange(startRow, collectingIdx + 1, numRows, 1).insertCheckboxes(); }
        catch (e) {
          // fallback: data validation to TRUE/FALSE
          try { sheetForms.getRange(startRow, collectingIdx + 1, numRows, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['TRUE','FALSE']).setAllowInvalid(true).build()); } catch (ee) {}
        }
      } catch (e) {
        // ignore per-sheet failures
      }
    }
  } catch (e) {
    // fallback: set headers directly
    try { sheetForms.getRange(1, 1, 1, HEADER_FORMS.length).setValues([HEADER_FORMS]); } catch (ee) {}
  }

  // migrate existing Options G:I -> Forms if present (pad to new column count)
  try {
    const lastOptRow = sheetOptions.getLastRow();
    if (lastOptRow >= 2) {
      const surveyRows = sheetOptions.getRange(2, 7, Math.max(0, lastOptRow - 1), 3).getValues();
      const exist = {};
      const lastFormRow = sheetForms.getLastRow();
      if (lastFormRow >= 2) {
        const cur = sheetForms.getRange(2,1,Math.max(0,lastFormRow-1),1).getValues().map(r=>String(r[0]||''));
        cur.forEach(v=>{ if (v) exist[v.trim()] = true; });
      }
      const toAppend = [];
      const padLen = HEADER_FORMS.length;
      surveyRows.forEach(r => {
        const sName = String(r[0] || '').trim();
        const url = String(r[1] || '').trim();
        const title = String(r[2] || '').trim();
        if (sName && !exist[sName]) {
          const row = new Array(padLen).fill('');
          row[0] = sName; row[1] = url; row[2] = title;
          toAppend.push(row);
          exist[sName]=true;
        }
      });
      if (toAppend.length>0) sheetForms.getRange(sheetForms.getLastRow()+1, 1, toAppend.length, padLen).setValues(toAppend);
    }
  } catch (e) { /* ignore migration errors */ }

  // 2. Usersシート
  let sheetUsers = ss.getSheetByName(SHEET_USERS);
  if (!sheetUsers) sheetUsers = ss.insertSheet(SHEET_USERS);
  sheetUsers.getRange(1, 1, 1, HEADER_USERS.length).setValues([HEADER_USERS]);

  // 3. Logsシート
  let sheetLogs = ss.getSheetByName(SHEET_LOGS);
  if (!sheetLogs) sheetLogs = ss.insertSheet(SHEET_LOGS);
  if (sheetLogs.getLastRow() === 0) sheetLogs.getRange(1, 1, 1, HEADER_LOGS.length).setValues([HEADER_LOGS]);

  // 4. Collections 系シート（集金）
  try { ensureCollectionsSheets(); } catch (e) { /* ignore if cannot create */ }

  // Tokens は PropertiesService に移行したためシートは作成しない

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
    const baseCol = COL.ORG_START + 1 + (k * 3); // 1-based
    sheetUsers.getRange(startRow, baseCol, numRows, 1).setDataValidation(ruleOrg);
    sheetUsers.getRange(startRow, baseCol + 1, numRows, 1).setDataValidation(ruleDept);
    sheetUsers.getRange(startRow, baseCol + 2, numRows, 1).setDataValidation(ruleRole);
  }

  // 車所有 (Y列) と Admin列をチェックボックスに変更（フォールバックあり）
  // Ensure boolean flags are checkboxes: 在籍, 次年度継続, 車所有, Admin
  const boolColsToCheckbox = [COL.RETIRED, COL.CONTINUE_NEXT, COL.CAR_OWNER, COL.ADMIN];
  boolColsToCheckbox.forEach(colIdx => {
    try {
      sheetUsers.getRange(startRow, colIdx + 1, numRows, 1).insertCheckboxes();
    } catch (e) {
      // fallback: set data validation to TRUE/FALSE list
      try {
        const rule = SpreadsheetApp.newDataValidation().requireValueInList(['TRUE', 'FALSE']).setAllowInvalid(true).build();
        sheetUsers.getRange(startRow, colIdx + 1, numRows, 1).setDataValidation(rule);
      } catch (ee) {
        // ignore
      }
    }
  });

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
    const colOrgIndex = COL.ORG_START + (k * 3) + 1; // 1-based: 所属局のカラムインデックス
    const colDeptIndex = COL.ORG_START + (k * 3) + 2; // 1-based: 所属部門のカラムインデックス
    const colOrgLet = getColLetter(colOrgIndex);
    const colDeptLet = getColLetter(colDeptIndex);
    const range = sheetUsers.getRange(`${colDeptLet}2:${colDeptLet}`);
    const formula = `=AND(${colDeptLet}2<>"", COUNTIFS(INDIRECT("${SHEET_OPTIONS}!\\$E:\\$E"), ${colOrgLet}2, INDIRECT("${SHEET_OPTIONS}!\\$F:\\$F"), ${colDeptLet}2)=0)`;
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula).setBackground("#FFFF00").setRanges([range]).build());
  }
  sheetUsers.setConditionalFormatRules(rules);

  installTriggers();
  // マイグレーション: 既存の 'TRUE'/'FALSE' 文字列を boolean に変換
  try {
    const lastRowUsers = sheetUsers.getLastRow();
    if (lastRowUsers >= startRow) {
      // convert 'TRUE'/'FALSE' strings to booleans for checkbox columns
      const boolCols = [COL.CAR_OWNER, COL.ADMIN, COL.RETIRED, COL.CONTINUE_NEXT];
      boolCols.forEach(colIdx => {
        try {
          const range = sheetUsers.getRange(startRow, colIdx + 1, lastRowUsers - startRow + 1, 1);
          const vals = range.getValues().map(r => { const v = r[0]; if (v === true || v === 'TRUE') return [true]; if (v === false || v === 'FALSE') return [false]; return [false]; });
          range.setValues(vals);
        } catch (e) {
          // ignore per-column failures
        }
      });
    }
  } catch (e) {
    console.warn('Checkbox migration failed:', e.toString());
  }
  console.log("セットアップ完了");
}

function installTriggers() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
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

    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const usersSheet = ss.getSheetByName(SHEET_USERS);
    if (!usersSheet) return { status: 'error', message: "DB構成エラー" };

    const sessions = _loadStore(SESSIONS_PROP_KEY);
    const entry = sessions[sessionToken];
    if (!entry) return { status: 'guest' };

    const now = Date.now();
    const created = entry.created || 0;
    const diffDays = (now - created) / (1000 * 60 * 60 * 24);
    if (diffDays > SESSION_DURATION_DAYS) {
      delete sessions[sessionToken];
      _saveStore(SESSIONS_PROP_KEY, sessions);
      return { status: 'guest', message: 'セッション有効期限切れ' };
    }

    const userEmail = entry.email;
    const tokensByEmail = _loadStore(TOKENS_BY_EMAIL_PROP_KEY);
    const tokenEntry = tokensByEmail[userEmail] || {};
    const slackToken = tokenEntry.slackToken || '';

    // ユーザー情報取得
    const userData = usersSheet.getDataRange().getValues();
    const userRow = userData.find(r => String(r[COL.EMAIL] || '').trim().toLowerCase() === String(userEmail).trim().toLowerCase());
    if (!userRow) return { status: 'error', message: "ユーザー情報が見つかりません" };

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

function _requireAdmin(sessionToken) {
  const login = getLoginUser(sessionToken);
  if (!login || login.status !== 'authorized') throw new Error('認証されていません');
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const usersSheet = ss.getSheetByName(SHEET_USERS);
  if (!usersSheet) throw new Error('Users シートが見つかりません');
  const usersData = usersSheet.getDataRange().getValues();
  const row = usersData.find(r => String(r[COL.EMAIL] || '').trim().toLowerCase() === String(login.user.email || '').trim().toLowerCase());
  const isAdmin = row && (row[COL.ADMIN] === 'TRUE' || row[COL.ADMIN] === true);
  if (!isAdmin) throw new Error('管理者権限が必要です');
  return login;
}

// 1-A. OTPリクエスト (BotからDM送信)
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
  const newSessionToken = Utilities.getUuid();
  const nowTs = Date.now();
  const sessions = _loadStore(SESSIONS_PROP_KEY);
  sessions[newSessionToken] = { email: targetEmail, created: nowTs };
  _saveStore(SESSIONS_PROP_KEY, sessions);
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

    const userEmailRaw = infoJson.user.profile.email;
    const userEmail = String(userEmailRaw || '').trim().toLowerCase();
    const ss = SpreadsheetApp.openById(getSpreadsheetId());

    // ユーザー登録チェック
    const usersSheet = ss.getSheetByName(SHEET_USERS);
    const userData = usersSheet.getDataRange().getValues();
    const userExists = userData.some((r, i) => i > 0 && String(r[COL.EMAIL] || '').trim().toLowerCase() === userEmail);

    if (!userExists) return HtmlService.createHtmlOutput(`<h2 style="color:red; text-align:center;">未登録ユーザー (${userEmail})</h2>`);

    // セッションとトークンを別々に保存
    const newSessionToken = Utilities.getUuid();
    const now = Date.now();
    const sessions = _loadStore(SESSIONS_PROP_KEY);
    sessions[newSessionToken] = { email: userEmail, created: now };
    _saveStore(SESSIONS_PROP_KEY, sessions);

    const tokensByEmail = _loadStore(TOKENS_BY_EMAIL_PROP_KEY);
    tokensByEmail[userEmail] = { slackToken: userSlackToken, created: now };
    _saveStore(TOKENS_BY_EMAIL_PROP_KEY, tokensByEmail);

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
          <p>認証が完了しました。ツールへ戻るをクリックしてください。</p>
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

  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const sheet = ss.getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();

  // 管理者判定はログインユーザーの Users 行の Z 列 (COL.ADMIN)
  const loginRow = data.find(r => String(r[COL.EMAIL] || '').trim().toLowerCase() === String(login.user.email).trim().toLowerCase());
  const isAdmin = loginRow && (loginRow[COL.ADMIN] === 'TRUE' || loginRow[COL.ADMIN] === true);

  const emailToFetch = (targetEmail && isAdmin) ? targetEmail : login.user.email;
  const row = data.find(r => String(r[COL.EMAIL] || '').trim().toLowerCase() === String(emailToFetch).trim().toLowerCase());
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
    retired: row[COL.RETIRED] === 'TRUE' || row[COL.RETIRED] === true,
    continueNext: row[COL.CONTINUE_NEXT] === 'TRUE' || row[COL.CONTINUE_NEXT] === true,
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

  const ss = SpreadsheetApp.openById(getSpreadsheetId());
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
  // 在籍(退局)フラグ: 管理者は他ユーザーの退局フラグを変更可能
  if (typeof formData.retired !== 'undefined') {
    const currentValRet = sheet.getRange(rowIndex, COL.RETIRED + 1).getValue();
    const currentBoolRet = (currentValRet === true || currentValRet === 'TRUE');
    const requestedRet = !!formData.retired;

    // 他ユーザーの退局変更は管理者のみ
    if (String(emailToSave).trim().toLowerCase() !== String(login.user.email).trim().toLowerCase() && !isAdmin) {
      throw new Error('退局フラグを変更する権限がありません');
    }

    // 非管理者は在籍(false) -> 退局(true) のみ許可（退局->在籍は不可）
    if (!isAdmin) {
      if (currentBoolRet === true && requestedRet === false) throw new Error('退局から復帰する権限はありません');
    }

    sheet.getRange(rowIndex, COL.RETIRED + 1).setValue(requestedRet ? true : false);
  }

  // 次年度継続スイッチの制御: 常にデータは保存するが、UIの操作は制限される可能性がある
  if (typeof formData.continueNext !== 'undefined') {
    const currentVal = sheet.getRange(rowIndex, COL.CONTINUE_NEXT + 1).getValue();
    const currentBool = (currentVal === true || currentVal === 'TRUE');
    const requested = !!formData.continueNext;
    if (!isAdmin) {
      // 非管理者は OFF -> ON のみ許可（ON->OFF は不可）
      if (currentBool === true && requested === false) throw new Error('次年度継続を取り消す権限はありません');
    }
    sheet.getRange(rowIndex, COL.CONTINUE_NEXT + 1).setValue(requested ? true : false);
  }

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

  // 次年度継続スイッチの制御: 常にデータは保存するが、UIの操作は制限される可能性がある
  if (typeof formData.continueNext !== 'undefined') {
    const currentVal = sheet.getRange(rowIndex, COL.CONTINUE_NEXT + 1).getValue();
    const currentBool = (currentVal === true || currentVal === 'TRUE');
    const requested = !!formData.continueNext;
    if (!isAdmin) {
      // 非管理者は OFF -> ON のみ許可（ON->OFF は不可）
      if (currentBool === true && requested === false) throw new Error('次年度継続を取り消す権限はありません');
    }
    sheet.getRange(rowIndex, COL.CONTINUE_NEXT + 1).setValue(requested ? true : false);
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

  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const usersSheet = ss.getSheetByName(SHEET_USERS);
  if (!usersSheet) throw new Error('Users シートが見つかりません');

  // 管理者判定
  const allUsers = usersSheet.getDataRange().getValues();
  const loginRow = allUsers.find(r => String(r[COL.EMAIL] || '').trim().toLowerCase() === String(login.user.email).trim().toLowerCase());
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

  // 退局 / 次年度継続 (退局 default: false => 在籍)
  row[COL.RETIRED] = (typeof userObj.retired !== 'undefined') ? !!userObj.retired : false;
  row[COL.CONTINUE_NEXT] = (typeof userObj.continueNext !== 'undefined') ? !!userObj.continueNext : false;

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
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
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
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
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
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
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
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const sheet = ss.getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  const results = [];
  const q = criteria.query ? criteria.query.toLowerCase() : "";
  const filterGrade = criteria.grade || "";
  const filterField = criteria.field || "";
  const filterOrg = criteria.org || "";
  const filterDept = criteria.dept || "";
  const filterRole = criteria.role || "";
  const filterStatus = criteria.status || "active"; // 'active', 'retired', 'all'

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const nameJp = row[COL.NAME_JP];
    const nameEn = row[COL.NAME_EN];
    const studentId = row[COL.STUDENT_ID] || "";
    const grade = row[COL.GRADE];
    const field = row[COL.FIELD];
    const email = row[COL.EMAIL];
    const almaMater = row[COL.ALMA_MATER] || "";
    const retired = row[COL.RETIRED] === true || row[COL.RETIRED] === 'TRUE';
    const searchString = `${nameJp} ${nameEn} ${email} ${almaMater} ${studentId}`.toLowerCase();

    if (q && !searchString.includes(q)) continue;

    // 在籍フィルタ処理
    if (filterStatus === 'active' && retired) continue;
    if (filterStatus === 'retired' && !retired) continue;
    // filterStatus === 'all' の場合は全て表示

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

  // 所属局列を検出（1-based）
  const startCol = COL.ORG_START + 1;
  const lastCol = COL.ORG_START + (5 * 3); // 1-based last column for roles
  if (col < startCol || col > lastCol) return;
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
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) throw new Error('Users シートが見つかりません');

  const lastCol = Math.max(sheet.getLastColumn(), COL.ADMIN + 1);
  const headerRange = sheet.getRange(1, 1, 1, lastCol);
  const headers = headerRange.getValues()[0];

  // ヘッダが短ければ拡張
  const lastRequiredCol = Math.max(COL.RETIRED, COL.CONTINUE_NEXT, COL.ADMIN, COL.CAR_OWNER) + 1; // 0-based -> +1
  if (headers.length < lastRequiredCol) {
    const newHeaders = headers.slice();
    for (let i = headers.length; i < lastRequiredCol; i++) newHeaders[i] = '';
    // set known header names
    newHeaders[COL.RETIRED] = '退局';
    newHeaders[COL.CONTINUE_NEXT] = '次年度継続';
    newHeaders[COL.ADMIN] = 'Admin';
    newHeaders[COL.CAR_OWNER] = '車所有';
    sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  } else {
    // ensure headers exist for these columns
    sheet.getRange(1, COL.RETIRED + 1).setValue(headers[COL.RETIRED] || '退局');
    sheet.getRange(1, COL.CONTINUE_NEXT + 1).setValue(headers[COL.CONTINUE_NEXT] || '次年度継続');
    sheet.getRange(1, COL.ADMIN + 1).setValue(headers[COL.ADMIN] || 'Admin');
    sheet.getRange(1, COL.CAR_OWNER + 1).setValue(headers[COL.CAR_OWNER] || '車所有');
  }

  // データ行の Admin 列が空の行は FALSE に設定
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, message: 'ヘッダのみです' };

  // Ensure data rows have boolean values and checkbox formatting for boolean columns
  const boolCols = [COL.RETIRED, COL.CONTINUE_NEXT, COL.ADMIN, COL.CAR_OWNER];
  let updated = 0;
  boolCols.forEach(colIdx => {
    try {
      const colRange = sheet.getRange(2, colIdx + 1, lastRow - 1, 1);
      const colVals = colRange.getValues();
      for (let i = 0; i < colVals.length; i++) {
        const v = colVals[i][0];
        if (v === '' || v === null || typeof v === 'undefined') { colVals[i][0] = false; updated++; }
        else if (v === 'TRUE') colVals[i][0] = true;
        else if (v === 'FALSE') colVals[i][0] = false;
      }
      if (updated > 0) colRange.setValues(colVals);
    } catch (e) {
      // ignore per-column failures
    }
    // ensure checkbox formatting
    try { sheet.getRange(2, colIdx + 1, lastRow - 1, 1).insertCheckboxes(); } catch (e) { /* ignore */ }
  });
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

function isContinueSwitchEnabled() {
  try {
    return getScriptProperty('NEXT_YEAR_CONTINUE_ENABLED') === 'true';
  } catch (e) {
    return false;
  }
}

function createRosterSpreadsheet(sessionToken, selectedFields, folderId, filename) {
  // sessionToken: to validate user and permissions
  try {
    const login = getLoginUser(sessionToken);
    if (login.status !== 'authorized') throw new Error('認証されていません');

    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const usersSheet = ss.getSheetByName(SHEET_USERS);
    if (!usersSheet) throw new Error('Users シートが見つかりません');

    // 管理者判定
    const usersData = usersSheet.getDataRange().getValues();
    const loginRow = usersData.find(r => String(r[COL.EMAIL] || '').trim().toLowerCase() === String(login.user.email).trim().toLowerCase());
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
      else if (f === '学籍番号') pushIf('学籍番号', COL.STUDENT_ID);
      else if (f === '学年') pushIf('学年', COL.GRADE);
      else if (f === '分野') pushIf('分野', COL.FIELD);
      else if (f === 'メールアドレス') pushIf('メールアドレス', COL.EMAIL);
      else if (f === '電話番号') pushIf('電話番号', COL.PHONE);
      else if (f === '生年月日') pushIf('生年月日', COL.BIRTHDAY);
      else if (f === '出身校') pushIf('出身校', COL.ALMA_MATER);
      else if (f === '退局' || f === '在籍') pushIf('退局', COL.RETIRED);
      else if (f === '次年度継続') pushIf('次年度継続', COL.CONTINUE_NEXT);
      else if (f === '所属局1') pushIf('所属局1', COL.ORG_START + 0 * 3);
      else if (f === '所属部門1') pushIf('所属部門1', COL.ORG_START + 0 * 3 + 1);
      else if (f === '役職1') pushIf('役職1', COL.ORG_START + 0 * 3 + 2);
      else if (f === '所属局2') pushIf('所属局2', COL.ORG_START + 1 * 3);
      else if (f === '所属部門2') pushIf('所属部門2', COL.ORG_START + 1 * 3 + 1);
      else if (f === '役職2') pushIf('役職2', COL.ORG_START + 1 * 3 + 2);
      else if (f === '所属局3') pushIf('所属局3', COL.ORG_START + 2 * 3);
      else if (f === '所属部門3') pushIf('所属部門3', COL.ORG_START + 2 * 3 + 1);
      else if (f === '役職3') pushIf('役職3', COL.ORG_START + 2 * 3 + 2);
      else if (f === '所属局4') pushIf('所属局4', COL.ORG_START + 3 * 3);
      else if (f === '所属部門4') pushIf('所属部門4', COL.ORG_START + 3 * 3 + 1);
      else if (f === '役職4') pushIf('役職4', COL.ORG_START + 3 * 3 + 2);
      else if (f === '所属局5') pushIf('所属局5', COL.ORG_START + 4 * 3);
      else if (f === '所属部門5') pushIf('所属部門5', COL.ORG_START + 4 * 3 + 1);
      else if (f === '役職5') pushIf('役職5', COL.ORG_START + 4 * 3 + 2);
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
      const outRow = indices.map(ci => {
        let v = row[ci];
        if (v === undefined || v === null) return '';
        // Format birthday as YYYY/MM/DD for CSV (Windows Excel friendly)
        if (ci === COL.BIRTHDAY) {
          try {
            if (Object.prototype.toString.call(v) === '[object Date]') {
              return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy/MM/dd');
            }
            // if it's a string like 2004-04-28 or 2004/04/28
            const s = String(v).trim();
            if (s.indexOf('-') !== -1) return s.replace(/-/g, '/');
            return s;
          } catch (e) {
            return String(v);
          }
        }
        return v;
      });
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

    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const usersSheet = ss.getSheetByName(SHEET_USERS);
    if (!usersSheet) throw new Error('Users シートが見つかりません');

    // 管理者判定
    const usersData = usersSheet.getDataRange().getValues();
    const loginRow = usersData.find(r => String(r[COL.EMAIL] || '').trim().toLowerCase() === String(login.user.email).trim().toLowerCase());
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
      else if (f === '退局' || f === '在籍') add(COL.RETIRED);
      else if (f === '次年度継続') add(COL.CONTINUE_NEXT);
      else if (f === '所属局1') { add(COL.ORG_START + 0 * 3); }
      else if (f === '所属部門1') { add(COL.ORG_START + 0 * 3 + 1); }
      else if (f === '役職1') { add(COL.ORG_START + 0 * 3 + 2); }
      else if (f === '所属局2') { add(COL.ORG_START + 1 * 3); }
      else if (f === '所属部門2') { add(COL.ORG_START + 1 * 3 + 1); }
      else if (f === '役職2') { add(COL.ORG_START + 1 * 3 + 2); }
      else if (f === '所属局3') { add(COL.ORG_START + 2 * 3); }
      else if (f === '所属部門3') { add(COL.ORG_START + 2 * 3 + 1); }
      else if (f === '役職3') { add(COL.ORG_START + 2 * 3 + 2); }
      else if (f === '所属局4') { add(COL.ORG_START + 3 * 3); }
      else if (f === '所属部門4') { add(COL.ORG_START + 3 * 3 + 1); }
      else if (f === '役職4') { add(COL.ORG_START + 3 * 3 + 2); }
      else if (f === '所属局5') { add(COL.ORG_START + 4 * 3); }
      else if (f === '所属部門5') { add(COL.ORG_START + 4 * 3 + 1); }
      else if (f === '役職5') { add(COL.ORG_START + 4 * 3 + 2); }
      else if (f === '車所有') add(COL.CAR_OWNER);
      else if (f === 'Admin') add(COL.ADMIN);
    }

    const indices = Array.from(allowedIdxSet).sort((a,b)=>a-b);
    const headersOut = indices.map(i => HEADER_USERS[i] || '');

    if (indices.length === 0) throw new Error('出力項目が選択されていません');

    // フィルタ処理 (status: 'active'|'retired'|'all'), grade, field
    const statusFilter = (filter && filter.status) ? filter.status : 'active';
    if ((statusFilter === 'retired' || statusFilter === 'all') && !isAdmin) {
      throw new Error('退局者または全員の出力は管理者のみ可能です');
    }
    // Build filtered list of data rows first, then sort according to Options ordering
    const filteredRows = [];
    for (let i = 1; i < usersData.length; i++) {
      const row = usersData[i];
      if (!row[COL.EMAIL]) continue;
      if (filter && filter.grade) { if ((row[COL.GRADE] || '') !== String(filter.grade)) continue; }
      if (filter && filter.field) { if ((row[COL.FIELD] || '') !== String(filter.field)) continue; }

      let include = false;
      if (statusFilter === 'all') include = true;
      else if (statusFilter === 'active') { if (!(row[COL.RETIRED] === true || row[COL.RETIRED] === 'TRUE')) include = true; }
      else if (statusFilter === 'retired') { if (row[COL.RETIRED] === true || row[COL.RETIRED] === 'TRUE') include = true; }
      if (!include) continue;

      if (!filter || filter.type === 'all') {
        // ok
      } else if (filter.type === 'orgs' && Array.isArray(filter.selections) && filter.selections.length > 0) {
        const mode = filter.orgMatchMode === 'mainOnly' ? 'mainOnly' : 'allAffiliations';
        let matched = false;
        for (const sel of filter.selections) {
          const targetOrg = sel.org;
          const targetDept = sel.dept || '';
          if (mode === 'mainOnly') {
            const org1 = row[COL.ORG_START];
            const dept1 = row[COL.ORG_START + 1];
            if (targetOrg && org1 === targetOrg) {
              if (!targetDept || dept1 === targetDept) { matched = true; break; }
            }
          } else {
            for (let k = 0; k < 5; k++) {
              const o = row[COL.ORG_START + k * 3];
              const d = row[COL.ORG_START + k * 3 + 1];
              if (o && o === targetOrg) {
                if (!targetDept || d === targetDept) { matched = true; break; }
              }
            }
            if (matched) break;
          }
        }
        if (!matched) continue;
      }
      filteredRows.push(row);
    }

    // Load Options ordering for sorting
    let gradeOrder = [];
    let fieldOrder = [];
    let orgOrder = [];
    const deptMap = {}; // { orgName: [dept1, dept2...] }
    try {
      const optSheet = ss.getSheetByName(SHEET_OPTIONS);
      if (optSheet) {
        const odata = optSheet.getDataRange().getValues();
        for (let i = 1; i < odata.length; i++) {
          const r = odata[i] || [];
          const g = String(r[0] || '').trim(); if (g && gradeOrder.indexOf(g) === -1) gradeOrder.push(g);
          const f = String(r[1] || '').trim(); if (f && fieldOrder.indexOf(f) === -1) fieldOrder.push(f);
          const org = String(r[3] || '').trim(); if (org && orgOrder.indexOf(org) === -1) orgOrder.push(org);
          const deptOrg = String(r[4] || '').trim(); const deptName = String(r[5] || '').trim();
          if (deptOrg && deptName) {
            if (!deptMap[deptOrg]) deptMap[deptOrg] = [];
            if (deptMap[deptOrg].indexOf(deptName) === -1) deptMap[deptOrg].push(deptName);
          }
        }
      }
    } catch (e) { /* ignore */ }

    const idxIn = (arr, v) => { if (!arr || !arr.length) return -1; if (!v) return arr.length + 1; const i = arr.indexOf(String(v)); return i === -1 ? arr.length : i; };

    filteredRows.sort((A, B) => {
      // org1
      const aOrg = String(A[COL.ORG_START] || '');
      const bOrg = String(B[COL.ORG_START] || '');
      const ai = idxIn(orgOrder, aOrg);
      const bi = idxIn(orgOrder, bOrg);
      if (ai !== bi) return ai - bi;
      // dept1
      const aDept = String(A[COL.ORG_START + 1] || '');
      const bDept = String(B[COL.ORG_START + 1] || '');
      const deptList = deptMap[aOrg] || [];
      const adi = deptList.indexOf(aDept); const bdi = deptList.indexOf(bDept);
      if (adi !== bdi) return (adi === -1 ? 1 : adi) - (bdi === -1 ? 1 : bdi);
      // grade
      const aGrade = String(A[COL.GRADE] || '');
      const bGrade = String(B[COL.GRADE] || '');
      const agi = idxIn(gradeOrder, aGrade);
      const bgi = idxIn(gradeOrder, bGrade);
      if (agi !== bgi) return agi - bgi;
      // field
      const aField = String(A[COL.FIELD] || '');
      const bField = String(B[COL.FIELD] || '');
      const afi = idxIn(fieldOrder, aField);
      const bfi = idxIn(fieldOrder, bField);
      if (afi !== bfi) return afi - bfi;
      // fallback: name jp
      const an = String(A[COL.NAME_JP] || '').toLowerCase();
      const bn = String(B[COL.NAME_JP] || '').toLowerCase();
      if (an < bn) return -1; if (an > bn) return 1; return 0;
    });

    const outRows = filteredRows.map(row => indices.map(ci => row[ci] === undefined || row[ci] === null ? '' : row[ci]));

    // CSV 生成（Excelの文字化け対策: UTF-8 BOM を先頭に付与）
    const escape = (v) => {
      if (v === null || typeof v === 'undefined') return '';
      const s = String(v);
      if (s.indexOf('"') !== -1) return '"' + s.replace(/"/g, '""') + '"';
      if (s.indexOf(',') !== -1 || s.indexOf('\n') !== -1 || s.indexOf('\r') !== -1) return '"' + s + '"';
      return s;
    };

    // 正規化: 生年月日を必ず YYYY/MM/DD に変換し、Date オブジェクトはフォーマットする
    const formattedOutRows = outRows.map(r => r.map((cell, j) => {
      const origCol = indices[j];
      if (origCol === COL.BIRTHDAY) {
        if (!cell && cell !== 0) return '';
        if (Object.prototype.toString.call(cell) === '[object Date]' || cell instanceof Date) {
          try { return Utilities.formatDate(new Date(cell), 'Asia/Tokyo', 'yyyy/MM/dd'); } catch (e) { return String(cell); }
        }
        // try parseable string
        const s = String(cell).trim();
        const d = new Date(s);
        if (!isNaN(d.getTime())) return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
        // common separators
        return s.replace(/-/g, '/').replace(/\./g, '/');
      }
      if (cell === true) return 'TRUE';
      if (cell === false) return 'FALSE';
      return cell === undefined || cell === null ? '' : cell;
    }));

    // optionally append selected survey responses to the right of roster (join by email)
    let surveyAppendHeaders = [];
    let surveyByEmail = {};
    if (params && params.surveyRef) {
      try {
        const sid = _extractSpreadsheetId(params.surveyRef) || String(params.surveyRef);
        if (sid) {
          try {
            const targetSs = SpreadsheetApp.openById(sid);
            const surveyTitle = targetSs.getName();
            const sSh = targetSs.getSheets()[0];
            const sLastCol = Math.max(1, sSh.getLastColumn());
            const sHeaderRow = _detectHeaderRow(sSh, sLastCol);
            const surveyHeadersRaw = sSh.getRange(sHeaderRow, 1, 1, sLastCol).getValues()[0] || [];
            const emailIdx = _findHeaderIndex(surveyHeadersRaw, ['メール', 'メールアドレス', '^email$','^e-mail$']);
            const timeIdx = _findHeaderIndex(surveyHeadersRaw, ['タイムスタンプ','Timestamp','回答日時','回答日','日時']);
            const appendIdxs = [];
            for (let i = 0; i < surveyHeadersRaw.length; i++) {
              if (i === emailIdx || i === timeIdx) continue;
              appendIdxs.push(i);
            }
            if (emailIdx >= 0 && appendIdxs.length > 0) {
              surveyAppendHeaders = appendIdxs.map(i => {
                const h = String(surveyHeadersRaw[i] || '').trim() || ('col' + (i + 1));
                return (surveyTitle ? (surveyTitle + ' - ' + h) : h);
              });

              const sDataCount = Math.max(0, sSh.getLastRow() - sHeaderRow);
              const sData = (sDataCount > 0) ? sSh.getRange(sHeaderRow + 1, 1, sDataCount, sLastCol).getValues() : [];
              sData.forEach(r => {
                const email = String(r[emailIdx] || '').trim().toLowerCase();
                if (!email) return;
                let ts = 0;
                if (timeIdx >= 0) {
                  const t = r[timeIdx];
                  if (t instanceof Date) ts = t.getTime();
                  else {
                    const tt = Date.parse(String(t || ''));
                    if (!isNaN(tt)) ts = tt;
                  }
                }
                const values = appendIdxs.map(i => _safeValueForClient(r[i]));
                if (!surveyByEmail[email] || ts >= surveyByEmail[email].ts) {
                  surveyByEmail[email] = { ts: ts, values: values };
                }
              });
            }
          } catch (e) {
            // ignore survey append errors
          }
        }
      } catch (e) { /* ignore */ }
    }

    const finalHeaders = headersOut.concat(surveyAppendHeaders);
    const finalRows = formattedOutRows.map((r, i) => {
      if (!surveyAppendHeaders.length) return r;
      const email = String(filteredRows[i][COL.EMAIL] || '').trim().toLowerCase();
      const extra = surveyByEmail[email] ? surveyByEmail[email].values : new Array(surveyAppendHeaders.length).fill('');
      return r.concat(extra);
    });

    const rows = [];
    rows.push(finalHeaders.map(escape).join(','));
    finalRows.forEach(r => rows.push(r.map(escape).join(',')));

    const body = rows.join('\r\n');

    // Prepare matrix for spreadsheet export (ensure consistent column count)
    const matrix = [finalHeaders].concat(finalRows.map(r => r.map(c => c === undefined || c === null ? '' : c)));

    // ファイル名: クライアントが指定すればそれを優先、なければサーバ側の Tokyo 時刻で生成
    const ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');
    const filename = (params && params.filename) ? String(params.filename) : ('list_' + ts + '.csv');

    // Create Excel (.xlsx) by writing matrix to a temporary Spreadsheet and exporting
    try {
      const tempName = 'tmp_export_' + ts;
      const tempSs = SpreadsheetApp.create(tempName);
      const sh = tempSs.getSheets()[0];
      // ensure dimensions
      const numRows = matrix.length;
      const numCols = finalHeaders.length || 1;
      // pad rows to numCols
      const norm = matrix.map(r => {
        const nr = r.slice();
        while (nr.length < numCols) nr.push('');
        return nr;
      });
      sh.getRange(1,1, norm.length, numCols).setValues(norm);
      // export as xlsx
      const file = DriveApp.getFileById(tempSs.getId());
      const xlsxBlob = file.getBlob().getAs('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      const excelB64 = Utilities.base64Encode(xlsxBlob.getBytes());
      // remove temp spreadsheet
      try { DriveApp.getFileById(tempSs.getId()).setTrashed(true); } catch(e) {}

      // also prepare CSV (UTF-8 BOM)
      const bom = '\uFEFF';
      const csv = bom + body;
      const encoded = Utilities.newBlob(csv, 'text/csv;charset=utf-8');
      const bytes = encoded.getBytes();
      const csvB64 = Utilities.base64Encode(bytes);

      return { success: true, csvBase64: csvB64, csv: csv, excelBase64: excelB64, filename: filename, encoding: 'utf-8-bom' };
    } catch (e) {
      // fallback: return CSV with BOM
      const bom = '\uFEFF';
      const csv = bom + body;
      return { success: true, csv: csv, filename: filename, encoding: 'utf-8' };
    }
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/* --------------------------------------------------------------------------
 * 8. アンケート（Google Form 回答）表示機能
 * - スプレッドシート内の各シートを走査し、以下のいずれかでアンケートシートと判定する
 *   1) A1 に URL が含まれる
 *   2) ヘッダにメールアドレス(メール|Email) とタイムスタンプ(Timestamp|回答日等)が含まれる
 * - listSurveys(): アンケート一覧（タイトル、最新スコア、最新回答日）を返す
 * - getSurveyDetails(sheetName): 指定シートの全回答・最新回答（メールで一意化）・スコア統計を返す
 * -------------------------------------------------------------------------- */

function _isUrl(s) {
  if (!s) return false;
  try { return /https?:\/\//i.test(String(s)); } catch (e) { return false; }
}

function _extractSpreadsheetId(urlOrId) {
  if (!urlOrId) return null;
  const s = String(urlOrId).trim();
  // direct id
  const direct = s.match(/^[-\w]{25,}$/);
  if (direct) return direct[0];
  // url
  const m = s.match(/\/d\/([a-zA-Z0-9-_]+)\//);
  if (m && m[1]) return m[1];
  const m2 = s.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (m2 && m2[1]) return m2[1];
  return null;
}

function _findHeaderIndex(headers, patterns) {
  if (!headers || !headers.length) return -1;
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || '').trim();
    for (let p of patterns) {
      const re = new RegExp(p, 'i');
      if (re.test(h)) return i;
    }
  }
  return -1;
}

// Try to detect which row contains the header (search first few rows for common header patterns)
function _detectHeaderRow(sheet, lastCol) {
  try {
    const maxLook = Math.min(5, Math.max(1, sheet.getLastRow()));
    for (let r = 1; r <= maxLook; r++) {
      const rowVals = sheet.getRange(r, 1, 1, lastCol).getValues()[0] || [];
      const emailIdx = _findHeaderIndex(rowVals, ['メール', 'メールアドレス', '^email$','^e-mail$']);
      const timeIdx = _findHeaderIndex(rowVals, ['タイムスタンプ','Timestamp','回答日時','回答日','日時']);
      if (emailIdx >= 0 || timeIdx >= 0) return r;
    }
  } catch (e) {
    // ignore and fallback to 1
  }
  return 1;
}

function _safeValueForClient(v) {
  if (v === null || typeof v === 'undefined') return '';
  if (Object.prototype.toString.call(v) === '[object Date]' || v instanceof Date) {
    try {
      return Utilities.formatDate(new Date(v), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    } catch (e) {
      return String(v);
    }
  }
  if (typeof v === 'number' || typeof v === 'boolean') return v;
  return String(v);
}

function listSurveys(sessionToken) {
  try {
    const ssMainLog = SpreadsheetApp.openById(getSpreadsheetId());
    const sheetLogs = ssMainLog.getSheetByName(SHEET_LOGS) || ssMainLog.insertSheet(SHEET_LOGS);
    sheetLogs.appendRow([new Date(), 'listSurveys', sessionToken || '', 'start', '']);
  } catch (e) {
    // ignore logging failure
  }
  const login = getLoginUser(sessionToken);
  // require login for viewing details (we need an identity to check user's response)
  if (!login || login.status !== 'authorized') {
    return { success: false, message: '参照するにはログインが必要です。ログインしてください。' };
  }
  const userEmail = (login && login.user) ? String(login.user.email || '').trim().toLowerCase() : '';
  // try to get student id from Users sheet if available
  let userStudentId = null;
  try {
    if (userEmail) {
      const ssMain = SpreadsheetApp.openById(getSpreadsheetId());
      const usersSheet = ssMain.getSheetByName(SHEET_USERS);
      if (usersSheet) {
        const data = usersSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          if (String(row[COL.EMAIL] || '').trim().toLowerCase() === userEmail) { userStudentId = String(row[COL.STUDENT_ID] || '').trim(); break; }
        }
      }
    }
  } catch (e) { /* ignore */ }

  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const formsSheet = ss.getSheetByName(SHEET_FORMS);
  const out = [];
  if (!formsSheet) return out;
  const lastRow = formsSheet.getLastRow();
  if (lastRow < 2) return out;
  const cols = Math.max(3, HEADER_FORMS.length);
  const rows = formsSheet.getRange(2,1,Math.max(0,lastRow-1), cols).getValues();
  rows.forEach((r, idx) => {
    try {
      const spreadRef = String(r[0] || '').trim();
      const formUrl = String(r[1] || '').trim();
      const title = String(r[2] || '').trim() || spreadRef;
      const inChargeOrg = String(r[3] || '').trim() || '';
      const inChargeDept = String(r[4] || '').trim() || '';
      const collectingRaw = r[5];
      const collecting = (collectingRaw === true) || (String(collectingRaw || '').toLowerCase() === 'true');
      const scoreName = String(r[6] || '').trim() || null;
      const scoreUnit = String(r[7] || '').trim() || null;
      const sid = _extractSpreadsheetId(spreadRef) || _extractSpreadsheetId(formUrl);
      if (!sid) {
        out.push({ title: title, spreadsheetId: null, spreadsheetUrl: null, formUrl: formUrl || null, inChargeOrg: inChargeOrg, inChargeDept: inChargeDept, collecting: collecting, scoreName: scoreName, scoreUnit: scoreUnit, userLatestRowIndex: null, available: false, latestResponseDate: null, latestScore: null });
        return;
      }
      // Determine whether the current user has a response in this survey sheet
      let available = false;
      let latestResponseDate = null;
      let latestScore = null;
      let userLatestRowIndex = null;
      try {
        const targetSs = SpreadsheetApp.openById(sid);
        const sSh = targetSs.getSheets()[0];
        const sLastCol = Math.max(1, sSh.getLastColumn());
        const sHeaderRow = _detectHeaderRow(sSh, sLastCol);
        const headers = sSh.getRange(sHeaderRow, 1, 1, sLastCol).getValues()[0] || [];
        const timeIdx = _findHeaderIndex(headers, ['タイムスタンプ','Timestamp','回答日時','回答日','日時']);
        const emailIdx = _findHeaderIndex(headers, ['メール', 'メールアドレス', '^email$','^e-mail$']);
        const sidIdx = _findHeaderIndex(headers, ['学籍番号', 'student id', 'studentid', '学籍']);
        const scoreIdx = _findHeaderIndex(headers, ['スコア','Score','合計','点数']);

        const dataStart = sHeaderRow + 1;
        const dataCount = Math.max(0, sSh.getLastRow() - sHeaderRow);
        if (dataCount > 0 && (emailIdx >= 0 || sidIdx >= 0)) {
          const data = sSh.getRange(dataStart, 1, dataCount, sLastCol).getValues();
          let latestTs = -1;
          for (let i = 0; i < data.length; i++) {
            const row = data[i] || [];
            const rowEmail = (emailIdx >= 0) ? String(row[emailIdx] || '').trim().toLowerCase() : '';
            const rowSid = (sidIdx >= 0) ? String(row[sidIdx] || '').trim() : '';
            const emailMatch = userEmail && rowEmail && rowEmail === userEmail;
            const sidMatch = userStudentId && rowSid && rowSid === userStudentId;
            if (!emailMatch && !sidMatch) continue;

            let ts = i; // fallback when timestamp column is missing or invalid
            if (timeIdx >= 0) {
              const t = row[timeIdx];
              if (t instanceof Date) ts = t.getTime();
              else {
                const tt = Date.parse(String(t || ''));
                if (!isNaN(tt)) ts = tt;
              }
            }
            if (ts >= latestTs) {
              latestTs = ts;
              available = true;
              latestResponseDate = ts;
              latestScore = (scoreIdx >= 0) ? row[scoreIdx] : null;
              userLatestRowIndex = dataStart + i;
            }
          }
        }
      } catch (e) {
        // ignore per-sheet failures
      }
      // Performance improvement: avoid opening each spreadsheet synchronously.
      // Instead, fetch Drive file metadata (modifiedDate) and return lightweight info.
      let latestDate = null;
      try {
        // Try Advanced Drive API first (faster for metadata)
        try {
          const meta = Drive.Files.get(sid, { fields: 'modifiedDate' });
          if (meta && meta.modifiedDate) latestDate = (new Date(meta.modifiedDate)).getTime();
        } catch (e) {
          // Fallback to DriveApp
          try { const f = DriveApp.getFileById(sid); if (f && f.getLastUpdated) latestDate = f.getLastUpdated().getTime(); } catch (ee) { /* ignore */ }
        }
      } catch (e) {
        // ignore
      }
      const spreadsheetUrl = (spreadRef && String(spreadRef).indexOf('http')===0) ? spreadRef : ('https://docs.google.com/spreadsheets/d/' + sid + '/edit');
      const latestScoreFormatted = (latestScore !== null && latestScore !== undefined) ? (Number(latestScore).toLocaleString() + (scoreUnit ? (' ' + scoreUnit) : '')) : null;
      out.push({
        title: title,
        spreadsheetId: sid,
        spreadsheetUrl: spreadsheetUrl,
        formUrl: formUrl || null,
        inChargeOrg: inChargeOrg,
        inChargeDept: inChargeDept,
        collecting: collecting,
        scoreName: scoreName,
        scoreUnit: scoreUnit,
        userLatestRowIndex: userLatestRowIndex,
        available: available,
        latestResponseDate: available ? latestResponseDate : (latestDate || null),
        latestScore: available ? latestScore : null,
        latestScoreFormatted: available ? latestScoreFormatted : null
      });
    } catch (e) { out.push({ title: String(r[2]||r[0]||''), spreadsheetId: null, spreadsheetUrl: null, userLatestRowIndex: null, available: false, latestResponseDate: null, latestScore: null }); }
  });
  // sort by latestResponseDate desc
  out.sort((a,b) => (b.latestResponseDate || 0) - (a.latestResponseDate || 0));
  try {
    // sort by latestResponseDate desc
    out.sort((a,b) => (b.latestResponseDate || 0) - (a.latestResponseDate || 0));
    try {
      const ssMainLog2 = SpreadsheetApp.openById(getSpreadsheetId());
      const sheetLogs2 = ssMainLog2.getSheetByName(SHEET_LOGS) || ssMainLog2.insertSheet(SHEET_LOGS);
      sheetLogs2.appendRow([new Date(), 'listSurveys', sessionToken || '', 'end', JSON.stringify({ count: out.length })]);
    } catch (e) {}
    return out;
  } catch (e) {
    try {
      const ssMainLog3 = SpreadsheetApp.openById(getSpreadsheetId());
      const sheetLogs3 = ssMainLog3.getSheetByName(SHEET_LOGS) || ssMainLog3.insertSheet(SHEET_LOGS);
      sheetLogs3.appendRow([new Date(), 'listSurveys', sessionToken || '', 'error', String(e)]);
    } catch (ee) {}
    return out;
  }
}

function listFormDefinitions(sessionToken) {
  try {
    const login = getLoginUser(sessionToken);
    if (!login || login.status !== 'authorized') throw new Error('認証されていません');
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const formsSheet = ss.getSheetByName(SHEET_FORMS);
    if (!formsSheet) return { success: true, items: [] };
    const lastRow = formsSheet.getLastRow();
    if (lastRow < 2) return { success: true, items: [] };
    const cols = Math.max(HEADER_FORMS.length, formsSheet.getLastColumn());
    const rows = formsSheet.getRange(2, 1, lastRow - 1, cols).getValues();
    const items = rows.map((r, i) => {
      const collectingRaw = r[5];
      const collecting = (collectingRaw === true) || (String(collectingRaw || '').toLowerCase() === 'true');
      return {
        rowIndex: i + 2,
        spreadsheetRef: String(r[0] || '').trim(),
        formUrl: String(r[1] || '').trim(),
        title: String(r[2] || '').trim(),
        inChargeOrg: String(r[3] || '').trim(),
        inChargeDept: String(r[4] || '').trim(),
        collecting: collecting,
        scoreName: String(r[6] || '').trim(),
        scoreUnit: String(r[7] || '').trim()
      };
    });
    return { success: true, items: items };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function saveFormDefinition(sessionToken, payload) {
  try {
    const login = getLoginUser(sessionToken);
    if (!login || login.status !== 'authorized') throw new Error('認証されていません');
    const data = payload || {};
    const spreadsheetRef = String(data.spreadsheetRef || '').trim();
    const formUrl = String(data.formUrl || '').trim();
    if (!spreadsheetRef && !formUrl) throw new Error('アンケートシートまたはフォームURLのどちらかを入力してください');
    const rowIndex = Number(data.rowIndex || 0);
    const row = new Array(HEADER_FORMS.length).fill('');
    row[0] = spreadsheetRef;
    row[1] = formUrl;
    row[2] = String(data.title || '').trim();
    row[3] = String(data.inChargeOrg || '').trim();
    row[4] = String(data.inChargeDept || '').trim();
    row[5] = data.collecting ? true : false;
    row[6] = String(data.scoreName || '').trim();
    row[7] = String(data.scoreUnit || '').trim();

    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const formsSheet = ss.getSheetByName(SHEET_FORMS) || ss.insertSheet(SHEET_FORMS);
    // ensure header row exists
    if (formsSheet.getLastRow() === 0) {
      formsSheet.getRange(1, 1, 1, HEADER_FORMS.length).setValues([HEADER_FORMS]);
    }

    if (rowIndex && rowIndex >= 2) {
      formsSheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    } else {
      formsSheet.appendRow(row);
    }
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getSurveyDetails(sessionToken, spreadsheetRef, rowIndex) {
  try {
    try {
      const ssMainLog = SpreadsheetApp.openById(getSpreadsheetId());
      const logs = ssMainLog.getSheetByName(SHEET_LOGS) || ssMainLog.insertSheet(SHEET_LOGS);
      logs.appendRow([new Date(), 'getSurveyDetails', spreadsheetRef || '', 'start', String(rowIndex || '')]);
    } catch (e) {}
  // spreadsheetRef: spreadsheet URL or ID (from Forms.A)
  const login = getLoginUser(sessionToken);
  const userEmail = (login && login.user) ? String(login.user.email || '').trim().toLowerCase() : '';
  // get user student id
  let userStudentId = null;
  try {
    if (userEmail) {
      const ssMain = SpreadsheetApp.openById(getSpreadsheetId());
      const usersSheet = ssMain.getSheetByName(SHEET_USERS);
      if (usersSheet) {
        const data = usersSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          if (String(row[COL.EMAIL] || '').trim().toLowerCase() === userEmail) { userStudentId = String(row[COL.STUDENT_ID] || '').trim(); break; }
        }
      }
    }
  } catch (e) { /* ignore */ }

  const sid = _extractSpreadsheetId(spreadsheetRef);
  if (!sid) throw new Error('無効なスプレッドシート参照です');
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  let targetSs;
  try { targetSs = SpreadsheetApp.openById(sid); } catch (e) { throw new Error('対象スプレッドシートを開けません: ' + e.toString()); }
  const sheets = targetSs.getSheets();
  if (!sheets || sheets.length === 0) throw new Error('対象シートがありません');
  const sh = sheets[0];

  const lastCol = Math.max(1, sh.getLastColumn());
  const headerRow = _detectHeaderRow(sh, lastCol);
  const headers = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0] || [];

  const timeIdx = _findHeaderIndex(headers, ['タイムスタンプ','Timestamp','回答日時','回答日','日時']);
  const emailIdx = _findHeaderIndex(headers, ['メール', 'メールアドレス', '^email$','^e-mail$']);
  const sidIdx = _findHeaderIndex(headers, ['学籍番号', 'student id', 'studentid', '学籍']);
  const scoreIdx = _findHeaderIndex(headers, ['スコア','Score','合計','点数']);

  // If rowIndex is provided, return single response for that sheet row (sheet-row index expected)
  if (typeof rowIndex !== 'undefined' && rowIndex !== null) {
    const ri = Number(rowIndex);
    if (isNaN(ri) || ri <= headerRow || ri > sh.getLastRow()) return { success: false, message: '無効な行番号です' };
    const r = sh.getRange(ri, 1, 1, lastCol).getValues()[0] || [];
    const obj = { answers: {}, timestamp: null, email: null, score: null, studentId: null };
    if (timeIdx >= 0) {
      const t = r[timeIdx];
      if (t instanceof Date) obj.timestamp = t;
      else if (String(t || '').trim()) { const dd = new Date(String(t)); if (!isNaN(dd.getTime())) obj.timestamp = dd; }
    }
    if (emailIdx >= 0) obj.email = String(r[emailIdx] || '').trim().toLowerCase();
    if (sidIdx >= 0) obj.studentId = String(r[sidIdx] || '').trim();
    if (scoreIdx >= 0) obj.score = r[scoreIdx];
    for (let i = 0; i < headers.length; i++) obj.answers[headers[i] || ('col' + (i+1))] = _safeValueForClient(r[i]);

    // ensure requester owns this response (email or studentId)
    let ok = false;
    if (userEmail && obj.email && String(obj.email).trim().toLowerCase() === userEmail) ok = true;
    if (!ok && userStudentId && obj.studentId && String(obj.studentId).trim() === userStudentId) ok = true;
    if (!ok) return { success: false, message: 'あなたの回答がないため参照できません' };

    const conv = Object.assign({}, obj);
    conv.timestamp = conv.timestamp ? (conv.timestamp instanceof Date ? conv.timestamp.getTime() : Number(conv.timestamp)) : null;
    conv.score = _safeValueForClient(conv.score);

    // fetch scoreName/scoreUnit from Forms sheet if available
    let scoreName = null, scoreUnit = null;
    try {
      const formsSheet = ss.getSheetByName(SHEET_FORMS);
      if (formsSheet) {
        const cols = Math.max(3, HEADER_FORMS.length);
        const metaRows = formsSheet.getRange(2,1,Math.max(0, formsSheet.getLastRow()-1), cols).getValues();
        for (let i = 0; i < metaRows.length; i++) {
          const a = String(metaRows[i][0] || '').trim();
          const b = String(metaRows[i][1] || '').trim();
          const fid = _extractSpreadsheetId(a) || _extractSpreadsheetId(b);
          if (fid === sid || a === spreadsheetRef || b === spreadsheetRef) {
            scoreName = String(metaRows[i][6] || '').trim() || null;
            scoreUnit = String(metaRows[i][7] || '').trim() || null;
            break;
          }
        }
      }
    } catch (e) { /* ignore */ }

    conv.scoreFormatted = (conv.score !== null && conv.score !== undefined) ? (Number(conv.score).toLocaleString() + (scoreUnit ? (' ' + scoreUnit) : '')) : null;

    try {
      const ssMainLogEnd = SpreadsheetApp.openById(getSpreadsheetId());
      const logsEnd = ssMainLogEnd.getSheetByName(SHEET_LOGS) || ssMainLogEnd.insertSheet(SHEET_LOGS);
      logsEnd.appendRow([new Date(), 'getSurveyDetails', spreadsheetRef || '', 'end', 'rowIndex:' + ri]);
    } catch (e) {}
    return { success: true, sheetRef: spreadsheetRef, rowIndex: ri, headers: headers, response: conv, scoreName: scoreName, scoreUnit: scoreUnit };
  }

  // otherwise fall back to full-sheet behavior (backwards compatible)
  const dataStart = headerRow + 1;
  const dataCount = Math.max(0, sh.getLastRow() - headerRow);
  const rows = (dataCount > 0) ? sh.getRange(dataStart,1,dataCount,lastCol).getValues() : [];

  // parse rows
  const parsed = rows.map(r => {
    const obj = { answers: {}, timestamp: null, email: null, score: null, studentId: null };
    if (timeIdx >= 0) {
      const t = r[timeIdx];
      if (t instanceof Date) obj.timestamp = t;
      else if (String(t || '').trim()) { const dd = new Date(String(t)); if (!isNaN(dd.getTime())) obj.timestamp = dd; }
    }
    if (emailIdx >= 0) obj.email = String(r[emailIdx] || '').trim().toLowerCase();
    if (sidIdx >= 0) obj.studentId = String(r[sidIdx] || '').trim();
    if (scoreIdx >= 0) obj.score = r[scoreIdx];
    for (let i = 0; i < headers.length; i++) obj.answers[headers[i] || ('col' + (i+1))] = _safeValueForClient(r[i]);
    return obj;
  }).filter(p => p.timestamp || p.email || Object.keys(p.answers).length > 0);

  // determine whether current user has response (email first, then studentId)
  let hasResponse = false;
  if (userEmail) {
    if (emailIdx >= 0) {
      hasResponse = parsed.some(p => p.email && String(p.email).trim().toLowerCase() === userEmail);
    }
  }
  if (!hasResponse && userStudentId && sidIdx >= 0) {
    hasResponse = parsed.some(p => p.studentId && String(p.studentId).trim() === userStudentId);
  }

  if (!hasResponse) return { success: false, message: 'あなたの回答がないため参照できません' };

  // sort by timestamp desc
  parsed.sort((a,b) => {
    const ta = a.timestamp ? (a.timestamp instanceof Date ? a.timestamp.getTime() : Number(a.timestamp)) : 0;
    const tb = b.timestamp ? (b.timestamp instanceof Date ? b.timestamp.getTime() : Number(b.timestamp)) : 0;
    return tb - ta;
  });

  // latest per email (メールで一意化)、fallback to studentId grouping if email missing
  const latestByKey = {};
  parsed.forEach(p => {
    let key = p.email || (p.studentId ? ('sid:' + p.studentId) : ('row' + Math.random()));
    if (!latestByKey[key] || (p.timestamp && latestByKey[key].timestamp && p.timestamp.getTime() > latestByKey[key].timestamp.getTime())) {
      latestByKey[key] = p;
    }
  });
  const latestList = Object.keys(latestByKey).map(k => latestByKey[k]);
  latestList.sort((a,b) => (b.timestamp?b.timestamp.getTime():0) - (a.timestamp?a.timestamp.getTime():0));

  const scores = parsed.map(p => (typeof p.score === 'number' ? p.score : (p.score ? Number(p.score) : null))).filter(v => v !== null && !isNaN(v));
  const stats = { count: scores.length, min: null, max: null, avg: null, distribution: {} };
  if (scores.length > 0) {
    const s = scores.slice().sort((a,b)=>a-b);
    stats.min = s[0]; stats.max = s[s.length-1]; stats.avg = s.reduce((a,b)=>a+b,0)/s.length;
    s.forEach(v => { const k = String(v); stats.distribution[k] = (stats.distribution[k] || 0) + 1; });
  }

  // also include scoreName/scoreUnit from Forms sheet when available
  let scoreName = null, scoreUnit = null;
  try {
    const formsSheet = ss.getSheetByName(SHEET_FORMS);
    if (formsSheet) {
      const cols = Math.max(3, HEADER_FORMS.length);
      const metaRows = formsSheet.getRange(2,1,Math.max(0, formsSheet.getLastRow()-1), cols).getValues();
      for (let i = 0; i < metaRows.length; i++) {
        const a = String(metaRows[i][0] || '').trim();
        const b = String(metaRows[i][1] || '').trim();
        const fid = _extractSpreadsheetId(a) || _extractSpreadsheetId(b);
        if (fid === sid || a === spreadsheetRef || b === spreadsheetRef) {
          scoreName = String(metaRows[i][6] || '').trim() || null;
          scoreUnit = String(metaRows[i][7] || '').trim() || null;
          break;
        }
      }
    }
  } catch (e) { /* ignore */ }

  // format stats
  stats.minFormatted = stats.min !== null ? Number(stats.min).toLocaleString() + (scoreUnit ? (' ' + scoreUnit) : '') : null;
  stats.maxFormatted = stats.max !== null ? Number(stats.max).toLocaleString() + (scoreUnit ? (' ' + scoreUnit) : '') : null;
  stats.avgFormatted = stats.avg !== null ? Number(stats.avg).toLocaleString() + (scoreUnit ? (' ' + scoreUnit) : '') : null;
  stats.distributionFormatted = {};
  Object.keys(stats.distribution).forEach(k => {
    const formatted = Number(k).toLocaleString() + (scoreUnit ? (' ' + scoreUnit) : '');
    stats.distributionFormatted[formatted] = stats.distribution[k];
  });

  // convert timestamps to epoch millis for safe JSON serialization to client
  const convParsed = parsed.map(p => {
    const copy = Object.assign({}, p);
    copy.timestamp = copy.timestamp ? (copy.timestamp instanceof Date ? copy.timestamp.getTime() : Number(copy.timestamp)) : null;
    copy.scoreFormatted = (copy.score !== null && copy.score !== undefined) ? (Number(copy.score).toLocaleString() + (scoreUnit ? (' ' + scoreUnit) : '')) : null;
    return copy;
  });
  const convLatest = latestList.map(p => {
    const copy = Object.assign({}, p);
    copy.timestamp = copy.timestamp ? (copy.timestamp instanceof Date ? copy.timestamp.getTime() : Number(copy.timestamp)) : null;
    copy.scoreFormatted = (copy.score !== null && copy.score !== undefined) ? (Number(copy.score).toLocaleString() + (scoreUnit ? (' ' + scoreUnit) : '')) : null;
    return copy;
  });

  try {
    const ssMainLogEnd2 = SpreadsheetApp.openById(getSpreadsheetId());
    const logsEnd2 = ssMainLogEnd2.getSheetByName(SHEET_LOGS) || ssMainLogEnd2.insertSheet(SHEET_LOGS);
    logsEnd2.appendRow([new Date(), 'getSurveyDetails', spreadsheetRef || '', 'end', 'rows:' + convParsed.length]);
  } catch (e) {}

  return { success: true, sheetRef: spreadsheetRef, headers: headers, allResponses: convParsed, latestPerEmail: convLatest, scoreStats: stats, scoreName: scoreName, scoreUnit: scoreUnit };
  } catch (e) {
    const msg = (e && e.toString) ? e.toString() : String(e);
    try {
      const ss = SpreadsheetApp.openById(getSpreadsheetId());
      const logs = ss.getSheetByName(SHEET_LOGS) || ss.insertSheet(SHEET_LOGS);
      const time = new Date();
      logs.appendRow([time, 'getSurveyDetails', spreadsheetRef || '', 'error', msg]);
    } catch (ee) {
      // ignore logging failure
    }
    return { success: false, message: 'サーバー処理中にエラーが発生しました: ' + msg };
  }
}

// Try to extract Form ID from a URL like https://docs.google.com/forms/d/FORM_ID/
function _extractFormId(formUrl) {
  if (!formUrl) return null;
  try {
    // support URLs like /forms/d/ID/ and /forms/d/e/ID/
    const m = String(formUrl).match(/\/forms\/d\/(?:e\/)?([-_0-9A-Za-z]+)/);
    if (m && m[1]) return m[1];
  } catch (e) {}
  return null;
}

/* -----------------------------
 * Collections (集金) 機能
 * -----------------------------*/
const SHEET_COLLECTIONS = 'Collections';
const SHEET_COLLECTIONS_LOG = 'Collections_log';

const HEADER_COLLECTIONS = ['id','タイトル','スプレッドシートURL','担当局','担当部門','作成日時','作成者'];
const HEADER_COLLECTIONS_LOG = ['Collections_id','タイムスタンプ','取引先メールアドレス','取引種別','取引金額','担当者メールアドレス'];

function ensureCollectionsSheets() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  let s = ss.getSheetByName(SHEET_COLLECTIONS);
  if (!s) s = ss.insertSheet(SHEET_COLLECTIONS);
  if (s.getLastRow() === 0) s.getRange(1,1,1,HEADER_COLLECTIONS.length).setValues([HEADER_COLLECTIONS]);

  let sl = ss.getSheetByName(SHEET_COLLECTIONS_LOG);
  if (!sl) sl = ss.insertSheet(SHEET_COLLECTIONS_LOG);
  if (sl.getLastRow() === 0) sl.getRange(1,1,1,HEADER_COLLECTIONS_LOG.length).setValues([HEADER_COLLECTIONS_LOG]);

  return { success: true };
}

function _generateCollectionId() {
  return 'COL' + String(Date.now());
}

function listCollections(sessionToken) {
  // write start log
  try {
    const ssLog = SpreadsheetApp.openById(getSpreadsheetId());
    const logSheet = ssLog.getSheetByName(SHEET_LOGS) || ssLog.insertSheet(SHEET_LOGS);
    logSheet.appendRow([new Date(), 'listCollections', sessionToken || '', 'start', '']);
  } catch (e) { /* ignore logging errors */ }

  try {
    ensureCollectionsSheets();
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const s = ss.getSheetByName(SHEET_COLLECTIONS);
    if (!s) {
      try { const ssLog2 = SpreadsheetApp.openById(getSpreadsheetId()); const ls2 = ssLog2.getSheetByName(SHEET_LOGS) || ssLog2.insertSheet(SHEET_LOGS); ls2.appendRow([new Date(), 'listCollections', sessionToken || '', 'end', JSON.stringify({ count:0 })]); } catch(e){}
      return [];
    }
    const lr = s.getLastRow();
    if (lr < 2) {
      try { const ssLog3 = SpreadsheetApp.openById(getSpreadsheetId()); const ls3 = ssLog3.getSheetByName(SHEET_LOGS) || ssLog3.insertSheet(SHEET_LOGS); ls3.appendRow([new Date(), 'listCollections', sessionToken || '', 'end', JSON.stringify({ count:0 })]); } catch(e){}
      return [];
    }
    const rows = s.getRange(2,1,lr-1, Math.max(HEADER_COLLECTIONS.length, s.getLastColumn())).getValues();
    const out = rows.map(r => {
      const createdRaw = r[5];
      let createdVal = null;
      try {
        if (createdRaw instanceof Date) createdVal = createdRaw.getTime();
        else if (typeof createdRaw === 'number') createdVal = createdRaw;
        else if (createdRaw) createdVal = String(createdRaw);
        else createdVal = null;
      } catch (e) { createdVal = String(createdRaw); }
      return { id: String(r[0]||''), title: String(r[1]||''), spreadsheetUrl: String(r[2]||''), inChargeOrg: String(r[3]||''), inChargeDept: String(r[4]||''), createdAt: createdVal, createdBy: String(r[6]||'') };
    });
    try { const ssLog4 = SpreadsheetApp.openById(getSpreadsheetId()); const ls4 = ssLog4.getSheetByName(SHEET_LOGS) || ssLog4.insertSheet(SHEET_LOGS); ls4.appendRow([new Date(), 'listCollections', sessionToken || '', 'end', JSON.stringify({ count: out.length })]); } catch(e){}
    return out;
  } catch (e) {
    try {
      const ssLog = SpreadsheetApp.openById(getSpreadsheetId());
      const logSheet = ssLog.getSheetByName(SHEET_LOGS) || ssLog.insertSheet(SHEET_LOGS);
      logSheet.appendRow([new Date(), 'listCollections', sessionToken || '', 'error', String(e)]);
    } catch (ee) {
      console.error('listCollections logging failed', ee.toString());
    }
    return [];
  }
}

function createCollection(sessionToken, payload) {
  // payload: { title, spreadsheetUrl, inChargeOrg, inChargeDept }
  ensureCollectionsSheets();
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const s = ss.getSheetByName(SHEET_COLLECTIONS);
  const id = _generateCollectionId();
  const now = new Date();
  let createdBy = 'unknown';
  try { const sessions = _loadStore(SESSIONS_PROP_KEY); if (sessionToken && sessions[sessionToken]) createdBy = sessions[sessionToken].email || createdBy; } catch(e){}
  // normalize placeholder values
  let orgVal = payload.inChargeOrg || '';
  let deptVal = payload.inChargeDept || '';
  if (orgVal === '選択' || orgVal === '選択してください') orgVal = '';
  if (deptVal === '選択' || deptVal === '選択してください') deptVal = '';
  const row = [id, payload.title || '', payload.spreadsheetUrl || '', orgVal || '', deptVal || '', now, createdBy];
  s.appendRow(row);
  // log creation
  try {
    const ls = ss.getSheetByName(SHEET_LOGS) || ss.insertSheet(SHEET_LOGS);
    ls.appendRow([new Date(), 'createCollection', sessionToken || '', 'created', JSON.stringify({ id: id, title: payload.title || '' })]);
  } catch (e) {}
  return { success: true, id };
}

// Use the primary helpers defined earlier in the file.
// Provide a utility that returns all matching header indices (plural) without overriding the single-index helper.
function _findHeaderIndices(headers, patterns) {
  const matches = [];
  if (!headers || !headers.length) return matches;
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || '').trim();
    for (let p of patterns) {
      const re = new RegExp(p, 'i');
      if (re.test(h)) { matches.push(i); break; }
    }
  }
  return matches;
}

function parseSourceSpreadsheet(spreadsheetRef) {
  const sid = _extractSpreadsheetId(spreadsheetRef) || spreadsheetRef;
  if (!sid) return { success: false, message: 'スプレッドシートIDが取得できません' };
  let target;
  try { target = SpreadsheetApp.openById(sid); } catch (e) { return { success:false, message: 'スプレッドシートを開けません: '+ String(e) }; }
  const sheet = target.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  if (!data || data.length < 1) return { success: true, headers: [], rows: [] };
  const headers = (data[0] || []).map(c => String(c||''));

  // fuzzy match email
  const emailPatterns = ['メールアドレス','メール','アドレス','Email','E-mail','e-mail'];
  const amountPatterns = ['金額','集金額','支払金額','請求額','スコア','score','amount'];

  const emailMatches = _findHeaderIndices(headers, emailPatterns);
  const amountMatches = _findHeaderIndices(headers, amountPatterns);

  if (emailMatches.length === 0) return { success:false, message: 'メールアドレス列が見つかりません' };
  if (amountMatches.length === 0) return { success:false, message: '金額列が見つかりません' };
  if (emailMatches.length > 1) return { success:false, message: 'メールアドレス列が複数見つかりました' };
  if (amountMatches.length > 1) return { success:false, message: '金額列が複数見つかりました' };

  const emailIdx = emailMatches[0];
  const amountIdx = amountMatches[0];

  const rows = [];
  const emailsSeen = {};
  const dupEmails = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const mail = String(row[emailIdx]||'').trim();
    let amount = row[amountIdx];
    if (typeof amount === 'string') {
      amount = Number(String(amount).replace(/[^0-9\-\.]/g,'')) || 0;
    }
    amount = Number(amount) || 0;
    if (mail) {
      const key = mail.toLowerCase();
      if (emailsSeen[key]) dupEmails.push(mail);
      emailsSeen[key] = true;
    }
    rows.push({ rowIndex: i+1, email: mail, amount });
  }
  return { success: true, headers, rows, emailColumn: emailIdx, amountColumn: amountIdx, duplicateEmails: dupEmails };
}

function fetchCollectionSummary(sessionToken, collectionId) {
  // log start
  try { const ssLog = SpreadsheetApp.openById(getSpreadsheetId()); const ls = ssLog.getSheetByName(SHEET_LOGS) || ssLog.insertSheet(SHEET_LOGS); ls.appendRow([new Date(), 'fetchCollectionSummary', sessionToken || '', 'start', collectionId]); } catch(e){}
  ensureCollectionsSheets();
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const s = ss.getSheetByName(SHEET_COLLECTIONS);
  const lr = s.getLastRow();
  if (lr < 2) return { success: false, message: 'Collectionsが空です' };
  const rows = s.getRange(2,1,lr-1, HEADER_COLLECTIONS.length).getValues();
  const found = rows.find(r => String(r[0]||'') === String(collectionId));
  if (!found) return { success:false, message: '指定のCollectionが見つかりません' };
  const col = { id: String(found[0]||''), title: String(found[1]||''), spreadsheetUrl: String(found[2]||''), inChargeOrg: String(found[3]||''), inChargeDept: String(found[4]||''), createdAt: found[5], createdBy: String(found[6]||'') };

  const parsed = parseSourceSpreadsheet(col.spreadsheetUrl);
  if (!parsed.success) return parsed;

  // read logs
  const sl = ss.getSheetByName(SHEET_COLLECTIONS_LOG);
  const lrl = sl.getLastRow();
  const logs = lrl >= 2 ? sl.getRange(2,1,lrl-1, HEADER_COLLECTIONS_LOG.length).getValues().map(r=>({ collectionId: String(r[0]||''), timestamp: r[1], email: String(r[2]||''), type: String(r[3]||''), amount: Number(r[4]||0), handler: String(r[5]||'')})) : [];

  // normalize log timestamps to primitive values for JSON serialization
  const normalizeTimestamp = (t) => {
    if (!t && t !== 0) return null;
    if (t instanceof Date) return t.getTime();
    if (typeof t === 'number') return t;
    return String(t);
  };

  const logsOf = logs.filter(l => l.collectionId === col.id).map(l => ({
    collectionId: l.collectionId,
    timestamp: normalizeTimestamp(l.timestamp),
    email: l.email,
    type: l.type,
    amount: l.amount,
    handler: l.handler
  }));

  const collectedByEmail = {};
  logsOf.forEach(l => {
    const key = (l.email || '').toLowerCase();
    if (!collectedByEmail[key]) collectedByEmail[key] = { email: l.email, total: 0, entries: [] };
    let sign = 1;
    if (l.type === 'おつり') sign = -1; // treat change as negative
    if (l.type === '返金') sign = -1; // refund reduces collected
    collectedByEmail[key].total += Number(l.amount || 0) * sign;
    collectedByEmail[key].entries.push(l);
  });

  // aggregate by handler (collector) so UI can show "受け取った人" (担当者) breakdown
  const collectedByHandler = {};
  logsOf.forEach(l => {
    const handlerKey = (l.handler || '').toLowerCase();
    if (!collectedByHandler[handlerKey]) collectedByHandler[handlerKey] = { handler: l.handler, total: 0, entries: [] };
    let signH = 1;
    if (l.type === 'おつり') signH = -1;
    if (l.type === '返金') signH = -1;
    collectedByHandler[handlerKey].total += Number(l.amount || 0) * signH;
    collectedByHandler[handlerKey].entries.push(l);
  });

  const perPerson = parsed.rows.map(r => {
    const key = (r.email || '').toLowerCase();
    const expected = Number(r.amount || 0);
    const collected = (collectedByEmail[key] && Number(collectedByEmail[key].total)) || 0;
    let status = '正';
    if (collected < expected) status = '不足';
    if (collected > expected) status = '過払い';
    const entries = (collectedByEmail[key] && Array.isArray(collectedByEmail[key].entries)) ? collectedByEmail[key].entries.map(e => ({
      collectionId: e.collectionId,
      timestamp: e.timestamp,
      email: e.email,
      type: e.type,
      amount: e.amount,
      handler: e.handler
    })) : [];
    return { email: r.email, expected, collected, status, entries };
  });

  const expectedTotal = parsed.rows.reduce((s,r)=>s + Number(r.amount||0),0);
  const expectedCount = parsed.rows.filter(r=>r.email).length;
  const collectedTotal = Object.keys(collectedByEmail).reduce((s,k)=>s + Number(collectedByEmail[k].total||0),0);
  const collectedCount = Object.keys(collectedByEmail).length;

  // normalize collection.createdAt
  try { if (col && col.createdAt instanceof Date) col.createdAt = col.createdAt.getTime(); } catch(e) {}
  try {
    const ls = ss.getSheetByName(SHEET_LOGS) || ss.insertSheet(SHEET_LOGS);
    const sampleEmails = (perPerson || []).slice(0,10).map(p => p.email || '').filter(Boolean);
    ls.appendRow([new Date(), 'fetchCollectionSummary', sessionToken || '', 'end', JSON.stringify({ collectionId: collectionId, expectedCount: expectedCount, collectedCount: collectedCount, perPersonCount: (perPerson||[]).length, sampleEmails: sampleEmails })]);
  } catch(e){}
  // prepare perCollector array for UI (受け取った人の一覧)
  const perCollector = Object.keys(collectedByHandler).map(k => ({ handler: collectedByHandler[k].handler, total: collectedByHandler[k].total, entries: collectedByHandler[k].entries }));

  return { success: true, collection: col, expectedTotal, expectedCount, collectedTotal, collectedCount, perPerson, perCollector, duplicateEmails: parsed.duplicateEmails };
}

function recordPayment(sessionToken, collectionId, recipientEmail, amount, type, handlerEmail) {
  ensureCollectionsSheets();
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const sl = ss.getSheetByName(SHEET_COLLECTIONS_LOG);
  const ts = new Date();
  const a = Number(amount) || 0;
  const row = [collectionId, ts, recipientEmail || '', type || '支払', a, handlerEmail || ''];
  sl.appendRow(row);
  return { success: true };
}

function recordPaymentWithChange(sessionToken, collectionId, recipientEmail, receivedAmount, expectedAmount, handlerEmail) {
  // receivedAmount: actual received (positive). expectedAmount: expected. We record full received as 支払, then record おつり as negative if needed.
  ensureCollectionsSheets();
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const sl = ss.getSheetByName(SHEET_COLLECTIONS_LOG);
  const ts = new Date();
  const r = Number(receivedAmount) || 0;
  const e = Number(expectedAmount) || 0;
  // record full received
  sl.appendRow([collectionId, ts, recipientEmail || '', '支払', r, handlerEmail || '']);
  const change = r - e;
  if (change > 0) {
    // record change as おつり with negative amount
    sl.appendRow([collectionId, ts, recipientEmail || '', 'おつり', -Math.abs(change), handlerEmail || '']);
  }
  return { success: true };
}

function updateCollection(sessionToken, collectionId, payload) {
  try {
    const login = getLoginUser(sessionToken);
    if (!login || login.status !== 'authorized') throw new Error('認証されていません');
    ensureCollectionsSheets();
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const s = ss.getSheetByName(SHEET_COLLECTIONS);
    const lr = s.getLastRow();
    if (lr < 2) throw new Error('Collectionsが空です');
    const rows = s.getRange(2,1,lr-1, HEADER_COLLECTIONS.length).getValues();
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (String(r[0]||'') === String(collectionId)) {
        const rowIndex = i + 2;
        const title = String(payload.title || r[1] || '');
        const spreadsheetUrl = String(payload.spreadsheetUrl || r[2] || '');
        // normalize placeholder values from client
        let inChargeOrg = (typeof payload.inChargeOrg !== 'undefined') ? String(payload.inChargeOrg) : String(r[3] || '');
        let inChargeDept = (typeof payload.inChargeDept !== 'undefined') ? String(payload.inChargeDept) : String(r[4] || '');
        if (inChargeOrg === '選択' || inChargeOrg === '選択してください') inChargeOrg = '';
        if (inChargeDept === '選択' || inChargeDept === '選択してください') inChargeDept = '';
        const createdAt = r[5] || new Date();
        const createdBy = r[6] || '';
        const newRow = [collectionId, title, spreadsheetUrl, inChargeOrg, inChargeDept, createdAt, createdBy];
        s.getRange(rowIndex, 1, 1, newRow.length).setValues([newRow]);
        try { const ls = ss.getSheetByName(SHEET_LOGS) || ss.insertSheet(SHEET_LOGS); ls.appendRow([new Date(), 'updateCollection', sessionToken || '', 'updated', JSON.stringify({ id: collectionId, title: title })]); } catch(e) {}
        return { success: true };
      }
    }
    throw new Error('指定のCollectionが見つかりません');
  } catch (e) {
    try { const ss2 = SpreadsheetApp.openById(getSpreadsheetId()); const ls2 = ss2.getSheetByName(SHEET_LOGS) || ss2.insertSheet(SHEET_LOGS); ls2.appendRow([new Date(), 'updateCollection', sessionToken || '', 'error', String(e)]); } catch(err) {}
    return { success: false, message: e.toString() };
  }
}


function deleteCollection(sessionToken, collectionId) {
  try {
    const login = getLoginUser(sessionToken);
    if (!login || login.status !== 'authorized') throw new Error('認証されていません');
    ensureCollectionsSheets();
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const s = ss.getSheetByName(SHEET_COLLECTIONS);
    const lr = s.getLastRow();
    if (lr < 2) throw new Error('Collectionsが空です');
    const rows = s.getRange(2,1,lr-1, HEADER_COLLECTIONS.length).getValues();
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (String(r[0]||'') === String(collectionId)) {
        const rowIndex = i + 2;
        s.deleteRow(rowIndex);
        try { const ls = ss.getSheetByName(SHEET_LOGS) || ss.insertSheet(SHEET_LOGS); ls.appendRow([new Date(), 'deleteCollection', sessionToken || '', 'deleted', collectionId]); } catch(e) {}
        return { success: true };
      }
    }
    throw new Error('指定のCollectionが見つかりません');
  } catch (e) {
    try { const ss2 = SpreadsheetApp.openById(getSpreadsheetId()); const ls2 = ss2.getSheetByName(SHEET_LOGS) || ss2.insertSheet(SHEET_LOGS); ls2.appendRow([new Date(), 'deleteCollection', sessionToken || '', 'error', String(e)]); } catch(err) {}
    return { success: false, message: e.toString() };
  }
}
