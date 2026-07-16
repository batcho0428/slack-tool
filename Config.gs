function getScriptProperty(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}
function getSpreadsheetId() {
  const id = getScriptProperty('SPREADSHEET_ID');
  if (!id) throw new Error('SPREADSHEET_IDがScript Propertiesに設定されていません');
  return id;
}
function setupSpreadsheet() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());

  const ensureSheetWithHeader = (name, header) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, header.length).setValues([header]);
    else sh.getRange(1, 1, 1, header.length).setValues([header]);
    return sh;
  };

  const sheetGrade = ensureSheetWithHeader(SHEET_GRADE, HEADER_GRADE);
  const sheetOrg = ensureSheetWithHeader(SHEET_ORG, HEADER_ORG);
  const sheetDept = ensureSheetWithHeader(SHEET_DEPT, HEADER_DEPT);
  const sheetRole = ensureSheetWithHeader(SHEET_ROLE, HEADER_ROLE);
  const sheetField = ensureSheetWithHeader(SHEET_FIELD, HEADER_FIELD);

  _applyCheckboxColumn(sheetOrg, 5);
  _applyCheckboxColumn(sheetDept, 5);
  _applyCheckboxColumn(sheetRole, 4);

  // 1b. Formsシート (アンケート情報を専用シートに移行)
  let sheetForms = ss.getSheetByName(SHEET_FORMS);
  if (!sheetForms) sheetForms = ss.insertSheet(SHEET_FORMS);
  // Remove legacy '担当局' column in Forms if present
  try {
    const hdrCols = Math.max(1, sheetForms.getLastColumn());
    const hdrRow = sheetForms.getRange(1, 1, 1, hdrCols).getValues()[0] || [];
    const bureauIdx = hdrRow.findIndex(h => String(h || '').trim() === '担当局');
    if (bureauIdx >= 0) {
      sheetForms.deleteColumn(bureauIdx + 1);
    }
  } catch (e) {}
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

  // 2. Usersシート
  let sheetUsers = ss.getSheetByName(SHEET_USERS);
  if (!sheetUsers) sheetUsers = ss.insertSheet(SHEET_USERS);
  // Remove legacy stray column if it exists.
  try {
    const lastCol = Math.max(1, sheetUsers.getLastColumn());
    const headerRow = sheetUsers.getRange(1, 1, 1, lastCol).getValues()[0] || [];
    let idx = headerRow.findIndex(h => String(h || '').trim() === 'Admin (12)');
    while (idx >= 0) {
      sheetUsers.deleteColumn(idx + 1);
      const newLastCol = Math.max(1, sheetUsers.getLastColumn());
      const newHeaderRow = sheetUsers.getRange(1, 1, 1, newLastCol).getValues()[0] || [];
      idx = newHeaderRow.findIndex(h => String(h || '').trim() === 'Admin (12)');
    }
  } catch (e) {
    // ignore cleanup errors
  }
  sheetUsers.getRange(1, 1, 1, HEADER_USERS.length).setValues([HEADER_USERS]);

  // 3. Logsシート
  let sheetLogs = ss.getSheetByName(SHEET_LOGS);
  if (!sheetLogs) sheetLogs = ss.insertSheet(SHEET_LOGS);
  if (sheetLogs.getLastRow() === 0) sheetLogs.getRange(1, 1, 1, HEADER_LOGS.length).setValues([HEADER_LOGS]);

  // 4. Collections 系シート（集金）
  try { ensureCollectionsSheets(); } catch (e) { /* ignore if cannot create */ }

  // Remove legacy '担当局' column in Collections if present
  try {
    const sheetCollections = ss.getSheetByName(SHEET_COLLECTIONS);
    if (sheetCollections) {
      const hdrColsC = Math.max(1, sheetCollections.getLastColumn());
      const hdrRowC = sheetCollections.getRange(1, 1, 1, hdrColsC).getValues()[0] || [];
      const bureauIdxC = hdrRowC.findIndex(h => String(h || '').trim() === '担当局');
      if (bureauIdxC >= 0) sheetCollections.deleteColumn(bureauIdxC + 1);
    }
  } catch (e) {}

  // Tokens は PropertiesService に移行したためシートは作成しない

  // --- A. 書式設定 ---
  const startRow = 2;
  const numRows = 999;
  sheetUsers.getRange(startRow, 3, numRows, 5).setNumberFormat('@');
  sheetUsers.getRange(startRow, 8, numRows, 1).setNumberFormat('yyyy/mm/dd');
  sheetForms.getRange(startRow, 8, numRows, 1).setNumberFormat('yyyy/mm/dd');

  // --- B. 入力規則 ---
  const rangeGradeOpt = sheetGrade.getRange('A2:A');
  const rangeFieldOpt = sheetField.getRange('A2:A');
  const rangeRoleOpt  = sheetRole.getRange('A2:A');
  const rangeDeptOpt  = sheetDept.getRange('A2:A');

  const buildRule = (range) => SpreadsheetApp.newDataValidation().requireValueInRange(range).setAllowInvalid(true).build();
  const ruleGrade = buildRule(rangeGradeOpt);
  const ruleField = buildRule(rangeFieldOpt);
  const ruleRole  = buildRule(rangeRoleOpt);
  const ruleDept  = buildRule(rangeDeptOpt);

  sheetUsers.getRange(startRow, 4, numRows, 1).setDataValidation(ruleGrade);
  sheetUsers.getRange(startRow, 5, numRows, 1).setDataValidation(ruleField);

  for (let k = 0; k < AFFILIATION_SLOTS; k++) {
    const deptCol = _affiliationDeptCol(k) + 1;
    const roleCol = _affiliationRoleCol(k) + 1;
    sheetUsers.getRange(startRow, deptCol, numRows, 1).setDataValidation(ruleDept);
    sheetUsers.getRange(startRow, roleCol, numRows, 1).setDataValidation(ruleRole);
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

  for (let k = 0; k < AFFILIATION_SLOTS; k++) {
    const colDeptIndex = _affiliationDeptCol(k) + 1;
    const colDeptLet = getColLetter(colDeptIndex);
    const range = sheetUsers.getRange(`${colDeptLet}2:${colDeptLet}`);
    const formula = `=AND(${colDeptLet}2<>"", COUNTIF(INDIRECT("${SHEET_DEPT}!\\$A:\\$A"), ${colDeptLet}2)=0, COUNTIF(INDIRECT("${SHEET_ORG}!\\$A:\\$A"), ${colDeptLet}2)=0)`;
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

  const dailyFuncName = 'closeExpiredSurveyCollectionsDaily';
  const dailyExists = triggers.some(t => t.getHandlerFunction() === dailyFuncName);
  if (!dailyExists) {
    ScriptApp.newTrigger(dailyFuncName).timeBased().everyDays(1).atHour(0).create();
  }
}
function _getFrontendRedirectBaseUrl() {
  const raw = String(getScriptProperty('FRONTEND_REDIRECT_URL') || '').trim();
  if (!raw) throw new Error('FRONTEND_REDIRECT_URL を Script Properties に設定してください');
  if (/script\.google\.com\/macros\//i.test(raw)) {
    throw new Error('FRONTEND_REDIRECT_URL には ngrok フロントのURLを設定してください');
  }
  return raw;
}
function _getFrontendOAuthCallbackUrl() {
  const base = _getFrontendRedirectBaseUrl().replace(/\/+$/, '');
  return base + '/auth/slack/callback';
}
function isContinueSwitchEnabled() {
  try {
    return getScriptProperty('NEXT_YEAR_CONTINUE_ENABLED') === 'true';
  } catch (e) {
    return false;
  }
}
