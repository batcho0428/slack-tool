function closeExpiredSurveyCollectionsDaily() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const formsSheet = ss.getSheetByName(SHEET_FORMS);
  if (!formsSheet || formsSheet.getLastRow() < 2) {
    return { success: true, updated: 0, checked: 0, message: 'Formsのデータがありません' };
  }

  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
  const yesterdayKey = Utilities.formatDate(yesterday, tz, 'yyyy-MM-dd');
  const collectingCol = 5;
  const deadlineCol = 8;

  const rowCount = formsSheet.getLastRow() - 1;
  const cols = Math.max(HEADER_FORMS.length, formsSheet.getLastColumn());
  const rows = formsSheet.getRange(2, 1, rowCount, cols).getValues();

  let updated = 0;
  let checked = 0;
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    const collectingRaw = row[collectingCol - 1];
    const collecting = (collectingRaw === true) || (String(collectingRaw || '').toLowerCase() === 'true');
    if (!collecting) continue;

    checked++;
    const deadlineKey = _formatDateOnlyYmd(row[deadlineCol - 1]);
    if (!deadlineKey) continue;

    // 締め切り日が前日以前になっていれば収集中を false にする(トリガーの実行漏れがあっても取りこぼさないよう、前日と一致する場合だけでなくそれ以前も対象にする)。
    if (deadlineKey <= yesterdayKey) {
      formsSheet.getRange(i + 2, collectingCol).setValue(false);
      updated++;
    }
  }

  try {
    const logSheet = ss.getSheetByName(SHEET_LOGS) || ss.insertSheet(SHEET_LOGS);
    logSheet.appendRow([new Date(), 'system', '-', 'Survey Collecting Auto Close', 'updated=' + updated + ', checked=' + checked + ', targetDate=' + yesterdayKey]);
  } catch (e) {}

  return { success: true, updated: updated, checked: checked, targetDate: yesterdayKey };
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
  const masters = _loadMasterMaps(ss);
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
      const aff = _parseAffiliationCode(r[3], masters);
      const collectingRaw = r[4];
      const collecting = (collectingRaw === true) || (String(collectingRaw || '').toLowerCase() === 'true');
      const scoreName = String(r[5] || '').trim() || null;
      const scoreUnit = String(r[6] || '').trim() || null;
      const deadlineDate = _formatDateOnlyYmd(r[7]) || null;
      const sid = _extractSpreadsheetId(spreadRef) || _extractSpreadsheetId(formUrl);
      if (!sid) {
        out.push({ title: title, spreadsheetId: null, spreadsheetUrl: null, formUrl: formUrl || null, inChargeOrg: aff.orgPid || '', inChargeDept: aff.deptPid || '', inChargeCode: aff.code || '', inChargeOrgLabel: aff.org || '', inChargeDeptLabel: aff.dept || aff.org || '', collecting: collecting, scoreName: scoreName, scoreUnit: scoreUnit, deadlineDate: deadlineDate, userLatestRowIndex: null, available: false, latestResponseDate: null, latestScore: null });
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
        inChargeOrg: aff.orgPid || '',
        inChargeDept: aff.deptPid || '',
        inChargeCode: aff.code || '',
        inChargeOrgLabel: aff.org || '',
        inChargeDeptLabel: aff.dept || aff.org || '',
        collecting: collecting,
        scoreName: scoreName,
        scoreUnit: scoreUnit,
        deadlineDate: deadlineDate,
        userLatestRowIndex: userLatestRowIndex,
        available: available,
        latestResponseDate: available ? latestResponseDate : (latestDate || null),
        latestScore: available ? latestScore : null,
        latestScoreFormatted: available ? latestScoreFormatted : null
      });
    } catch (e) { out.push({ title: String(r[2]||r[0]||''), spreadsheetId: null, spreadsheetUrl: null, userLatestRowIndex: null, available: false, latestResponseDate: null, latestScore: null }); }
  });
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
    const masters = _loadMasterMaps(ss);
    const cols = Math.max(HEADER_FORMS.length, formsSheet.getLastColumn());
    const rows = formsSheet.getRange(2, 1, lastRow - 1, cols).getValues();
    const items = rows.map((r, i) => {
        const aff = _parseAffiliationCode(r[3], masters);
      const collectingRaw = r[4];
      const collecting = (collectingRaw === true) || (String(collectingRaw || '').toLowerCase() === 'true');
      return {
        rowIndex: i + 2,
        spreadsheetRef: String(r[0] || '').trim(),
        formUrl: String(r[1] || '').trim(),
        title: String(r[2] || '').trim(),
          inChargeOrg: aff.orgPid || '',
          inChargeDept: aff.deptPid || '',
          inChargeCode: aff.code || '',
          inChargeOrgLabel: aff.org || '',
          inChargeDeptLabel: aff.dept || aff.org || '',
        collecting: collecting,
        scoreName: String(r[5] || '').trim(),
        scoreUnit: String(r[6] || '').trim(),
        deadlineDate: _formatDateOnlyYmd(r[7])
      };
    });
    return { success: true, items: items };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
function _buildReminderUsers(ss, masters) {
  const usersSheet = ss.getSheetByName(SHEET_USERS);
  if (!usersSheet) return [];
  const rows = usersSheet.getDataRange().getValues();
  const users = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const email = String(row[COL.EMAIL] || '').trim().toLowerCase();
    const name = String(row[COL.NAME_JP] || '').trim();
    if (!email || !name) continue;

    const affiliations = [];
    const orgLabels = [];
    const deptLabels = [];
    const roleLabels = [];
    const deptTexts = [];
    for (let k = 0; k < AFFILIATION_SLOTS; k++) {
      const affCode = String(row[_affiliationDeptCol(k)] || '').trim();
      const aff = _parseAffiliationCode(affCode, masters);
      const rolePid = String(row[_affiliationRoleCol(k)] || '').trim();
      const roleLabel = _toLabelOrEmpty(rolePid, masters.role.byPid);
      if (!aff.org && !aff.dept && !roleLabel) continue;
      affiliations.push({ org: aff.org || '', dept: aff.dept || '', role: roleLabel || '' });
      if (aff.org) orgLabels.push(aff.org);
      if (aff.dept) deptLabels.push(aff.dept);
      if (roleLabel) roleLabels.push(roleLabel);
      const affLabel = aff.dept ? (aff.org ? (aff.org + '/' + aff.dept) : aff.dept) : (aff.org || '');
      deptTexts.push([affLabel, roleLabel].filter(Boolean).join(' '));
    }

    const retired = row[COL.RETIRED] === true || row[COL.RETIRED] === 'TRUE';
    users.push({
      name: name,
      email: email,
      studentId: String(row[COL.STUDENT_ID] || '').trim(),
      grade: _toLabelOrEmpty(row[COL.GRADE], masters.grade.byPid),
      field: _toLabelOrEmpty(row[COL.FIELD], masters.field.byPid),
      retired: retired,
      affiliations: affiliations,
      org: Array.from(new Set(orgLabels)),
      department: Array.from(new Set(deptLabels)),
      role: Array.from(new Set(roleLabels)),
      departmentText: deptTexts.join(', ') || '所属なし',
      mainOrg: affiliations.length > 0 ? String(affiliations[0].org || '') : '',
      mainDept: affiliations.length > 0 ? String(affiliations[0].dept || '') : ''
    });
  }

  return users;
}
function _requireAdminLogin(sessionToken) {
  const login = getLoginUser(sessionToken);
  if (!login || login.status !== 'authorized') {
    return { ok: false, message: '認証されていません' };
  }
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const usersSheet = ss.getSheetByName(SHEET_USERS);
  const usersData = usersSheet ? usersSheet.getDataRange().getValues() : [];
  const loginEmail = String(login.user && login.user.email ? login.user.email : '').trim().toLowerCase();
  const loginRow = usersData.find(r => String(r[COL.EMAIL] || '').trim().toLowerCase() === loginEmail);
  const isAdmin = !!(loginRow && (loginRow[COL.ADMIN] === 'TRUE' || loginRow[COL.ADMIN] === true));
  if (!isAdmin) {
    return { ok: false, message: '権限がありません' };
  }
  return { ok: true, login: login, ss: ss };
}

function _computeReminderStatus(ss, masters, surveyRowIndices) {
  const formsSheet = ss.getSheetByName(SHEET_FORMS);
  if (!formsSheet || formsSheet.getLastRow() < 2) {
    return { surveys: [], users: _buildReminderUsers(ss, masters), unansweredByEmail: {} };
  }

  const selectedRows = {};
  (surveyRowIndices || []).forEach(v => {
    const n = Number(v);
    if (!isNaN(n) && n >= 2) selectedRows[n] = true;
  });

  const cols = Math.max(HEADER_FORMS.length, formsSheet.getLastColumn());
  const formRows = formsSheet.getRange(2, 1, formsSheet.getLastRow() - 1, cols).getValues();
  const selectedSurveys = [];

  for (let i = 0; i < formRows.length; i++) {
    const rowIndex = i + 2;
    if (Object.keys(selectedRows).length > 0 && !selectedRows[rowIndex]) continue;
    const r = formRows[i] || [];
    const spreadRef = String(r[0] || '').trim();
    const formUrl = String(r[1] || '').trim();
    const title = String(r[2] || '').trim() || spreadRef || formUrl || ('アンケート' + rowIndex);
    const sid = _extractSpreadsheetId(spreadRef) || _extractSpreadsheetId(formUrl);
    if (!sid) continue;

    const collectingRaw = r[4];
    const collecting = (collectingRaw === true) || (String(collectingRaw || '').toLowerCase() === 'true');
    const deadlineDate = _formatDateOnlyYmd(r[7]);
    const surveyInfo = {
      rowIndex: rowIndex,
      spreadsheetId: sid,
      title: title,
      formUrl: formUrl,
      collecting: collecting,
      deadlineDate: deadlineDate,
      respondedEmails: {},
      respondedStudentIds: {},
      error: ''
    };

    try {
      const targetSs = SpreadsheetApp.openById(sid);
      const sh = targetSs.getSheets()[0];
      const lastCol = Math.max(1, sh.getLastColumn());
      const headerRow = _detectHeaderRow(sh, lastCol);
      const headers = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0] || [];
      const emailIdx = _findHeaderIndex(headers, ['メール', 'メールアドレス', '^email$','^e-mail$']);
      const sidIdx = _findHeaderIndex(headers, ['学籍番号', 'student id', 'studentid', '学籍']);
      const dataCount = Math.max(0, sh.getLastRow() - headerRow);
      if (dataCount > 0 && (emailIdx >= 0 || sidIdx >= 0)) {
        const data = sh.getRange(headerRow + 1, 1, dataCount, lastCol).getValues();
        for (let rIdx = 0; rIdx < data.length; rIdx++) {
          const row = data[rIdx] || [];
          if (emailIdx >= 0) {
            const em = String(row[emailIdx] || '').trim().toLowerCase();
            if (em) surveyInfo.respondedEmails[em] = true;
          }
          if (sidIdx >= 0) {
            const st = String(row[sidIdx] || '').trim();
            if (st) surveyInfo.respondedStudentIds[st] = true;
          }
        }
      }
    } catch (e) {
      surveyInfo.error = String(e && e.message ? e.message : e);
    }

    selectedSurveys.push(surveyInfo);
  }

  const users = _buildReminderUsers(ss, masters);
  const unansweredByEmail = {};

  users.forEach(u => {
    const email = String(u.email || '').trim().toLowerCase();
    if (!email) return;
    const pending = [];
    selectedSurveys.forEach(s => {
      if (s.error) return;
      if (!s.collecting) return;
      const byEmail = !!s.respondedEmails[email];
      const byStudentId = !!(u.studentId && s.respondedStudentIds[u.studentId]);
      if (!byEmail && !byStudentId) {
        pending.push({
          rowIndex: s.rowIndex,
          title: s.title,
          formUrl: s.formUrl || '',
          deadlineDate: s.deadlineDate || ''
        });
      }
    });
    unansweredByEmail[email] = pending;
  });

  return {
    surveys: selectedSurveys.map(s => ({
      rowIndex: s.rowIndex,
      spreadsheetId: s.spreadsheetId,
      title: s.title,
      formUrl: s.formUrl,
      collecting: s.collecting,
      deadlineDate: s.deadlineDate || '',
      error: s.error || ''
    })),
    users: users,
    unansweredByEmail: unansweredByEmail
  };
}
function collectSurveyReminderStatus(sessionToken, surveyRowIndices) {
  const auth = _requireAdminLogin(sessionToken);
  if (!auth.ok) return { success: false, message: auth.message };

  const masters = _loadMasterMaps(auth.ss);
  const result = _computeReminderStatus(auth.ss, masters, surveyRowIndices);
  return { success: true, surveys: result.surveys, users: result.users, unansweredByEmail: result.unansweredByEmail };
}
function _buildSurveyReminderText(mention, unansweredSurveys) {
  const list = Array.isArray(unansweredSurveys) ? unansweredSurveys : [];
  if (list.length === 0) {
    return '回答が必要なアンケートはありません。\nアンケート回答へのご協力ありがとうございました。';
  }

  let text = mention + ' さん\nあなたの未回答のアンケートをお知らせします\n期限までに回答へのご協力お願いします\n\n';
  list.forEach(item => {
    const title = String(item && item.title ? item.title : 'アンケート').trim();
    const deadlineLabel = _formatDateOnlyYmd(item && item.deadlineDate ? item.deadlineDate : '') || '期限未設定';
    const line = '～' + deadlineLabel + ' ' + title;
    text += line + '\n';
    const url = String(item && item.formUrl ? item.formUrl : '').trim();
    if (url) text += url + '\n';
    text += '\n';
  });
  return text.trim();
}
function sendSurveyReminderDMs(sessionToken, payload) {
  const auth = _requireAdminLogin(sessionToken);
  if (!auth.ok) {
    return { success: 0, failed: [{ email: '', error: auth.message }] };
  }

  const senderEmail = String(auth.login.user && auth.login.user.email ? auth.login.user.email : '').trim().toLowerCase();
  const requestedRecipients = (payload && Array.isArray(payload.recipients)) ? payload.recipients : [];
  const botToken = _normalizeSlackCredential(getScriptProperty('SLACK_BOT_TOKEN'));
  if (!botToken) {
    return { success: 0, failed: requestedRecipients.map(r => ({ email: String(r && r.email ? r.email : ''), error: 'SLACK_BOT_TOKEN が設定されていません' })) };
  }

  // 送信内容(未回答アンケート一覧)はクライアント入力を信頼せず、サーバー側で選択されたアンケートの実データから再計算する。
  const masters = _loadMasterMaps(auth.ss);
  const reminderStatus = _computeReminderStatus(auth.ss, masters, payload && payload.surveyRowIndices);
  const registeredEmails = {};
  reminderStatus.users.forEach(u => {
    const email = String(u.email || '').trim().toLowerCase();
    if (email) registeredEmails[email] = true;
  });

  const logSheet = auth.ss.getSheetByName(SHEET_LOGS) || auth.ss.insertSheet(SHEET_LOGS);
  try {
    logSheet.appendRow([new Date(), senderEmail, '-', 'Survey Remind Trigger', 'count=' + requestedRecipients.length]);
  } catch (e) {}

  let successCount = 0;
  const failedList = [];

  requestedRecipients.forEach((r) => {
    const recipientEmail = String(r && r.email ? r.email : '').trim().toLowerCase();
    const unanswered = reminderStatus.unansweredByEmail[recipientEmail] || [];
    try {
      if (!recipientEmail) throw new Error('メールアドレスが不正です');
      if (!registeredEmails[recipientEmail]) throw new Error('登録されていないメールアドレスです');
      if (unanswered.length === 0) throw new Error('未回答のアンケートがありません');
      const uid = getSlackID(botToken, recipientEmail);
      if (!uid) throw new Error('Slackアカウントなし');
      const text = _buildSurveyReminderText('<@' + uid + '>', unanswered);
      const res = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
        method: 'post',
        contentType: 'application/json',
        headers: { 'Authorization': 'Bearer ' + botToken },
        payload: JSON.stringify({ channel: uid, text: text, unfurl_links: false, unfurl_media: false }),
        muteHttpExceptions: true
      });
      const json = JSON.parse(res.getContentText());
      if (!json.ok) throw new Error(json.error || 'Unknown Error');
      successCount++;
      logSheet.appendRow([new Date(), senderEmail, recipientEmail, 'Survey Remind Success', 'unanswered=' + unanswered.length]);
    } catch (e) {
      const msg = String(e && e.message ? e.message : e);
      failedList.push({ email: recipientEmail, name: String(r && r.name ? r.name : ''), error: msg });
      try {
        logSheet.appendRow([new Date(), senderEmail, recipientEmail, 'Survey Remind Failed', msg]);
      } catch (e2) {}
    }
    Utilities.sleep(1200);
  });

  return { success: successCount, failed: failedList };
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
    const masters = _loadMasterMaps(SpreadsheetApp.openById(getSpreadsheetId()));
    const row = new Array(HEADER_FORMS.length).fill('');
    row[0] = spreadsheetRef;
    row[1] = formUrl;
    row[2] = String(data.title || '').trim();
    row[3] = _buildAffiliationStorageCode(data.inChargeOrg, data.inChargeDept, masters);
    row[4] = data.collecting ? true : false;
    row[5] = String(data.scoreName || '').trim();
    row[6] = String(data.scoreUnit || '').trim();
    row[7] = _parseDateOnlyValue(data.deadlineDate || '');

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
    try {
      const lastRow = formsSheet.getLastRow();
      if (lastRow >= 2) formsSheet.getRange(2, 8, Math.max(1, lastRow - 1), 1).setNumberFormat('yyyy/mm/dd');
    } catch (e) {}
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
            scoreName = String(metaRows[i][5] || '').trim() || null;
            scoreUnit = String(metaRows[i][6] || '').trim() || null;
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
          scoreName = String(metaRows[i][5] || '').trim() || null;
          scoreUnit = String(metaRows[i][6] || '').trim() || null;
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
