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
function doGet(e) {
  try {
    const baseUrl = _getFrontendRedirectBaseUrl();
    const p = (e && e.parameter) ? e.parameter : {};
    if (p.code || p.error) {
      const callbackUrl = _getFrontendOAuthCallbackUrl();
      const query = [];
      if (p.code) query.push('code=' + encodeURIComponent(String(p.code)));
      if (p.state) query.push('state=' + encodeURIComponent(String(p.state)));
      if (p.error) query.push('error=' + encodeURIComponent(String(p.error)));
      if (p.error_description) query.push('error_description=' + encodeURIComponent(String(p.error_description)));
      const target = callbackUrl + (query.length ? ('?' + query.join('&')) : '');
      return _buildRedirectHtml(target);
    }

    return _buildRedirectHtml(baseUrl);
  } catch (err) {
    return HtmlService
      .createHtmlOutput('FRONTEND_REDIRECT_URL を ngrok フロントURLで設定してください。')
      .setTitle(APP_NAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
function getScriptUrl() { return ScriptApp.getService().getUrl(); }
function handleSpreadsheetEdit(e) {
  return;
}
function _generateCollectionId() {
  return 'COL' + String(Date.now());
}
