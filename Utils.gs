function _normalizeSlackCredential(value) {
  const s = String(value || '').trim();
  if (!s) return '';
  if ((s.startsWith('"') && s.endsWith('"')) || (s.startsWith("'") && s.endsWith("'"))) {
    return s.slice(1, -1).trim();
  }
  return s;
}
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
function _escapeHtmlAttribute(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}
function _buildRedirectHtml(targetUrl) {
  const safeTarget = String(targetUrl || '').trim();
  const metaRefreshTarget = _escapeHtmlAttribute(safeTarget);
  const jsTargetLiteral = JSON.stringify(safeTarget);
  return HtmlService.createHtmlOutput(`
    <html>
      <head>
        <meta charset="UTF-8" />
        <meta http-equiv="refresh" content="0; url=${metaRefreshTarget}" />
        <script>
          window.top.location.replace(${jsTargetLiteral});
        </script>
      </head>
      <body>Redirecting...</body>
    </html>
  `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function _nextNumericPid(existingPids, width) {
  let max = 0;
  existingPids.forEach(p => {
    const n = parseInt(String(p || '').replace(/[^0-9]/g, ''), 10);
    if (!isNaN(n) && n > max) max = n;
  });
  const next = max + 1;
  return Utilities.formatString('%0' + width + 'd', next);
}
function _nextDataRowByPidColumn(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  if (lastRow < 2) return 2;
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let lastDataRow = 1;
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim()) lastDataRow = i + 2;
  }
  return lastDataRow + 1;
}
function _applyCheckboxColumn(sheet, colIndex1Based) {
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const numRows = Math.max(1, lastRow - 1);
  const range = sheet.getRange(2, colIndex1Based, numRows, 1);
  try {
    const values = range.getValues();
    const normalized = values.map(row => {
      const value = row[0];
      if (value === true || value === 'TRUE' || String(value).toLowerCase() === 'true') return [true];
      if (value === false || value === 'FALSE' || String(value).toLowerCase() === 'false') return [false];
      return [false];
    });
    range.setValues(normalized);
  } catch (e) {}
  try {
    range.insertCheckboxes();
  } catch (e) {
    try {
      const rule = SpreadsheetApp.newDataValidation().requireValueInList(['TRUE', 'FALSE']).setAllowInvalid(true).build();
      range.setDataValidation(rule);
    } catch (ee) {}
  }
}
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
function _parseDateOnlyValue(value) {
  if (!value && value !== 0) return null;
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return null;
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  const raw = String(value || '').trim();
  if (!raw) return null;
  const normalized = raw.replace(/\./g, '/').replace(/-/g, '/');
  const m = normalized.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if (m) {
    const y = Number(m[1]);
    const mo = Number(m[2]) - 1;
    const d = Number(m[3]);
    const dt = new Date(y, mo, d);
    if (!isNaN(dt.getTime())) return dt;
  }

  const parsed = new Date(raw);
  if (isNaN(parsed.getTime())) return null;
  return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}
function _formatDateOnlyYmd(value) {
  const d = _parseDateOnlyValue(value);
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function _extractFormId(formUrl) {
  if (!formUrl) return null;
  try {
    // support URLs like /forms/d/ID/ and /forms/d/e/ID/
    const m = String(formUrl).match(/\/forms\/d\/(?:e\/)?([-_0-9A-Za-z]+)/);
    if (m && m[1]) return m[1];
  } catch (e) {}
  return null;
}
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
