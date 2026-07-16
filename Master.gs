function listAffiliationMasters(sessionToken) {
  _requireAdmin(sessionToken);
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const masters = _loadMasterMaps(ss);
  const isActive = (statusVal) => {
    const s = String(statusVal || '').trim();
    return s === '' || s === '0' || s.toLowerCase() === 'false';
  };

  const orgs = Object.keys(masters.org.byPid).map(pid => ({
    pid: pid,
    org: masters.org.byPid[pid],
    status: String(masters.org.statusByPid[pid] || ''),
    not_main_org: !!masters.org.notMainByPid[pid],
    gen: (() => {
      const row = masters.org.rows.find(r => String(r[0] || '').trim() === pid);
      return row ? row[2] : '';
    })(),
    active: isActive(masters.org.statusByPid[pid])
  })).sort((a, b) => a.pid.localeCompare(b.pid));

  const depts = Object.keys(masters.dept.byPid).map(pid => {
    const orgPid = masters.deptToOrgPid[pid] || '';
    return {
      pid: pid,
      dept: masters.dept.byPid[pid],
      orgPid: orgPid,
      org: _toLabelOrEmpty(orgPid, masters.org.byPid),
      status: String(masters.dept.statusByPid[pid] || ''),
      not_main_dept: !!masters.dept.notMainByPid[pid],
      active: isActive(masters.dept.statusByPid[pid])
    };
  }).sort((a, b) => a.pid.localeCompare(b.pid));

  const roles = Object.keys(masters.role.byPid).map(pid => ({
    pid: pid,
    role: masters.role.byPid[pid],
    gen: (() => {
      const row = masters.role.rows.find(r => String(r[0] || '').trim() === pid);
      return row ? row[2] : '';
    })(),
    status: String(masters.role.statusByPid[pid] || ''),
    not_main_role: !!masters.role.notMainByPid[pid],
    active: isActive(masters.role.statusByPid[pid])
  })).sort((a, b) => a.pid.localeCompare(b.pid));

  return { orgs: orgs, depts: depts, roles: roles };
}
function saveOrgMaster(sessionToken, payload) {
  _requireAdmin(sessionToken);
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const sh = ss.getSheetByName(SHEET_ORG);
  if (!sh) throw new Error('org シートが見つかりません');
  const data = sh.getDataRange().getValues();
  const pidIn = String((payload && payload.pid) || '').trim();
  const orgName = String((payload && payload.org) || '').trim();
  if (!orgName) throw new Error('局名は必須です');
  const genRaw = (payload && typeof payload.gen !== 'undefined') ? String(payload.gen).trim() : '';
  const gen = genRaw === '' ? '' : Number(genRaw);
  const status = String((payload && payload.status) || '').trim();
  const notMain = !!(payload && payload.not_main_org);

  // If not_main_org is true, gen is required
  if (notMain && (genRaw === '' || isNaN(gen))) {
    throw new Error('not_main_org が True の場合、gen は必須です');
  }

  if (pidIn) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === pidIn) {
        sh.getRange(i + 1, 1, 1, 5).setValues([[pidIn, orgName, gen, status, notMain]]);
        _applyCheckboxColumn(sh, 5);
        return { success: true, pid: pidIn };
      }
    }
  }

  let nextPid = '';
  if (notMain) {
    // gen is required and provides first two digits; append single-digit sequence
    const genStr = Utilities.formatString('%02d', gen);
    const siblings = data.slice(1).map(r => String(r[0] || '').trim()).filter(pid => pid.startsWith(genStr) && pid.length === 3);
    const seqs = siblings.map(pid => parseInt(pid.slice(2), 10)).filter(n => !isNaN(n));
    const maxSeq = seqs.length ? Math.max.apply(null, seqs) : 0;
    const nextSeq = maxSeq + 1;
    nextPid = genStr + String(nextSeq);
  } else {
    nextPid = _nextNumericPid(data.slice(1).map(r => r[0]), 2);
  }
  const nextRow = _nextDataRowByPidColumn(sh);
  sh.getRange(nextRow, 1, 1, 5).setValues([[nextPid, orgName, (notMain ? gen : ''), status, notMain]]);
  _applyCheckboxColumn(sh, 5);
  return { success: true, pid: nextPid };
}
function saveDeptMaster(sessionToken, payload) {
  _requireAdmin(sessionToken);
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const sh = ss.getSheetByName(SHEET_DEPT);
  if (!sh) throw new Error('dept シートが見つかりません');
  const masters = _loadMasterMaps(ss);
  const data = sh.getDataRange().getValues();
  const pidIn = String((payload && payload.pid) || '').trim();
  const deptName = String((payload && payload.dept) || '').trim();
  const orgInput = String((payload && payload.orgPid) || '').trim();
  let orgPid = '';
  if (orgInput) {
    if (masters.org.byPid && masters.org.byPid[orgInput]) orgPid = orgInput;
    else if (masters.org.byName && masters.org.byName[orgInput]) orgPid = masters.org.byName[orgInput];
    else orgPid = orgInput;
  }
  if (!deptName) throw new Error('部門名は必須です');
  if (!orgPid) throw new Error('所属局は必須です');
  const status = String((payload && payload.status) || '').trim();
  const notMain = !!(payload && payload.not_main_dept);

  _assertActiveSelection('所属局', orgPid, masters.org, '所属局');

  if (pidIn) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === pidIn) {
        sh.getRange(i + 1, 1, 1, 5).setValues([[pidIn, deptName, orgPid, status, notMain]]);
        _applyCheckboxColumn(sh, 5);
        return { success: true, pid: pidIn };
      }
    }
  }

  const siblings = data.slice(1).map(r => String(r[0] || '').trim()).filter(pid => pid.startsWith(orgPid) && pid.length >= 4);
  const seq = siblings.map(pid => parseInt(pid.slice(-2), 10)).filter(n => !isNaN(n));
  const maxSeq = seq.length ? Math.max.apply(null, seq) : 0;
  const nextPid = orgPid + Utilities.formatString('%02d', maxSeq + 1);
  const nextRow = _nextDataRowByPidColumn(sh);
  sh.getRange(nextRow, 1, 1, 5).setValues([[nextPid, deptName, orgPid, status, notMain]]);
  _applyCheckboxColumn(sh, 5);
  return { success: true, pid: nextPid };
}
function saveRoleMaster(sessionToken, payload) {
  _requireAdmin(sessionToken);
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const sh = ss.getSheetByName(SHEET_ROLE);
  if (!sh) throw new Error('role シートが見つかりません');
  const data = sh.getDataRange().getValues();
  const pidIn = String((payload && payload.pid) || '').trim();
  const roleName = String((payload && payload.role) || '').trim();
  if (!roleName) throw new Error('役職名は必須です');
  const genRaw = (payload && typeof payload.gen !== 'undefined') ? String(payload.gen).trim() : '';
  const gen = genRaw === '' ? '' : Number(genRaw);
  const notMain = !!(payload && payload.not_main_role);

  if (pidIn) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === pidIn) {
        sh.getRange(i + 1, 1, 1, 4).setValues([[pidIn, roleName, gen, notMain]]);
        _applyCheckboxColumn(sh, 4);
        return { success: true, pid: pidIn };
      }
    }
  }

  const nextPid = _nextNumericPid(data.slice(1).map(r => r[0]), 2);
  const nextRow = _nextDataRowByPidColumn(sh);
  sh.getRange(nextRow, 1, 1, 4).setValues([[nextPid, roleName, gen, notMain]]);
  _applyCheckboxColumn(sh, 4);
  return { success: true, pid: nextPid };
}
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
