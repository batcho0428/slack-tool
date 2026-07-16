function _affiliationDeptCol(slotIndex) {
  return COL.AFFILIATION_START + slotIndex * 2;
}
function _affiliationRoleCol(slotIndex) {
  return COL.AFFILIATION_START + slotIndex * 2 + 1;
}
function _loadMasterMaps(ss) {
  const readMap = (sheetName, keyIndex, valueIndex, statusIndex, notMainIndex) => {
    const out = { byPid: {}, byName: {}, rows: [], statusByPid: {}, notMainByPid: {} };
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return out;
    const statusIdx = (typeof statusIndex === 'number') ? statusIndex : -1;
    const notMainIdx = (typeof notMainIndex === 'number') ? notMainIndex : -1;
    const maxIndex = Math.max(valueIndex, statusIdx, notMainIdx);
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(sh.getLastColumn(), maxIndex + 1)).getValues();
    rows.forEach(r => {
      const pid = String(r[keyIndex] || '').trim();
      const name = String(r[valueIndex] || '').trim();
      if (!pid || !name) return;
      out.byPid[pid] = name;
      out.byName[name] = pid;
      out.statusByPid[pid] = statusIdx >= 0 ? String(r[statusIdx] || '').trim() : '';
      out.notMainByPid[pid] = notMainIdx >= 0 ? (r[notMainIdx] === true || String(r[notMainIdx] || '').toLowerCase() === 'true') : false;
      out.rows.push(r);
    });
    return out;
  };

  const grade = readMap(SHEET_GRADE, 0, 1, 2, 2);
  const org = readMap(SHEET_ORG, 0, 1, 3, 4);
  const dept = readMap(SHEET_DEPT, 0, 1, 3, 4);
  const role = readMap(SHEET_ROLE, 0, 1, null, 3);
  const field = readMap(SHEET_FIELD, 0, 1, 1, 1);

  const deptToOrgPid = {};
  const deptByOrgAndName = {}; // {orgPid: {deptName: deptPid}} to handle multi-org dept duplicates
  dept.rows.forEach(r => {
    const deptPid = String(r[0] || '').trim();
    const orgPid = String(r[2] || '').trim();
    if (deptPid && orgPid) {
      deptToOrgPid[deptPid] = orgPid;
      if (!deptByOrgAndName[orgPid]) deptByOrgAndName[orgPid] = {};
      const deptName = String(r[1] || '').trim();
      if (deptName) deptByOrgAndName[orgPid][deptName] = deptPid;
    }
  });
  dept.deptByOrgAndName = deptByOrgAndName;

  return { grade, org, dept, role, field, deptToOrgPid };
}
function _toPidOrEmpty(value, byName) {
  const s = String(value || '').trim();
  if (!s) return '';
  if (byName && byName[s]) return byName[s];
  return s;
}
function _toLabelOrEmpty(value, byPid) {
  const s = String(value || '').trim();
  if (!s) return '';
  return byPid && byPid[s] ? byPid[s] : s;
}
function _isMasterActive(statusVal) {
  const s = String(statusVal || '').trim();
  return s === '' || s === '0' || s.toLowerCase() === 'false';
}
function _assertActiveSelection(label, pid, master, kind) {
  const key = String(pid || '').trim();
  if (!key) return;
  if (!master || !master.byPid || !master.byPid[key]) {
    throw new Error(label + 'が見つかりません');
  }
  if (!master.statusByPid || !_isMasterActive(master.statusByPid[key])) {
    throw new Error('無効な' + kind + 'は選択できません');
  }
}
function _buildAffiliationCode(orgLabel, deptLabel, masters) {
  const orgPid = _toPidOrEmpty(orgLabel, masters && masters.org ? masters.org.byName : null);
  const deptPid = _toPidOrEmpty(deptLabel, masters && masters.dept ? masters.dept.byName : null);
  if (deptPid) return deptPid;
  if (orgPid) return orgPid;
  return '';
}
function _parseAffiliationCode(code, masters) {
  const raw = String(code || '').trim();
  if (!raw) return { code: '', orgPid: '', org: '', deptPid: '', dept: '' };

  // dept code (4桁想定) が優先
  if (masters && masters.dept && masters.dept.byPid && masters.dept.byPid[raw]) {
    const deptPid = raw;
    const orgPid = (masters.deptToOrgPid && masters.deptToOrgPid[deptPid]) ? masters.deptToOrgPid[deptPid] : '';
    return {
      code: raw,
      orgPid: orgPid,
      org: _toLabelOrEmpty(orgPid, masters.org.byPid),
      deptPid: deptPid,
      dept: _toLabelOrEmpty(deptPid, masters.dept.byPid)
    };
  }

  if (masters && masters.dept && masters.dept.byName && masters.dept.byName[raw]) {
    const deptPid = masters.dept.byName[raw];
    const orgPid = (masters.deptToOrgPid && masters.deptToOrgPid[deptPid]) ? masters.deptToOrgPid[deptPid] : '';
    return {
      code: deptPid,
      orgPid: orgPid,
      org: _toLabelOrEmpty(orgPid, masters.org.byPid),
      deptPid: deptPid,
      dept: _toLabelOrEmpty(deptPid, masters.dept.byPid)
    };
  }

  // org code (2桁想定)
  if (masters && masters.org && masters.org.byPid && masters.org.byPid[raw]) {
    return {
      code: raw,
      orgPid: raw,
      org: _toLabelOrEmpty(raw, masters.org.byPid),
      deptPid: '',
      dept: ''
    };
  }

  if (masters && masters.org && masters.org.byName && masters.org.byName[raw]) {
    const orgPid = masters.org.byName[raw];
    return {
      code: orgPid,
      orgPid: orgPid,
      org: _toLabelOrEmpty(orgPid, masters.org.byPid),
      deptPid: '',
      dept: ''
    };
  }

  // 既存データ互換: 不明値は文字列をそのままdept側に残す
  return { code: raw, orgPid: '', org: '', deptPid: raw, dept: raw };
}
function _buildAffiliationStorageCode(orgValue, deptValue, masters) {
  const orgPid = _toPidOrEmpty(orgValue, masters && masters.org ? masters.org.byName : null);
  const deptPid = _toPidOrEmpty(deptValue, masters && masters.dept ? masters.dept.byName : null);
  if (deptPid) return deptPid;
  if (orgPid) return orgPid;
  return '';
}
