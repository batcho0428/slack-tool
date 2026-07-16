function getLoginUser(sessionToken) {
  try {
    if (!sessionToken) return { status: 'guest' };

    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const masters = _loadMasterMaps(ss);
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
function getUserProfile(sessionToken, targetEmail) {
  const login = getLoginUser(sessionToken);
  if (login.status !== 'authorized') throw new Error("認証されていません");

  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const masters = _loadMasterMaps(ss);
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

  const affiliations = [];
  for (let k = 0; k < AFFILIATION_SLOTS; k++) {
    const affiliationCode = String(row[_affiliationDeptCol(k)] || '').trim();
    const aff = _parseAffiliationCode(affiliationCode, masters);
    const rolePid = String(row[_affiliationRoleCol(k)] || '').trim();
    const role = _toLabelOrEmpty(rolePid, masters.role.byPid);
    affiliations.push({ org: aff.org, dept: aff.dept, role: role, affiliationCode: aff.code });
  }

  return {
    name: row[COL.NAME_JP],
    nameEn: row[COL.NAME_EN],
    email: row[COL.EMAIL],
    studentId: row[COL.STUDENT_ID],
    grade: _toLabelOrEmpty(row[COL.GRADE], masters.grade.byPid),
    field: _toLabelOrEmpty(row[COL.FIELD], masters.field.byPid),
    phone: row[COL.PHONE],
    birthday: birthdayVal,
    almaMater: row[COL.ALMA_MATER],
    carOwner: row[COL.CAR_OWNER] === 'TRUE' || row[COL.CAR_OWNER] === true,
    retired: row[COL.RETIRED] === 'TRUE' || row[COL.RETIRED] === true,
    continueNext: row[COL.CONTINUE_NEXT] === 'TRUE' || row[COL.CONTINUE_NEXT] === true,
    orgs: affiliations,
    canEditNameEmail: isAdmin,
    isAdmin: viewedIsAdmin
  };
}
function updateUserProfile(sessionToken, formData, targetEmail) {
  const login = getLoginUser(sessionToken);
  if (login.status !== 'authorized') throw new Error("認証されていません");

  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const masters = _loadMasterMaps(ss);
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
  sheet.getRange(rowIndex, COL.GRADE + 1).setValue(_toPidOrEmpty(formData.grade, masters.grade.byName));
  sheet.getRange(rowIndex, COL.FIELD + 1).setValue(_toPidOrEmpty(formData.field, masters.field.byName));
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

  // 所属情報 (10セット)
  if (formData.orgs && Array.isArray(formData.orgs)) {
    for (let k = 0; k < AFFILIATION_SLOTS; k++) {
      if (k < formData.orgs.length) {
        const o = formData.orgs[k];
        const orgPid = _toPidOrEmpty(o.org, masters.org.byName);
        const deptPid = _toPidOrEmpty(o.dept, masters.dept.byName);
        const rolePid = _toPidOrEmpty(o.role, masters.role.byName);
        if (orgPid) _assertActiveSelection('所属局', orgPid, masters.org, '所属局');
        if (deptPid) _assertActiveSelection('所属部門', deptPid, masters.dept, '所属部門');
        if (rolePid) _assertActiveSelection('役職', rolePid, masters.role, '所属役職');
        const affCode = _buildAffiliationCode(o.org, o.dept, masters);
        sheet.getRange(rowIndex, _affiliationDeptCol(k) + 1).setValue(affCode);
        sheet.getRange(rowIndex, _affiliationRoleCol(k) + 1).setValue(rolePid);
      } else {
        sheet.getRange(rowIndex, _affiliationDeptCol(k) + 1).setValue('');
        sheet.getRange(rowIndex, _affiliationRoleCol(k) + 1).setValue('');
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

  const masters = _loadMasterMaps(ss);

  // 行データ作成
  const row = new Array(HEADER_USERS.length).fill('');
  row[COL.NAME_JP] = name;
  row[COL.NAME_EN] = userObj.nameEn || '';
  row[COL.STUDENT_ID] = userObj.studentId || '';
  row[COL.GRADE] = _toPidOrEmpty(userObj.grade, masters.grade.byName);
  row[COL.FIELD] = _toPidOrEmpty(userObj.field, masters.field.byName);
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

  // 所属 (10セット)
  if (userObj.orgs && Array.isArray(userObj.orgs)) {
    for (let k = 0; k < AFFILIATION_SLOTS; k++) {
      if (k < userObj.orgs.length) {
        const o = userObj.orgs[k] || {};
        const orgPid = _toPidOrEmpty(o.org, masters.org.byName);
        const deptPid = _toPidOrEmpty(o.dept, masters.dept.byName);
        const rolePid = _toPidOrEmpty(o.role, masters.role.byName);
        if (orgPid) _assertActiveSelection('所属局', orgPid, masters.org, '所属局');
        if (deptPid) _assertActiveSelection('所属部門', deptPid, masters.dept, '所属部門');
        if (rolePid) _assertActiveSelection('役職', rolePid, masters.role, '所属役職');
        row[_affiliationDeptCol(k)] = _buildAffiliationCode(o.org, o.dept, masters);
        row[_affiliationRoleCol(k)] = rolePid;
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
function getSearchOptions() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const options = { grades: [], fields: [], roles: [], orgs: [], deptMaster: [], orgMaster: [], roleMaster: [] };
  const masters = _loadMasterMaps(ss);

  const sortByPid = (a, b) => {
    const sa = String(a || '').trim();
    const sb = String(b || '').trim();
    const na = /^\d+$/.test(sa) ? parseInt(sa, 10) : null;
    const nb = /^\d+$/.test(sb) ? parseInt(sb, 10) : null;
    if (na !== null && nb !== null) {
      if (na !== nb) return na - nb;
      if (sa.length !== sb.length) return sa.length - sb.length;
      return sa.localeCompare(sb);
    }
    if (na !== null && nb === null) return -1;
    if (na === null && nb !== null) return 1;
    return sa.localeCompare(sb);
  };
  const sortByNotMainThenPid = (a, b) => {
    const an = !!a.notMain;
    const bn = !!b.notMain;
    if (an !== bn) return an ? 1 : -1; // FALSE -> TRUE
    return sortByPid(a.pid, b.pid);
  };

  const gradePidList = Object.keys(masters.grade.byPid).sort(sortByPid);
  const fieldPidList = Object.keys(masters.field.byPid).sort(sortByPid);
  options.grades = gradePidList.map(pid => masters.grade.byPid[pid]);
  options.fields = fieldPidList.map(pid => masters.field.byPid[pid]);
  const isActive = (statusVal) => {
    const s = String(statusVal || '').trim();
    return s === '' || s === '0' || s.toLowerCase() === 'false';
  };

  const orgMasterSorted = Object.keys(masters.org.byPid)
    .filter(pid => isActive(masters.org.statusByPid[pid]))
    .map(pid => ({ pid: pid, org: masters.org.byPid[pid], notMain: !!masters.org.notMainByPid[pid] }))
    .sort(sortByNotMainThenPid);

  const roleMasterSorted = Object.keys(masters.role.byPid)
    .filter(pid => isActive(masters.role.statusByPid[pid]))
    .map(pid => ({ pid: pid, role: masters.role.byPid[pid], notMain: !!masters.role.notMainByPid[pid] }))
    .sort(sortByNotMainThenPid);

  const deptMasterSorted = Object.keys(masters.dept.byPid)
    .filter(pid => isActive(masters.dept.statusByPid[pid]))
    .map(deptPid => {
      const orgPid = masters.deptToOrgPid[deptPid] || '';
      const orgName = masters.org.byPid[orgPid] || orgPid;
      return { org: orgName, dept: masters.dept.byPid[deptPid], pid: deptPid, orgPid: orgPid, notMain: !!masters.dept.notMainByPid[deptPid] };
    })
    .sort(sortByNotMainThenPid);

  options.orgMaster = orgMasterSorted;
  options.roleMaster = roleMasterSorted;
  options.deptMaster = deptMasterSorted;

  options.orgs = orgMasterSorted.map(v => v.org);
  options.roles = roleMasterSorted.map(v => v.role);

  const uniqueInOrder = function(arr) {
    const seen = {};
    const out = [];
    arr.forEach(v => {
      const s = String(v || '').trim();
      if (!s || seen[s]) return;
      seen[s] = true;
      out.push(s);
    });
    return out;
  };
  const seenDept = {};
  const dedupDeptMaster = [];
  for (let i = 0; i < options.deptMaster.length; i++) {
    const org = String(options.deptMaster[i].org || '').trim();
    const dept = String(options.deptMaster[i].dept || '').trim();
    if (!org || !dept) continue;
    const key = org + '||' + dept;
    if (seenDept[key]) continue;
    seenDept[key] = true;
    dedupDeptMaster.push(options.deptMaster[i]);
  }

  options.grades = uniqueInOrder(options.grades);
  options.fields = uniqueInOrder(options.fields);
  options.roles = uniqueInOrder(options.roles);
  options.orgs = uniqueInOrder(options.orgs);
  options.deptMaster = dedupDeptMaster;

  Logger.log('getSearchOptions result: ' + JSON.stringify(options));
  return options;
}
function searchRecipients(criteria) {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const sheet = ss.getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  const masters = _loadMasterMaps(ss);
  const results = [];
  const normalize = function(value) {
    return String(value || '').trim().replace(/\s+/g, ' ').toLowerCase();
  };

  const q = normalize(criteria.query || '');
  const filterGrade = normalize(criteria.grade || '');
  const filterField = normalize(criteria.field || '');
  const filterOrg = normalize(criteria.org || '');
  const filterDept = normalize(criteria.dept || '');
  const filterRole = normalize(criteria.role || '');
  const filterStatus = criteria.status || "active"; // 'active', 'retired', 'all'

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const nameJp = String(row[COL.NAME_JP] || '').trim();
    const nameEn = String(row[COL.NAME_EN] || '').trim();
    const studentId = row[COL.STUDENT_ID] || "";
    const grade = _toLabelOrEmpty(row[COL.GRADE], masters.grade.byPid);
    const field = _toLabelOrEmpty(row[COL.FIELD], masters.field.byPid);
    const email = String(row[COL.EMAIL] || '').trim();
    const almaMater = row[COL.ALMA_MATER] || "";
    const retired = row[COL.RETIRED] === true || row[COL.RETIRED] === 'TRUE';

    if (!nameJp || !email) continue;

    // 在籍フィルタ処理
    if (filterStatus === 'active' && retired) continue;
    if (filterStatus === 'retired' && !retired) continue;
    // filterStatus === 'all' の場合は全て表示

    if (filterGrade && normalize(grade) !== filterGrade) continue;
    if (filterField && normalize(field) !== filterField) continue;

    let isOrgMatch = !filterOrg;
    let isDeptMatch = !filterDept;
    let isRoleMatch = !filterRole;
    if (filterOrg || filterDept || filterRole) { isOrgMatch = false; isDeptMatch = false; isRoleMatch = false; }

    const depts = [];
    const affiliationTokens = [];
    const orgLabels = [];
    const deptLabels = [];
    const roleLabels = [];
    for (let k = 0; k < AFFILIATION_SLOTS; k++) {
      const deptCol = _affiliationDeptCol(k);
      const roleCol = _affiliationRoleCol(k);
      if (roleCol >= row.length) break;
      const affCode = String(row[deptCol] || '').trim();
      const aff = _parseAffiliationCode(affCode, masters);
      const rolePid = String(row[roleCol] || '').trim();
      const roleLabel = _toLabelOrEmpty(rolePid, masters.role.byPid);
      const deptLabel = aff.dept;
      const orgLabel = aff.org;

      const org = normalize(orgLabel);
      const dept = normalize(deptLabel);
      const role = normalize(roleLabel);
      const rawOrg = orgLabel;
      const rawDept = deptLabel;
      const rawRole = roleLabel;

      if (rawOrg || rawDept || rawRole) {
        const affiliationLabel = rawDept ? ((rawOrg ? (rawOrg + '/' + rawDept) : rawDept)) : rawOrg;
        depts.push([affiliationLabel, rawRole].filter(Boolean).join(' '));
      }

      if (rawOrg) orgLabels.push(rawOrg);
      if (rawDept) deptLabels.push(rawDept);
      if (rawRole) roleLabels.push(rawRole);

      if (org) affiliationTokens.push(org);
      if (dept) affiliationTokens.push(dept);
      if (role) affiliationTokens.push(role);

      if (filterOrg && org === filterOrg) isOrgMatch = true;
      if (filterDept && dept === filterDept) isDeptMatch = true;
      if (filterRole && role === filterRole) isRoleMatch = true;
    }

    const searchString = normalize(`${nameJp} ${nameEn} ${email} ${almaMater} ${studentId} ${affiliationTokens.join(' ')}`);
    if (q && !searchString.includes(q)) continue;

    if (filterOrg && !isOrgMatch) continue;
    if (filterDept && !isDeptMatch) continue;
    if (filterRole && !isRoleMatch) continue;

    results.push({
      name: nameJp,
      email: email,
      org: Array.from(new Set(orgLabels)),
      department: Array.from(new Set(deptLabels)),
      role: Array.from(new Set(roleLabels)),
      departmentText: depts.join(", ") || "所属なし",
      grade: grade,
      field: field
    });
  }
  return results;
}
