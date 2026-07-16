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

    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const masters = _loadMasterMaps(ss);
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
      else {
        const deptMatch = String(f || '').match(/^所属部門(\d{1,2})$/);
        const roleMatch = String(f || '').match(/^役職(\d{1,2})$/);
        const orgMatch = String(f || '').match(/^所属局(\d{1,2})$/);
        if (orgMatch) {
          const idx = Number(orgMatch[1]);
          if (idx >= 1 && idx <= AFFILIATION_SLOTS) pushIf('所属局' + idx, 'ORG_' + idx);
        } else if (deptMatch) {
          const idx = Number(deptMatch[1]);
          if (idx >= 1 && idx <= AFFILIATION_SLOTS) pushIf('所属部門' + idx, _affiliationDeptCol(idx - 1));
        } else if (roleMatch) {
          const idx = Number(roleMatch[1]);
          if (idx >= 1 && idx <= AFFILIATION_SLOTS) pushIf('役職' + idx, _affiliationRoleCol(idx - 1));
        } else if (f === '車所有') pushIf('車所有', COL.CAR_OWNER);
        else if (f === 'Admin') pushIf('Admin', COL.ADMIN);
      }
    }

    if (indices.length === 0) throw new Error('出力項目が選択されていません');

    // Read users data rows and build output rows
    const outRows = [];
    for (let i = 1; i < usersData.length; i++) {
      const row = usersData[i];
      // skip empty rows (メールアドレスが空なら無視)
      if (!row[COL.EMAIL]) continue;
      const outRow = indices.map(ci => {
        let v = '';
        if (typeof ci === 'string' && /^ORG_\d+$/.test(ci)) {
          const slot = Number(ci.replace('ORG_', '')) - 1;
          const aff = _parseAffiliationCode(row[_affiliationDeptCol(slot)], masters);
          v = aff.org || '';
        } else {
          v = row[ci];
        }
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
        if (ci === COL.GRADE) return _toLabelOrEmpty(v, masters.grade.byPid);
        if (ci === COL.FIELD) return _toLabelOrEmpty(v, masters.field.byPid);
        if (typeof ci === 'string' && /^ORG_\d+$/.test(ci)) return String(v || '');
        for (let k = 0; k < AFFILIATION_SLOTS; k++) {
          if (ci === _affiliationDeptCol(k)) {
            const aff = _parseAffiliationCode(v, masters);
            return aff.dept || aff.org || aff.code;
          }
          if (ci === _affiliationRoleCol(k)) return _toLabelOrEmpty(v, masters.role.byPid);
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
function createRosterCsv(sessionToken, params) {
  try {
    const login = getLoginUser(sessionToken);
    if (login.status !== 'authorized') throw new Error('認証されていません');

    params = params || {};
    const selectedFields = params.selectedFields || [];
    const filter = params.filter || { type: 'all' };

    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const masters = _loadMasterMaps(ss);
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
    const add = (idx) => { if ((typeof idx === 'number' && idx >= 0) || (typeof idx === 'string' && /^ORG_\d+$/.test(idx))) allowedIdxSet.add(idx); };

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
      else {
        const deptMatch = String(f || '').match(/^所属部門(\d{1,2})$/);
        const roleMatch = String(f || '').match(/^役職(\d{1,2})$/);
        const orgMatch = String(f || '').match(/^所属局(\d{1,2})$/);
        if (orgMatch) {
          const idx = Number(orgMatch[1]);
          if (idx >= 1 && idx <= AFFILIATION_SLOTS) add('ORG_' + idx);
        } else if (deptMatch) {
          const idx = Number(deptMatch[1]);
          if (idx >= 1 && idx <= AFFILIATION_SLOTS) add(_affiliationDeptCol(idx - 1));
        } else if (roleMatch) {
          const idx = Number(roleMatch[1]);
          if (idx >= 1 && idx <= AFFILIATION_SLOTS) add(_affiliationRoleCol(idx - 1));
        } else if (f === '車所有') add(COL.CAR_OWNER);
        else if (f === 'Admin') add(COL.ADMIN);
      }
    }

    const indices = Array.from(allowedIdxSet);
    const headersOut = indices.map(i => {
      if (typeof i === 'string' && /^ORG_\d+$/.test(i)) return '所属局' + i.replace('ORG_', '');
      return HEADER_USERS[i] || '';
    });

    if (indices.length === 0) throw new Error('出力項目が選択されていません');

    // フィルタ処理 (status: 'active'|'retired'|'all'), grade, field
    const statusFilter = (filter && filter.status) ? filter.status : 'active';
    const gradeFilter = (filter && typeof filter.grade !== 'undefined') ? filter.grade : null;
    const fieldFilter = (filter && typeof filter.field !== 'undefined') ? filter.field : null;
    if ((statusFilter === 'retired' || statusFilter === 'all') && !isAdmin) {
      throw new Error('退局者または全員の出力は管理者のみ可能です');
    }
    // Build filtered list of data rows first, then sort according to Options ordering
    const filteredRows = [];
    for (let i = 1; i < usersData.length; i++) {
      const row = usersData[i];
      if (!row[COL.EMAIL]) continue;
      if (gradeFilter) {
        if (Array.isArray(gradeFilter)) {
          if (gradeFilter.length > 0 && gradeFilter.indexOf(String(row[COL.GRADE] || '')) === -1) continue;
        } else {
          if ((row[COL.GRADE] || '') !== String(gradeFilter)) continue;
        }
      }
      if (fieldFilter) {
        if (Array.isArray(fieldFilter)) {
          if (fieldFilter.length > 0 && fieldFilter.indexOf(String(row[COL.FIELD] || '')) === -1) continue;
        } else {
          if ((row[COL.FIELD] || '') !== String(fieldFilter)) continue;
        }
      }

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
            const aff1 = _parseAffiliationCode(String(row[_affiliationDeptCol(0)] || '').trim(), masters);
            const dept1 = aff1.dept;
            const org1 = aff1.org;
            if (targetOrg && org1 === targetOrg) {
              if (!targetDept || dept1 === targetDept) { matched = true; break; }
            }
          } else {
            for (let k = 0; k < AFFILIATION_SLOTS; k++) {
              const aff = _parseAffiliationCode(String(row[_affiliationDeptCol(k)] || '').trim(), masters);
              const d = aff.dept;
              const o = aff.org;
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

    const gradeOrder = Object.keys(masters.grade.byPid);
    const fieldOrder = Object.keys(masters.field.byPid);

    const idxIn = (arr, v) => { if (!arr || !arr.length) return -1; if (!v) return arr.length + 1; const i = arr.indexOf(String(v)); return i === -1 ? arr.length : i; };

    filteredRows.sort((A, B) => {
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

    const outRows = filteredRows.map(row => indices.map(ci => {
      if (typeof ci === 'string' && /^ORG_\d+$/.test(ci)) {
        const slot = Number(ci.replace('ORG_', '')) - 1;
        const aff = _parseAffiliationCode(row[_affiliationDeptCol(slot)], masters);
        return aff.org || '';
      }
      return row[ci] === undefined || row[ci] === null ? '' : row[ci];
    }));

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
      if (origCol === COL.GRADE) return _toLabelOrEmpty(cell, masters.grade.byPid);
      if (origCol === COL.FIELD) return _toLabelOrEmpty(cell, masters.field.byPid);
      if (typeof origCol === 'string' && /^ORG_\d+$/.test(origCol)) return String(cell || '');
      for (let k = 0; k < AFFILIATION_SLOTS; k++) {
        if (origCol === _affiliationDeptCol(k)) {
          const aff = _parseAffiliationCode(cell, masters);
          return aff.dept || aff.org || aff.code;
        }
        if (origCol === _affiliationRoleCol(k)) return _toLabelOrEmpty(cell, masters.role.byPid);
      }
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

    // optionally append collection payment summary to the right of roster (join by email)
    let collectionAppendHeaders = [];
    let collectionByEmail = {};
    if (params && params.collectionId) {
      try {
        const sum = fetchCollectionSummary(sessionToken, params.collectionId);
        if (sum && sum.success && Array.isArray(sum.perPerson)) {
          collectionAppendHeaders = ['請求額', '受領額', '過不足'];
          sum.perPerson.forEach(p => {
            const em = String(p.email || '').trim().toLowerCase();
            if (!em) return;
            const expected = Number(p.expected || 0);
            const collected = Number(p.collected || 0);
            const diff = collected - expected;
            collectionByEmail[em] = [expected, collected, diff];
          });
        }
      } catch (e) {
        // ignore collection append errors
      }
    }

    const finalHeaders = headersOut.concat(surveyAppendHeaders).concat(collectionAppendHeaders);
    const finalRows = formattedOutRows.map((r, i) => {
      let out = r.slice();
      if (surveyAppendHeaders.length) {
        const email = String(filteredRows[i][COL.EMAIL] || '').trim().toLowerCase();
        const extra = surveyByEmail[email] ? surveyByEmail[email].values : new Array(surveyAppendHeaders.length).fill('');
        out = out.concat(extra);
      }
      if (collectionAppendHeaders.length) {
        const email = String(filteredRows[i][COL.EMAIL] || '').trim().toLowerCase();
        const extra = collectionByEmail[email] ? collectionByEmail[email] : new Array(collectionAppendHeaders.length).fill('');
        out = out.concat(extra);
      }
      return out;
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
