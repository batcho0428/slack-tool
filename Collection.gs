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
    const masters = _loadMasterMaps(ss);
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
      const aff = _parseAffiliationCode(r[3], masters);
      const createdRaw = r[4];
      let createdVal = null;
      try {
        if (createdRaw instanceof Date) createdVal = createdRaw.getTime();
        else if (typeof createdRaw === 'number') createdVal = createdRaw;
        else if (createdRaw) createdVal = String(createdRaw);
        else createdVal = null;
      } catch (e) { createdVal = String(createdRaw); }
      return { id: String(r[0]||''), title: String(r[1]||''), spreadsheetUrl: String(r[2]||''), inChargeOrg: aff.orgPid || '', inChargeDept: aff.deptPid || '', inChargeCode: aff.code || '', inChargeOrgLabel: aff.org || '', inChargeDeptLabel: aff.dept || aff.org || '', createdAt: createdVal, createdBy: String(r[5]||'') };
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
  // payload: { title, spreadsheetUrl, inChargeDept }
  ensureCollectionsSheets();
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const masters = _loadMasterMaps(ss);
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
  const row = [id, payload.title || '', payload.spreadsheetUrl || '', _buildAffiliationStorageCode(orgVal, deptVal, masters), now, createdBy];
  s.appendRow(row);
  // log creation
  try {
    const ls = ss.getSheetByName(SHEET_LOGS) || ss.insertSheet(SHEET_LOGS);
    ls.appendRow([new Date(), 'createCollection', sessionToken || '', 'created', JSON.stringify({ id: id, title: payload.title || '' })]);
  } catch (e) {}
  return { success: true, id };
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
function getCollectionRowDetails(sessionToken, collectionId, recipientEmail) {
  const login = getLoginUser(sessionToken);
  if (!login || login.status !== 'authorized') throw new Error('認証されていません');
  if (!collectionId || !recipientEmail) return { success: false, message: 'パラメータが不足しています' };
  ensureCollectionsSheets();
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const s = ss.getSheetByName(SHEET_COLLECTIONS);
  const lr = s.getLastRow();
  if (lr < 2) return { success: false, message: 'Collectionsが空です' };
  const rows = s.getRange(2,1,lr-1, HEADER_COLLECTIONS.length).getValues();
  const found = rows.find(r => String(r[0]||'') === String(collectionId));
  if (!found) return { success:false, message: '指定のCollectionが見つかりません' };
  const spreadsheetUrl = String(found[2] || '');
  const sid = _extractSpreadsheetId(spreadsheetUrl) || spreadsheetUrl;
  if (!sid) return { success: false, message: 'スプレッドシートIDが取得できません' };
  let target;
  try { target = SpreadsheetApp.openById(sid); } catch (e) { return { success:false, message: 'スプレッドシートを開けません: '+ String(e) }; }
  const sheet = target.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  if (!data || data.length < 1) return { success: false, message: 'スプレッドシートが空です' };
  const headers = (data[0] || []).map(c => String(c||''));
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
  const targetEmail = String(recipientEmail || '').trim().toLowerCase();
  let matchedRow = null;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const mail = String(row[emailIdx]||'').trim().toLowerCase();
    if (mail && mail === targetEmail) {
      matchedRow = row;
      break;
    }
  }
  if (!matchedRow) return { success:false, message: '該当メールの行が見つかりません' };
  const rowValues = matchedRow.map(v => _safeValueForClient(v));
  return { success: true, headers: headers, row: rowValues, emailColumn: emailIdx, amountColumn: amountIdx };
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
  const col = { id: String(found[0]||''), title: String(found[1]||''), spreadsheetUrl: String(found[2]||''), inChargeOrg: '', inChargeDept: String(found[3]||''), createdAt: found[4], createdBy: String(found[5]||'') };

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
    let amt = Number(l.amount || 0);
    if (l.type === 'おつり' || l.type === '返金') {
      if (amt > 0) amt = -Math.abs(amt);
    }
    collectedByEmail[key].total += amt;
    collectedByEmail[key].entries.push(l);
  });

  // aggregate by handler (collector) so UI can show "受け取った人" (担当者) breakdown
  const collectedByHandler = {};
  logsOf.forEach(l => {
    const handlerKey = (l.handler || '').toLowerCase();
    if (!collectedByHandler[handlerKey]) collectedByHandler[handlerKey] = { handler: l.handler, total: 0, entries: [] };
    let amtH = Number(l.amount || 0);
    if (l.type === 'おつり' || l.type === '返金') {
      if (amtH > 0) amtH = -Math.abs(amtH);
    }
    collectedByHandler[handlerKey].total += amtH;
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
  const row = [collectionId, ts, recipientEmail || '', type || '受領', a, handlerEmail || ''];
  sl.appendRow(row);
  return { success: true };
}
function recordPaymentWithChange(sessionToken, collectionId, recipientEmail, receivedAmount, expectedAmount, handlerEmail) {
  // receivedAmount: actual received (positive). expectedAmount: base amount (initial receive). We record full received as 受領, then record おつり as negative if needed.
  ensureCollectionsSheets();
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const sl = ss.getSheetByName(SHEET_COLLECTIONS_LOG);
  const ts = new Date();
  const r = Number(receivedAmount) || 0;
  const base = Number(expectedAmount) || 0;
  // record full received
  sl.appendRow([collectionId, ts, recipientEmail || '', '受領', r, handlerEmail || '']);
  if (base === 0 && r !== 0) {
    // when expected is zero, refund the full amount as change
    sl.appendRow([collectionId, ts, recipientEmail || '', 'おつり', -Math.abs(r), handlerEmail || '']);
    return { success: true };
  }
  const change = r - base;
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
    const masters = _loadMasterMaps(ss);
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
        let inChargeOrg = (typeof payload.inChargeOrg !== 'undefined') ? String(payload.inChargeOrg) : '';
        let inChargeDept = (typeof payload.inChargeDept !== 'undefined') ? String(payload.inChargeDept) : String(r[3] || '');
        if (inChargeOrg === '選択' || inChargeOrg === '選択してください') inChargeOrg = '';
        if (inChargeDept === '選択' || inChargeDept === '選択してください') inChargeDept = '';
        const createdAt = r[4] || new Date();
        const createdBy = r[5] || '';
        const newRow = [collectionId, title, spreadsheetUrl, _buildAffiliationStorageCode(inChargeOrg, inChargeDept, masters), createdAt, createdBy];
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
