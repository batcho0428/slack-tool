// @ts-nocheck
'use client';
/* eslint-disable */

import { useEffect, useRef, useState } from 'react';

        export default function CollectTab({ user, runGas, parseCsv, formatPreviewBirthday, DetailsModal, DialogModal }) {
                    const [collections, setCollections] = useState([]);
                    const [loading, setLoading] = useState(true);
                    const [selectedId, setSelectedId] = useState('');
                    const [summary, setSummary] = useState(null);
                    const [summaryLoading, setSummaryLoading] = useState(false);
                    const [modalOpen, setModalOpen] = useState(false);
                    const [newMode, setNewMode] = useState('sheet'); // 'sheet' or 'forms'
                    const [sheetUrl, setSheetUrl] = useState('');
                    const [title, setTitle] = useState('');
                    const [editingId, setEditingId] = useState('');
                    const [editSelectOpen, setEditSelectOpen] = useState(false);
                    const [editSelectLoading, setEditSelectLoading] = useState(false);
                    const [formsList, setFormsList] = useState([]);
                    const [formsLoading, setFormsLoading] = useState(false);
                    const [savingCollection, setSavingCollection] = useState(false);
                    const [regDoneOpen, setRegDoneOpen] = useState(false);
                    const [regDoneMessage, setRegDoneMessage] = useState('');
                    const [collectionsLoading, setCollectionsLoading] = useState(false);
                    const [orgOptions, setOrgOptions] = useState({ orgs: [], deptMaster: [] });
                    const [selectedOrg, setSelectedOrg] = useState('');
                    const [selectedDept, setSelectedDept] = useState('');
                    const [selectedFormRef, setSelectedFormRef] = useState('');
                    const [accountingOpen, setAccountingOpen] = useState(false);
                    const [accountTarget, setAccountTarget] = useState(null);
                    const [accountReceiptOpen, setAccountReceiptOpen] = useState(false);
                    const [accountReceiptLoading, setAccountReceiptLoading] = useState(false);
                    const [accountReceiptData, setAccountReceiptData] = useState(null);
                    const [accountFlowOpen, setAccountFlowOpen] = useState(false);
                    const [flowStep, setFlowStep] = useState('org'); // org -> grade -> person -> pay
                    const [flowOrgList, setFlowOrgList] = useState([]);
                    const [flowSelectedOrg, setFlowSelectedOrg] = useState('');
                    const [flowGrades, setFlowGrades] = useState([]);
                    const [flowSelectedGrade, setFlowSelectedGrade] = useState('');
                    const [flowPeople, setFlowPeople] = useState([]);
                    const [flowPeopleCache, setFlowPeopleCache] = useState({});
                    const [flowLoading, setFlowLoading] = useState(false);
                    const [flowShowOnlyUnbalanced, setFlowShowOnlyUnbalanced] = useState(false);
                    const [flowSelectedPerson, setFlowSelectedPerson] = useState(null);
                    const [flowReceiveAmount, setFlowReceiveAmount] = useState(0);
                    const [flowReceiveInitialAmount, setFlowReceiveInitialAmount] = useState(0);
                    const [quickReceiveAmount, setQuickReceiveAmount] = useState(0);
                    const [quickReceiveInitialAmount, setQuickReceiveInitialAmount] = useState(0);
                    const [flowHistoryLoading, setFlowHistoryLoading] = useState(false);
                    const [flowHistoryEntries, setFlowHistoryEntries] = useState([]);
                    const [flowHistoryEmail, setFlowHistoryEmail] = useState('');
                    const flowHistoryPromiseRef = useRef(null);
                    const [handlerNameMap, setHandlerNameMap] = useState({});
                    const [recipientNameMap, setRecipientNameMap] = useState({});
                    const [expandedCollectors, setExpandedCollectors] = useState({});
                    const [selectCollectionOpen, setSelectCollectionOpen] = useState(false);
                    const [myCollections, setMyCollections] = useState([]);
                    const [myCollectionsLoading, setMyCollectionsLoading] = useState(false);
                    const [myCollectionsModalOpen, setMyCollectionsModalOpen] = useState(false);
                    const [myCollectionsModalLoading, setMyCollectionsModalLoading] = useState(false);
                    const [myCollectionsModalDetails, setMyCollectionsModalDetails] = useState(null);

            useEffect(()=>{
                setLoading(true);
                const token = localStorage.getItem('slack_app_session');
                runGas('ensureCollectionsSheets').catch(()=>{});
                // initial lightweight load (dropdown will refresh on focus)
                runGas('listCollections', token).then(res=>{ setCollections(res || []); }).catch(()=>{}).finally(()=>setLoading(false));
                // prefetch org list once to avoid refetch on accounting open
                runGas('getSearchOptions', token).then(opts=>{
                    setFlowOrgList((opts && opts.orgs) ? opts.orgs : []);
                }).catch(()=>{});
            }, []);

            const loadMyCollections = async (listOverride) => {
                if (!user || !user.email) { setMyCollections([]); return; }
                setMyCollectionsLoading(true);
                try {
                    const token = localStorage.getItem('slack_app_session');
                    const list = listOverride || collections || [];
                    const targetEmail = String(user.email || '').toLowerCase();
                    const out = [];
                    for (const c of list) {
                        if (!c || !c.id) continue;
                        try {
                            const res = await runGas('fetchCollectionSummary', token, c.id);
                            const entry = (res && res.perPerson)
                                ? res.perPerson.find(p => (p.email || '').toLowerCase() === targetEmail)
                                : null;
                            if (!entry) continue;
                            const expected = Number(entry.expected || 0);
                            const collected = Number(entry.collected || 0);
                            out.push({
                                id: c.id,
                                title: c.title || c.spreadsheetUrl || '無題',
                                org: c.inChargeOrg || '',
                                dept: c.inChargeDept || '',
                                expected,
                                diff: expected - collected
                            });
                        } catch (e) {}
                    }
                    setMyCollections(out);
                } finally {
                    setMyCollectionsLoading(false);
                }
            };

            const openMyCollectionDetails = async (item) => {
                if (!item || !item.id || !user || !user.email) return;
                setMyCollectionsModalOpen(true);
                setMyCollectionsModalLoading(true);
                setMyCollectionsModalDetails(null);
                try {
                    const token = localStorage.getItem('slack_app_session');
                    const res = await runGas('fetchCollectionSummary', token, item.id);
                    let rowDetails = null;
                    try {
                        rowDetails = await runGas('getCollectionRowDetails', token, item.id, user.email);
                    } catch (e) {
                        rowDetails = null;
                    }
                    const targetEmail = String(user.email || '').toLowerCase();
                    const entry = (res && res.perPerson)
                        ? res.perPerson.find(p => (p.email || '').toLowerCase() === targetEmail)
                        : null;
                    if (!entry) {
                        setMyCollectionsModalDetails({ error: '対象データが見つかりません' });
                        return;
                    }
                    const expected = Number(entry.expected || 0);
                    const collected = Number(entry.collected || 0);
                    const diff = expected - collected;
                    const label = diff === 0 ? '差額' : (diff > 0 ? '不足' : '返金');
                    const fmtYen = (n) => Number(n || 0).toLocaleString() + '円';
                    const latestTs = Array.isArray(entry.entries) && entry.entries.length > 0
                        ? Math.max(...entry.entries.map(e => Number(e.timestamp || 0)))
                        : null;
                    const historyEntries = Array.isArray(entry.entries)
                        ? entry.entries.slice().sort((a, b) => Number(b.timestamp || 0) - Number(a.timestamp || 0))
                        : [];
                    const handlerEmails = Array.from(new Set(historyEntries.map(h => String(h.handler || '').trim().toLowerCase()).filter(Boolean)));
                    const handlerMap = {};
                    for (const em of handlerEmails) {
                        try {
                            const found = await runGas('searchRecipients', { query: em, status: 'all' });
                            if (found && found.length && found[0].name) handlerMap[em] = found[0].name;
                        } catch (e) {}
                    }
                    const historyEntriesWithNames = historyEntries.map(h => {
                        const em = String(h.handler || '').trim().toLowerCase();
                        return { ...h, handlerName: handlerMap[em] || '' };
                    });
                    const answers = {
                        '担当局/担当部門': (item.org || '-') + (item.dept ? (' / ' + item.dept) : ''),
                        '集金額': fmtYen(expected),
                        '受領済み': fmtYen(collected),
                        '過不足': label + ' ' + fmtYen(Math.abs(diff))
                    };
                    const orderedKeys = Object.keys(answers);
                    if (rowDetails && rowDetails.success && Array.isArray(rowDetails.headers) && Array.isArray(rowDetails.row)) {
                        rowDetails.headers.forEach((h, idx) => {
                            const key = String(h || '').trim();
                            if (!key) return;
                            const val = rowDetails.row[idx];
                            if (val === null || typeof val === 'undefined' || String(val).trim() === '') return;
                            if (typeof answers[key] !== 'undefined') return;
                            answers[key] = String(val);
                            orderedKeys.push(key);
                        });
                    }
                    setMyCollectionsModalDetails({
                        viewType: 'collection',
                        headers: orderedKeys,
                        response: {
                            timestamp: latestTs,
                            score: expected,
                            scoreFormatted: fmtYen(expected),
                            answers: answers
                        },
                        historyEntries: historyEntriesWithNames,
                        scoreName: '集金額',
                        scoreUnit: '円'
                    });
                } catch (e) {
                    setMyCollectionsModalDetails({ error: e && e.message ? e.message : String(e) });
                } finally {
                    setMyCollectionsModalLoading(false);
                }
            };

            useEffect(() => {
                loadMyCollections().catch(()=>{});
            }, [user, collections]);

            const openNew = async () => {
                setTitle(''); setSheetUrl(''); setNewMode('sheet'); setSelectedFormRef(''); setSelectedOrg(''); setSelectedDept(''); setEditingId('');
                setModalOpen(true);
                setFormsLoading(true);
                try {
                    const token = localStorage.getItem('slack_app_session');
                    const formsRes = await runGas('listFormDefinitions', token).catch(()=>({ success: true, items: [] }));
                    const opts = await runGas('getSearchOptions').catch(()=>({ orgs: [], deptMaster: [] }));
                    if (formsRes && formsRes.success) setFormsList(formsRes.items || []); else setFormsList([]);
                    if (opts) setOrgOptions({ orgs: opts.orgs || [], deptMaster: opts.deptMaster || [] });
                } catch (e) {
                    setFormsList([]);
                } finally {
                    setFormsLoading(false);
                }
            };

            const saveNew = async () => {
                setSavingCollection(true);
                try {
                    const token = localStorage.getItem('slack_app_session');
                    // If selected source is Forms, validate importability (email & amount columns exist)
                    if (newMode === 'forms' && sheetUrl) {
                        const parsed = await runGas('parseSourceSpreadsheet', token, sheetUrl);
                        if (!parsed || parsed.success === false) {
                            // show error modal
                            setRegDoneMessage('選択したソースはインポート要件を満たしていません:\n' + (parsed && parsed.message ? parsed.message : '不明なエラー'));
                            setRegDoneOpen(true);
                            setSavingCollection(false);
                            return;
                        }
                    }
                    // normalize '選択' placeholder to empty string when saving
                    let orgVal = selectedOrg;
                    let deptVal = selectedDept;
                    if (orgVal === '選択' || orgVal === '選択してください') orgVal = '';
                    if (deptVal === '選択' || deptVal === '選択してください') deptVal = '';
                    const payload = { title: title || ('無題 ' + (new Date()).toLocaleString()), spreadsheetUrl: sheetUrl || '', inChargeOrg: orgVal || '', inChargeDept: deptVal || '' };
                    if (editingId) {
                        const res = await runGas('updateCollection', token, editingId, payload);
                        if (!res || !res.success) throw new Error(res && res.message ? res.message : '更新に失敗しました');
                        setRegDoneMessage('更新が完了しました');
                    } else {
                        const res = await runGas('createCollection', token, payload);
                        if (!res || !res.success) throw new Error(res && res.message ? res.message : '作成に失敗しました');
                        setRegDoneMessage('登録が完了しました');
                    }
                    const list = await runGas('listCollections', token);
                    setCollections(list || []);
                    loadMyCollections(list || []).catch(()=>{});
                    setModalOpen(false);
                    setEditingId('');
                    setRegDoneOpen(true);
                } catch (e) { setRegDoneMessage('エラー: ' + (e && e.message ? e.message : String(e))); setRegDoneOpen(true); }
                finally { setSavingCollection(false); }
            };

            const deleteThisCollection = async () => {
                if (!editingId) return;
                if (!confirm('本当にこの集金を削除しますか？ この操作は取り消せません。')) return;
                setSavingCollection(true);
                try {
                    const token = localStorage.getItem('slack_app_session');
                    const res = await runGas('deleteCollection', token, editingId);
                    if (!res || !res.success) throw new Error(res && res.message ? res.message : '削除に失敗しました');
                    const list = await runGas('listCollections', token);
                    setCollections(list || []);
                    loadMyCollections(list || []).catch(()=>{});
                    setModalOpen(false);
                    setEditingId('');
                    setRegDoneMessage('削除しました');
                    setRegDoneOpen(true);
                } catch (e) {
                    setRegDoneMessage('エラー: ' + (e && e.message ? e.message : String(e)));
                    setRegDoneOpen(true);
                } finally {
                    setSavingCollection(false);
                }
            };

            const refreshSummary = async (idOverride, options = {}) => {
                const id = idOverride || selectedId;
                if (!id) return null;
                const showLoading = (typeof options.showLoading === 'boolean') ? options.showLoading : true;
                const fetchHandlerNames = (typeof options.fetchHandlerNames === 'boolean') ? options.fetchHandlerNames : true;
                if (showLoading) setSummaryLoading(true);
                const token = localStorage.getItem('slack_app_session');
                try {
                    const res = await runGas('fetchCollectionSummary', token, id);
                    setSummary(res);
                    if (fetchHandlerNames) {
                        // fetch handler names for perCollector
                        try {
                            const handlers = (res && res.perCollector) ? res.perCollector.map(h=>h.handler).filter(Boolean).map(x=>x.toLowerCase()) : [];
                            const unique = Array.from(new Set(handlers)).filter(Boolean);
                            const map = {};
                            for (const h of unique) {
                                if (!h) continue;
                                try {
                                    const found = await runGas('searchRecipients', { query: h, status: 'all' });
                                    if (found && found.length && found[0].name) map[h] = found[0].name;
                                } catch(e) {}
                            }
                            setHandlerNameMap(map);
                        } catch(e) { console.debug('handler name lookup failed', e); }
                    }
                    return res;
                } catch (e) {
                    setSummary({ error: e.message || e });
                    return null;
                } finally {
                    if (showLoading) setSummaryLoading(false);
                }
            };

            const fetchHistoryForEmail = async (email) => {
                if (!email || !selectedId) return [];
                const em = String(email || '').toLowerCase();
                if (flowHistoryEmail === em && flowHistoryLoading && flowHistoryPromiseRef.current) {
                    return flowHistoryPromiseRef.current;
                }
                setFlowHistoryLoading(true);
                setFlowHistoryEntries([]);
                setFlowHistoryEmail(em);
                const promise = (async () => {
                    try {
                        const token = localStorage.getItem('slack_app_session');
                        const res = await runGas('fetchCollectionSummary', token, selectedId);
                        const entry = (res && res.perPerson) ? res.perPerson.find(p => (p.email || '').toLowerCase() === em) : null;
                        const entries = entry && Array.isArray(entry.entries) ? entry.entries : [];
                        setFlowHistoryEntries(entries);
                        return entries;
                    } catch (e) {
                        setFlowHistoryEntries([]);
                        return [];
                    } finally {
                        setFlowHistoryLoading(false);
                        flowHistoryPromiseRef.current = null;
                    }
                })();
                flowHistoryPromiseRef.current = promise;
                return promise;
            };

            const buildReceiptData = (expected, received, initialReceive, mode, label, historyEntries) => {
                const init = Number(initialReceive || 0);
                const delta = Number(received || 0) - init;
                const lbl = String(label || '');
                const isRefund = (lbl === '返金' || Number(received || 0) < 0);
                if (isRefund) {
                    const refundAmount = Math.abs(Number(received || 0));
                    const baseInit = Math.abs(init);
                    const refundDiff = baseInit - refundAmount;
                    let changeLabel = '差額';
                    if (refundDiff > 0) changeLabel = '預り金';
                    else if (refundDiff < 0) changeLabel = '不足';
                    return {
                        expected: Number(expected || 0),
                        received: Number(received || 0),
                        initialReceive: init,
                        changeLabel,
                        changeAmount: Number(refundDiff || 0),
                        historyEntries: Array.isArray(historyEntries) ? historyEntries : []
                    };
                }

                let changeLabel = 'おつり';
                if (delta < 0) {
                    changeLabel = '不足';
                } else if (delta > 0) {
                    changeLabel = (mode === 'debt') ? '預り金' : 'おつり';
                } else {
                    // exact receive should be labeled as おつり
                    changeLabel = (mode === 'debt') ? '預り金' : 'おつり';
                }
                return {
                    expected: Number(expected || 0),
                    received: Number(received || 0),
                    initialReceive: init,
                    changeLabel,
                    changeAmount: Number(delta || 0),
                    historyEntries: Array.isArray(historyEntries) ? historyEntries : []
                };
            };

            const onSelect = async (id) => {
                setSelectedId(id);
                setSummary(null);
                if (!id) return;
                await refreshSummary(id);
            };

            const openAccounting = (person) => {
                // legacy single-person accounting (quick)
                setAccountTarget(person);
                if (person) {
                    const expected = Number(person.expected || 0);
                    const collected = Number(person.collected || 0);
                    const diff = expected - collected;
                    const init = diff < 0 ? -Math.abs(diff) : Math.abs(diff);
                    setQuickReceiveAmount(init);
                    setQuickReceiveInitialAmount(init);
                } else {
                    setQuickReceiveAmount(0);
                    setQuickReceiveInitialAmount(0);
                }
                setAccountingOpen(true);
            };

            const openAccountingFromEntry = async (email) => {
                if (!email) return;
                setFlowLoading(true);
                try {
                    let res = summary;
                    if (!res || !res.success || !selectedId) {
                        res = await refreshSummary(selectedId, { showLoading: false, fetchHandlerNames: false });
                    }
                    fetchHistoryForEmail(email).catch(()=>{});
                    const em = (email || '').toLowerCase();
                    const entry = (res && res.perPerson) ? res.perPerson.find(p=> (p.email||'').toLowerCase() === em) : null;
                    const name = recipientNameMap[em] || (entry && (entry.name || entry.displayName)) || email;
                    if (entry) {
                        setFlowSelectedPerson({ email: entry.email, name: name, grade: entry.grade, field: entry.field, department: entry.department });
                        const expected = Number(entry.expected || 0);
                        const collected = Number(entry.collected || 0);
                        const diff = expected - collected;
                        // overpaid => initialize as negative
                        const init = diff < 0 ? -Math.abs(diff) : Math.abs(diff);
                        setFlowReceiveAmount(init);
                        setFlowReceiveInitialAmount(init);
                        setFlowStep('pay');
                        setAccountFlowOpen(true);
                    } else {
                        // fallback to quick accounting
                        setAccountTarget({ email: email, expected: 0, collected: 0 });
                        setQuickReceiveAmount(0);
                        setAccountingOpen(true);
                    }
                } catch (e) {
                    alert(e && e.message ? e.message : String(e));
                } finally {
                    setFlowLoading(false);
                }
            };

            const openAccountFlow = async () => {
                setFlowStep('org');
                setFlowSelectedOrg('');
                setFlowSelectedGrade('');
                setFlowPeople([]);
                setFlowSelectedPerson(null);
                setFlowShowOnlyUnbalanced(false);
                setAccountFlowOpen(true);
                if (selectedId) {
                    // always recalc amounts when opening accounting flow (async)
                    refreshSummary(selectedId, { showLoading: false, fetchHandlerNames: false }).catch(()=>{});
                }
            };

            const loadPeopleForOrg = async (org) => {
                setFlowPeople([]);
                const cacheKey = org || 'ALL';
                const cached = flowPeopleCache[cacheKey];
                if (cached && cached.people) {
                    setFlowPeople(cached.people || []);
                    setFlowGrades(cached.grades || []);
                    return;
                }
                setFlowLoading(true);
                try {
                    const token = localStorage.getItem('slack_app_session');
                    const criteria = org ? { org: org, status: 'active' } : { status: 'active' };
                    const res = await runGas('searchRecipients', criteria);
                    const mapped = (res || []).map(r => ({ name: r.name, email: r.email, grade: r.grade, field: r.field, department: r.department }));
                    setFlowPeople(mapped);
                    // derive grades
                    const grades = Array.from(new Set(mapped.map(m=>m.grade).filter(Boolean)));
                    setFlowGrades(grades);
                    setFlowPeopleCache(prev => ({ ...prev, [cacheKey]: { people: mapped, grades: grades } }));
                } catch(e) { setFlowPeople([]); setFlowGrades([]); }
                finally { setFlowLoading(false); }
            };

            const selectPersonInFlow = async (person) => {
                if (!person) return;
                setFlowLoading(true);
                try {
                    setFlowSelectedPerson(person);
                    fetchHistoryForEmail(person.email).catch(()=>{});
                    let res = summary;
                    if (!res || !res.success) {
                        res = await refreshSummary(selectedId, { showLoading: false, fetchHandlerNames: false });
                    }
                    const entry = (res && res.perPerson && Array.isArray(res.perPerson)) ? res.perPerson.find(p=> (p.email||'').toLowerCase() === (person.email||'').toLowerCase()) : null;
                    const expected = entry ? Number(entry.expected || 0) : 0;
                    const collected = entry ? Number(entry.collected || 0) : 0;
                    const diff = expected - collected;
                    // overpaid => initialize as negative
                    const init = diff < 0 ? -Math.abs(diff) : Math.abs(diff);
                    setFlowReceiveAmount(init);
                    setFlowReceiveInitialAmount(init);
                    setFlowStep('pay');
                } catch (e) {
                    alert(e && e.message ? e.message : String(e));
                } finally { setFlowLoading(false); }
            };

            const doFlowReceive = async (mode, label) => {
                if (!flowSelectedPerson) return;
                const token = localStorage.getItem('slack_app_session');
                const handler = (user && user.email) ? user.email : '';
                const collId = selectedId;
                const email = flowSelectedPerson.email;
                const entry = (summary && summary.perPerson) ? summary.perPerson.find(p=> (p.email||'').toLowerCase() === (email||'').toLowerCase()) : null;
                const expected = entry ? Number(entry.expected || 0) : 0;
                const collected = entry ? Number(entry.collected || 0) : 0;
                const remaining = expected - collected;
                const val = Number(flowReceiveAmount || 0);
                const historyPromise = (flowHistoryEmail === String(email || '').toLowerCase() && flowHistoryPromiseRef.current)
                    ? flowHistoryPromiseRef.current
                    : Promise.resolve(flowHistoryEntries);
                setAccountReceiptData(buildReceiptData(expected, val, flowReceiveInitialAmount, mode, label || '', []));
                setAccountReceiptOpen(true);
                setAccountReceiptLoading(true);
                setFlowLoading(true);
                try {
                    if (mode === 'receive') {
                        await runGas('recordPayment', token, collId, email, Math.abs(val), label || '受領', handler);
                    } else if (mode === 'change') {
                        const baseExpected = Number(flowReceiveInitialAmount || 0);
                        await runGas('recordPaymentWithChange', token, collId, email, Math.abs(val), baseExpected, handler);
                    } else if (mode === 'debt') {
                        const lbl = label || '過不足';
                        if (lbl === '返金') {
                            await runGas('recordPayment', token, collId, email, -Math.abs(val), '返金', handler);
                        } else {
                            await runGas('recordPayment', token, collId, email, Number(val || 0), '受領', handler);
                        }
                    }
                    const historyEntries = await historyPromise;
                    setAccountReceiptData(buildReceiptData(expected, val, flowReceiveInitialAmount, mode, label || '', historyEntries));
                    const res = await runGas('fetchCollectionSummary', token, collId);
                    setSummary(res);
                    setAccountFlowOpen(false);
                    setFlowStep('org');
                    setAccountReceiptLoading(false);
                } catch (e) {
                    setAccountReceiptLoading(false);
                    alert(e.message || e);
                }
                finally { setFlowLoading(false); }
            };

            const doReceive = async (mode, label) => {
                // mode: 'receive' | 'change' | 'debt'
                if (!accountTarget) return;
                setFlowLoading(true);
                const token = localStorage.getItem('slack_app_session');
                const handler = (user && user.email) ? user.email : '';
                const collId = selectedId;
                const expected = Number(accountTarget.expected || 0);
                const collected = Number(accountTarget.collected || 0);
                const remaining = expected - collected;
                const input = document.getElementById('collect-input-amount');
                const val = input ? Number(input.value || 0) : remaining;
                const historyPromise = (accountTarget.email && flowHistoryPromiseRef.current)
                    ? flowHistoryPromiseRef.current
                    : Promise.resolve(flowHistoryEntries);
                setAccountReceiptData(buildReceiptData(expected, val, quickReceiveInitialAmount, mode, label || '', []));
                setAccountReceiptOpen(true);
                setAccountReceiptLoading(true);

                try {
                    if (mode === 'receive') {
                        await runGas('recordPayment', token, collId, accountTarget.email, Math.abs(val), label || '受領', handler);
                    } else if (mode === 'change') {
                        const baseExpected = Number(quickReceiveInitialAmount || 0);
                        await runGas('recordPaymentWithChange', token, collId, accountTarget.email, Math.abs(val), baseExpected, handler);
                    } else if (mode === 'debt') {
                        const lbl = label || '過不足';
                        if (lbl === '返金') {
                            // refund should be recorded as a negative amount
                            await runGas('recordPayment', token, collId, accountTarget.email, -Math.abs(val), '返金', handler);
                        } else {
                            // 過不足扱い: keep amount as-is but record as type '受領'
                            await runGas('recordPayment', token, collId, accountTarget.email, Number(val || 0), '受領', handler);
                        }
                    }
                    const historyEntries = await historyPromise;
                    setAccountReceiptData(buildReceiptData(expected, val, quickReceiveInitialAmount, mode, label || '', historyEntries));
                    const res = await runGas('fetchCollectionSummary', token, collId);
                    setSummary(res);
                    setAccountingOpen(false);
                    setAccountReceiptLoading(false);
                } catch (e) {
                    setAccountReceiptLoading(false);
                    alert(e.message || e);
                }
                finally { setFlowLoading(false); }
            };

            const fmt = (n) => {
                const v = Number(n||0);
                const abs = Math.abs(v).toLocaleString();
                return v < 0 ? '△' + abs : abs;
            };

            return (
                <div className="h-full overflow-auto p-4 bg-gray-50">
                    <div className="max-w-4xl mx-auto bg-white rounded-lg shadow-sm p-6 h-full">
                        <div className="flex flex-col md:flex-row md:items-start md:justify-between gap-2 mb-4">
                            <h3 className="text-xl font-bold">集金</h3>
                                <div className="flex flex-wrap items-center gap-2">
                                <button onClick={openNew} className="text-sm md:text-xs bg-green-100 text-green-900 px-4 py-2 rounded-lg">新規登録</button>
                                <button onClick={async ()=>{
                                        setEditSelectOpen(true);
                                        setEditSelectLoading(true);
                                        try {
                                            const token = localStorage.getItem('slack_app_session');
                                            const list = await runGas('listCollections', token).catch(()=>[]);
                                            setCollections(list || []);
                                        } catch(e) { }
                                        finally { setEditSelectLoading(false); }
                                    }} className="text-sm md:text-xs bg-yellow-100 text-yellow-900 px-4 py-2 rounded-lg">編集</button>
                            </div>
                        </div>

                        <div className="mb-4">
                            <div className="border rounded p-4 bg-white">
                                <div className="flex items-center justify-between">
                                    <div className="font-medium">集金マイページ</div>
                                    <div className="text-xs text-gray-500">{myCollectionsLoading ? '読み込み中...' : `${(myCollections || []).filter(c => Number(c.diff || 0) !== 0).length}件`}</div>
                                </div>
                                <div className="mt-3">
                                    <div className="flex items-center justify-end mb-2">
                                        <button onClick={()=>loadMyCollections().catch(()=>{})} className="text-xs text-blue-600 hover:underline">更新</button>
                                    </div>
                                    {myCollectionsLoading ? (
                                        <div className="text-xs text-gray-500">読み込み中...</div>
                                    ) : myCollections.length === 0 ? (
                                        <div className="text-xs text-gray-500">自分宛ての集金がありません。</div>
                                    ) : (
                                        <div className="overflow-auto">
                                            <table className="min-w-full text-sm">
                                                <thead className="bg-gray-100">
                                                    <tr>
                                                        <th className="px-3 py-2 text-left">タイトル</th>
                                                        <th className="px-3 py-2 text-left">担当局 / 担当部門</th>
                                                        <th className="px-3 py-2 text-right">集金額</th>
                                                        <th className="px-3 py-2 text-right">過不足</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    {[...myCollections].sort((a, b) => {
                                                        const aZero = Number(a.diff || 0) === 0;
                                                        const bZero = Number(b.diff || 0) === 0;
                                                        if (aZero === bZero) return 0;
                                                        return aZero ? 1 : -1;
                                                    }).map((c) => {
                                                        const diff = Number(c.diff || 0);
                                                        const label = diff === 0 ? '差額' : (diff > 0 ? '不足' : '返金');
                                                        return (
                                                            <tr key={c.id} role="button" tabIndex={0} onClick={()=>openMyCollectionDetails(c)} className="border-t hover:bg-gray-50 cursor-pointer">
                                                                <td className="px-3 py-2">{c.title}</td>
                                                                <td className="px-3 py-2">{(c.org || '-') + (c.dept ? (' / ' + c.dept) : '')}</td>
                                                                <td className="px-3 py-2 text-right">{fmt(c.expected)}円</td>
                                                                <td className="px-3 py-2 text-right">
                                                                    <span className={diff !== 0 ? 'text-red-600' : 'text-gray-600'}>{label}</span>
                                                                    <span className="ml-2">{fmt(Math.abs(diff))}円</span>
                                                                </td>
                                                            </tr>
                                                        );
                                                    })}
                                                </tbody>
                                            </table>
                                        </div>
                                    )}
                                </div>
                            </div>
                        </div>

                        <DetailsModal isOpen={myCollectionsModalOpen} loading={myCollectionsModalLoading} details={myCollectionsModalDetails} survey={null} onClose={() => { setMyCollectionsModalOpen(false); setMyCollectionsModalDetails(null); setMyCollectionsModalLoading(false); }} />

                        <details className="border rounded p-4 bg-white">
                            <summary className="cursor-pointer list-none flex items-center justify-between">
                                <div className="font-medium">集金機能</div>
                            </summary>
                            <div className="mt-3">
                                <div className="mb-4">
                                    <h4>集金選択</h4>
                                    <div className="flex gap-2 flex-wrap">
                                        <button onClick={async ()=>{
                                            setSelectCollectionOpen(true);
                                            setCollectionsLoading(true);
                                            try {
                                                const token = localStorage.getItem('slack_app_session');
                                                const list = await runGas('listCollections', token).catch(()=>[]);
                                                setCollections(list || []);
                                            } catch(e){} finally { setCollectionsLoading(false); }
                                        }} className="flex-1 text-left border px-3 py-3 rounded-lg text-base md:text-sm">{(collections.find(c=>c.id===selectedId) || { title: '選択してください' }).title}</button>
                                    </div>
                                    {collectionsLoading && <div className="text-xs text-gray-500 mt-1">一覧を更新しています…</div>}
                                </div>

                                {loading && <div className="text-gray-500">読み込み中...</div>}

                                {/* Large account button between summary and debug */}
                                {summary && summary.success && (
                                    <div className="my-4">
                                        <button onClick={openAccountFlow} className="w-full bg-amber-600 text-white text-xl py-4 rounded-lg shadow-md">会計</button>
                                    </div>
                                )}

                                {summary && summary.success && (
                                    <div className="border rounded p-4 bg-white">
                                        <div className="flex items-center justify-between gap-3 mb-2">
                                            <div>
                                                <h4>集金状況</h4>
                                                <br />
                                                {fmt(summary.collectedTotal)}円 / {fmt(summary.expectedTotal)}円（{summary.collectedCount}人/{summary.expectedCount}人）
                                                {(() => {
                                                    const list = summary.perPerson || [];
                                                    const count = list.filter(p => (p.entries || []).length > 0 && Number(p.expected || 0) !== Number(p.collected || 0)).length;
                                                    return (<div className="text-xs text-gray-500 mt-1">受領済みの過不足: {count}件</div>);
                                                })()}
                                            </div>
                                            <button onClick={() => refreshSummary()} className="w-9 h-9 flex items-center justify-center bg-gray-100 rounded-full hover:bg-gray-200" title="再読み込み" aria-label="再読み込み">
                                                <i className="fas fa-sync-alt"></i>
                                            </button>
                                        </div>
                                        {summaryLoading && <div className="text-xs text-gray-500 mb-2">読み込み中...</div>}
                                        <div className="mt-2 space-y-2">
                                            {(summary.perCollector || []).filter(c=>Number(c.total||0) !== 0).map((p,idx)=> (
                                                <div key={idx} className="p-2 border rounded">
                                                    <div className="flex items-center justify-between">
                                                        <div><strong>{(handlerNameMap[(p.handler||'').toLowerCase()] || p.handler) || '（不明）'}</strong></div>
                                                        <div className="text-sm">{fmt(p.total)}</div>
                                                    </div>
                                                    <div className="pl-4 mt-2">
                                                        <div className="flex items-center justify-between"><div className="text-sm">取引件数</div><div className="text-sm">{(p.entries && p.entries.length) || '-'}</div></div>
                                                        <div className="mt-2">
                                                            <button onClick={async ()=>{
                                                                const key = (p.handler||'').toLowerCase();
                                                                const emails = (p.entries||[]).map(e=> (e.email||'').toLowerCase()).filter(Boolean);
                                                                const uniq = Array.from(new Set(emails)).filter(Boolean);
                                                                const rmap = {};
                                                                for (const em of uniq) {
                                                                    if (recipientNameMap[em]) { rmap[em] = recipientNameMap[em]; continue; }
                                                                    try {
                                                                        const found = await runGas('searchRecipients', { query: em, status: 'all' });
                                                                        if (found && found.length && found[0].name) rmap[em] = found[0].name;
                                                                    } catch(e) {}
                                                                }
                                                                setRecipientNameMap(prev=>({ ...prev, ...rmap }));
                                                                setExpandedCollectors(prev=>({ ...prev, [key]: !prev[key] }));
                                                            }} className="text-sm text-blue-600 px-3 py-2 rounded hover:bg-blue-50">詳細を表示</button>
                                                        </div>
                                                        {expandedCollectors[(p.handler||'').toLowerCase()] && (
                                                            <div className="mt-3 space-y-2">
                                                                {p.entries && p.entries.map((e,ei)=> {
                                                                    const em = (e.email||'').toLowerCase();
                                                                    const name = recipientNameMap[em] || e.name || e.email || '-';
                                                                    let affiliation = '-';
                                                                    try {
                                                                        const person = (summary && summary.perPerson) ? summary.perPerson.find(x => (x.email||'').toLowerCase() === em) : null;
                                                                        if (person) {
                                                                            const parts = [];
                                                                            if (person.inChargeOrg || person.org) parts.push(person.inChargeOrg || person.org);
                                                                            if (person.department || person.inChargeDept) parts.push(person.department || person.inChargeDept);
                                                                            affiliation = parts.filter(Boolean).join(' / ') || '-';
                                                                        }
                                                                    } catch(_) {}
                                                                    return (
                                                                    <div key={ei} role="button" tabIndex={0} onClick={() => openAccountingFromEntry(e.email)} className="p-2 border rounded bg-gray-50 text-sm cursor-pointer">
                                                                        <div className="flex justify-between items-start">
                                                                            <div className="min-w-0">
                                                                                <div className="font-medium truncate">{name}</div>
                                                                                <div className="text-xs text-gray-500 truncate">{affiliation}</div>
                                                                            </div>
                                                                            <div className="text-xs ml-3">{e.timestamp ? new Date(e.timestamp).toLocaleString('ja-JP') : '-'}</div>
                                                                        </div>
                                                                        <div className="flex justify-between mt-1"><div className="text-xs">{e.type}</div><div className="text-sm">{fmt(e.amount)}</div></div>
                                                                    </div>
                                                                )})}
                                                            </div>
                                                        )}
                                                    </div>
                                                </div>
                                            ))}
                                        </div>
                                    </div>
                                )}

                                {summary && summary.error && <div className="text-red-600">{summary.error}</div>}
                            </div>
                        </details>
                        {/* New Collection Modal */}
                        {modalOpen && (
                            <div className="fixed inset-0 flex items-center justify-center modal-overlay p-4 z-[220]">
                                <div className="bg-white rounded-xl shadow-2xl w-full max-w-lg p-5">
                                    <div className="flex items-center justify-between mb-3"><h4 className="font-bold">集金登録</h4><button onClick={()=>setModalOpen(false)} className="text-sm text-gray-500">閉じる</button></div>
                                    <div className="space-y-3">
                                        <div>
                                            <label className="block text-xs text-gray-500 mb-1">ソース選択</label>
                                            <select disabled className="w-full border p-2 rounded" value={newMode}>
                                                <option value="sheet">スプレッドシートURLから登録</option>
                                                <option value="forms">Formsから選択</option>
                                            </select>
                                        </div>
                                        {formsLoading ? (
                                            <div className="flex flex-col items-center justify-center py-8"><div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mb-4"></div><div className="text-sm text-gray-600">読み込み中...</div></div>
                                        ) : (
                                        <>
                                        {newMode === 'sheet' ? (
                                            <div>
                                                <label className="block text-xs text-gray-500 mb-1">スプレッドシートURL</label>
                                                <input className="w-full border p-2 rounded" value={sheetUrl} onChange={e=>setSheetUrl(e.target.value)} />
                                            </div>
                                        ) : (
                                            <div>
                                                <label className="block text-xs text-gray-500 mb-1">Formsから選択</label>
                                                <select className="w-full border p-2 rounded" value={selectedFormRef} onChange={(e)=>{
                                                    const v = e.target.value; setSelectedFormRef(v);
                                                    const sel = (formsList || []).find(f => (f.spreadsheetRef||f.formUrl||'') === v);
                                                    if (sel) {
                                                        setSheetUrl(sel.spreadsheetRef || sel.formUrl || '');
                                                        if (!title) setTitle(sel.title || '');
                                                        setSelectedOrg(sel.inChargeOrg || '');
                                                        setSelectedDept(sel.inChargeDept || '');
                                                    } else {
                                                        setSheetUrl('');
                                                        setSelectedOrg(''); setSelectedDept('');
                                                    }
                                                }}>
                                                    <option value="">選択してください</option>
                                                    {formsList.map((f,idx)=>(<option key={idx} value={(f.spreadsheetRef || f.formUrl || '')}>{(f.title || f.spreadsheetRef || f.formUrl)}</option>))}
                                                </select>
                                                <div className={`mt-2 ${selectedFormRef ? 'opacity-80' : ''}`}>
                                                    <label className="block text-xs text-gray-500 mb-1">選択中のスプレッドシートURL</label>
                                                    <input className="w-full border p-2 rounded bg-gray-50" value={sheetUrl} disabled />
                                                </div>
                                            </div>
                                        )}

                                        <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                                            <div>
                                                <label className="block text-xs text-gray-500 mb-1">担当局</label>
                                                <select className="w-full border p-2 rounded" value={selectedOrg} onChange={e=>{ setSelectedOrg(e.target.value); setSelectedDept(''); }}>
                                                    <option value="">選択</option>
                                                    {(orgOptions && orgOptions.orgs ? Array.from(new Set(orgOptions.orgs)) : []).map((o2,idx)=>(<option key={idx} value={o2}>{o2}</option>))}
                                                </select>
                                            </div>
                                            <div>
                                                <label className="block text-xs text-gray-500 mb-1">担当部門</label>
                                                <select className="w-full border p-2 rounded" value={selectedDept} onChange={e=>setSelectedDept(e.target.value)} disabled={!selectedOrg}>
                                                    <option value="">選択</option>
                                                    {(orgOptions && orgOptions.deptMaster ? Array.from(new Set(orgOptions.deptMaster.filter(d=>d.org===selectedOrg).map(d=>d.dept))) : []).map((d2,idx)=>(<option key={idx} value={d2}>{d2}</option>))}
                                                </select>
                                            </div>
                                        </div>

                                        <div>
                                            <label className="block text-xs text-gray-500 mb-1">タイトル</label>
                                            <input className="w-full border p-2 rounded" value={title} onChange={e=>setTitle(e.target.value)} />
                                        </div>
                                        </>
                                        )}
                                        <div className="flex items-center justify-end gap-2 mt-3">
                                            {editingId ? (
                                                <button onClick={deleteThisCollection} disabled={savingCollection} className={`px-3 py-1 rounded text-white ${savingCollection ? 'bg-gray-400' : 'bg-red-600 hover:bg-red-700'}`}>
                                                    {savingCollection ? (<span><i className="fas fa-spinner fa-spin mr-2"></i>処理中...</span>) : '削除'}
                                                </button>
                                            ) : null}
                                            <button onClick={()=>{ setModalOpen(false); setEditingId(''); }} className="px-3 py-1 bg-gray-100 rounded">キャンセル</button>
                                            <button onClick={saveNew} disabled={savingCollection} className={`px-3 py-1 rounded text-white ${savingCollection ? 'bg-gray-400' : 'bg-blue-600 hover:bg-blue-700'}`}>
                                                {savingCollection ? (<span><i className="fas fa-spinner fa-spin mr-2"></i>登録中...</span>) : '登録'}
                                            </button>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        )}

                        <DialogModal isOpen={regDoneOpen} type={'notify'} message={regDoneMessage} onOk={() => { setRegDoneOpen(false); setRegDoneMessage(''); }} onCancel={() => { setRegDoneOpen(false); setRegDoneMessage(''); }} />

                        {/* Edit Select Modal */}
                        {editSelectOpen && (
                            <div className="fixed inset-0 flex items-center justify-center modal-overlay p-4 z-[225]">
                                <div className="bg-white rounded-xl shadow-2xl w-full max-w-2xl p-5 max-h-[80vh] overflow-auto">
                                    <div className="flex items-center justify-between mb-3">
                                        <h4 className="font-bold">集金一覧を選択</h4>
                                        <button onClick={()=>setEditSelectOpen(false)} className="text-sm text-gray-500">閉じる</button>
                                    </div>
                                    {editSelectLoading ? (
                                        <div className="flex items-center justify-center py-12"><div className="animate-spin rounded-full h-10 w-10 border-b-2 border-gray-900 mr-3"></div><div className="text-gray-600">読み込み中...</div></div>
                                    ) : (
                                    <div className="divide-y">
                                        {(collections || []).length === 0 && <div className="text-sm text-gray-500">登録済みの集金がありません。</div>}
                                        {(collections || []).map((c, idx) => (
                                            <div key={c.id || idx} className="py-3 flex items-center justify-between cursor-pointer" onClick={async ()=>{
                                                // open modal for editing with spinner while loading forms/options
                                                setEditingId(c.id || '');
                                                setTitle(c.title || '');
                                                setSheetUrl(c.spreadsheetUrl || '');
                                                setSelectedOrg(c.inChargeOrg || '');
                                                setSelectedDept(c.inChargeDept || '');
                                                setEditSelectOpen(false);
                                                setModalOpen(true);
                                                setFormsLoading(true);
                                                try {
                                                    const token = localStorage.getItem('slack_app_session');
                                                    const formsRes = await runGas('listFormDefinitions', token).catch(()=>({ success: true, items: [] }));
                                                    const opts = await runGas('getSearchOptions').catch(()=>({ orgs: [], deptMaster: [] }));
                                                    if (formsRes && formsRes.success) setFormsList(formsRes.items || []); else setFormsList([]);
                                                    if (opts) setOrgOptions({ orgs: opts.orgs || [], deptMaster: opts.deptMaster || [] });
                                                } catch(e) { setFormsList([]); }
                                                finally { setFormsLoading(false); }
                                            }}>
                                                <div className="min-w-0">
                                                    <div className="font-medium truncate">{c.title || c.spreadsheetUrl || '無題'}</div>
                                                    <div className="text-xs text-gray-500 truncate">{(c.inChargeOrg || '-') + (c.inChargeDept ? (' / ' + c.inChargeDept) : '')}</div>
                                                </div>
                                            </div>
                                        ))}
                                    </div>
                                    )}
                                </div>
                            </div>
                        )}

                        {selectCollectionOpen && (
                            <div className="fixed inset-0 flex items-center justify-center modal-overlay p-4 z-[225]">
                                <div className="bg-white rounded-xl shadow-2xl w-full max-w-2xl p-5 max-h-[80vh] overflow-auto">
                                    <div className="flex items-center justify-between mb-3">
                                        <h4 className="font-bold">集金を選択</h4>
                                        <button onClick={()=>setSelectCollectionOpen(false)} className="text-sm text-gray-500">閉じる</button>
                                    </div>
                                    {collectionsLoading ? (
                                        <div className="flex items-center justify-center py-12"><div className="animate-spin rounded-full h-10 w-10 border-b-2 border-gray-900 mr-3"></div><div className="text-gray-600">読み込み中...</div></div>
                                    ) : (
                                    <div className="divide-y">
                                        {(collections || []).length === 0 && <div className="text-sm text-gray-500">登録済みの集金がありません。</div>}
                                        {(collections || []).map((c, idx) => (
                                            <div key={c.id || idx} className="py-3 flex items-center justify-between cursor-pointer" onClick={async ()=>{
                                                setSelectedId(c.id || '');
                                                setSelectCollectionOpen(false);
                                                setSummary(null);
                                                await refreshSummary(c.id || '');
                                            }}>
                                                <div className="min-w-0">
                                                    <div className="font-medium truncate">{c.title || c.spreadsheetUrl || '無題'}</div>
                                                    <div className="text-xs text-gray-500 truncate">{(c.inChargeOrg || '-') + (c.inChargeDept ? (' / ' + c.inChargeDept) : '')}</div>
                                                </div>
                                            </div>
                                        ))}
                                    </div>
                                    )}
                                </div>
                            </div>
                        )}

                        {/* Accounting Modal (multi-step: org -> grade -> person -> pay) */}
                        {accountFlowOpen && (
                            <div className="fixed inset-0 flex items-center justify-center modal-overlay p-4 z-[230]">
                                <div className="bg-white rounded-xl shadow-2xl w-full max-w-2xl p-5 max-h-[85vh] overflow-auto relative">
                                    {(flowLoading && (flowStep === 'person' || flowStep === 'pay')) && (
                                        <div className="absolute inset-0 flex items-center justify-center bg-white/60 z-50">
                                            <div className="flex flex-col items-center">
                                                <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mb-3"></div>
                                                <div className="text-sm text-gray-600">読み込み中...</div>
                                            </div>
                                        </div>
                                    )}
                                    <div className="flex items-center justify-between mb-3">
                                        <h4 className="font-bold">会計</h4>
                                        <div className="flex items-center gap-2">
                                            <button onClick={()=>{ setAccountFlowOpen(false); setFlowStep('org'); }} className="text-sm text-gray-500">閉じる</button>
                                        </div>
                                    </div>

                                    {flowStep === 'org' && (
                                        <div>
                                            <div className="font-medium mb-2">所属局を選択</div>
                                            <div className="grid grid-cols-2 md:grid-cols-4 gap-2">
                                                <button onClick={()=>{
                                                    setFlowSelectedOrg('ALL');
                                                    setFlowStep('grade');
                                                    loadPeopleForOrg('').catch(()=>{ setFlowPeople([]); });
                                                }} className="text-left px-3 py-3 border rounded-lg hover:bg-gray-50 text-base md:text-sm">ALL</button>
                                                {(flowOrgList || []).map((o, oi) => (
                                                    <button key={oi} onClick={()=>{
                                                        setFlowSelectedOrg(o);
                                                        setFlowStep('grade');
                                                        loadPeopleForOrg(o).catch(()=>{ setFlowPeople([]); });
                                                    }} className="text-left px-3 py-3 border rounded-lg hover:bg-gray-50 text-base md:text-sm">{o}</button>
                                                ))}
                                                {(flowOrgList || []).length === 0 && <div className="text-xs text-gray-500">所属局が見つかりません。</div>}
                                            </div>
                                        </div>
                                    )}

                                    {flowStep === 'grade' && (
                                        <div>
                                            <div className="flex items-center justify-between mb-2">
                                                <div className="font-medium">学年を選択 — <span className="text-xs text-gray-500">{flowSelectedOrg}</span></div>
                                                <div>
                                                    <button onClick={()=>{ setFlowStep('org'); setFlowSelectedOrg(''); setFlowGrades([]); setFlowPeople([]); }} className="text-xs text-gray-400">戻る</button>
                                                </div>
                                            </div>
                                            <div className="flex flex-wrap gap-2">
                                                <button onClick={()=>{ setFlowSelectedGrade(''); setFlowStep('person'); }} className="px-3 py-3 border rounded-lg hover:bg-gray-50 text-base md:text-sm">ALL</button>
                                                {(flowGrades || []).map((g, gi) => (
                                                    <button key={gi} onClick={()=>{ setFlowSelectedGrade(g); setFlowStep('person'); }} className="px-3 py-3 border rounded-lg hover:bg-gray-50 text-base md:text-sm">{g}</button>
                                                ))}
                                                {(flowGrades || []).length === 0 && <div className="text-xs text-gray-500">学年データがありません。</div>}
                                            </div>
                                        </div>
                                    )}

                                    {flowStep === 'person' && (
                                        <div>
                                            <div className="flex items-center justify-between mb-2">
                                                <div className="font-medium">対象者一覧 — <span className="text-xs text-gray-500">{flowSelectedOrg} {flowSelectedGrade}</span></div>
                                                    <button onClick={()=>setFlowStep('grade')} className="text-xs text-gray-400 items-right">戻る</button>
                                            </div>
                                            <div className="space-y-2">
                                                <div className="flex items-center gap-2 items-right">
                                                    <label className="text-sm flex items-center gap-2"><input type="checkbox" checked={flowShowOnlyUnbalanced} onChange={(e)=>setFlowShowOnlyUnbalanced(e.target.checked)} /> 過不足なしを含む</label>
                                                </div>
                                                {(() => {
                                                    const filtered = (flowPeople || []).filter(p => (!flowSelectedGrade || p.grade === flowSelectedGrade));
                                                    const boxed = filtered.filter(p => {
                                                        // checkbox means "過不足なしを含む" (when checked include zero-diff persons)
                                                        if (flowShowOnlyUnbalanced) return true; // include all
                                                        // otherwise, hide persons whose expected - collected == 0
                                                        if (!summary || !summary.perPerson) return false;
                                                        const e = summary.perPerson.find(x => (x.email||'').toLowerCase() === (p.email||'').toLowerCase());
                                                        if (!e) return false;
                                                        return Number(e.expected || 0) !== Number(e.collected || 0);
                                                    });
                                                    if (boxed.length === 0) return (<div className="text-sm text-gray-500">対象者が見つかりません。</div>);
                                                    return boxed.map((p, pi) => {
                                                        const e = summary && summary.perPerson ? summary.perPerson.find(x => (x.email||'').toLowerCase() === (p.email||'').toLowerCase()) : null;
                                                        const expected = e ? Number(e.expected || 0) : 0;
                                                        const collected = e ? Number(e.collected || 0) : 0;
                                                        const diff = expected - collected;
                                                        const label = (() => {
                                                            if (!e || (e.entries || []).length === 0) return '未収';
                                                            if (diff > 0) return '不足';
                                                            if (diff < 0) return '返金';
                                                            return '差額';
                                                        })();
                                                        const isOverpaid = diff < 0;
                                                        return (
                                                            <div key={pi} role="button" tabIndex={0} onClick={() => selectPersonInFlow(p)} className="p-2 border rounded flex items-center justify-between cursor-pointer">
                                                                <div>
                                                                    <div className="font-medium">{p.name || p.email}</div>
                                                                    <div className="text-xs text-gray-500">{p.grade || '-'} / {p.field || '-'} / {p.department || '-'}</div>
                                                                </div>
                                                                <div className="text-right">
                                                                    <div className={`text-2xl font-bold ${isOverpaid ? 'text-red-600' : ''}`}>{fmt(Math.abs(diff))}</div>
                                                                    <div className={`text-xs mt-1 ${isOverpaid ? 'text-red-600' : 'text-gray-500'}`}>{label}</div>
                                                                </div>
                                                            </div>
                                                        );
                                                    });
                                                })()}
                                            </div>
                                        </div>
                                    )}

                                    {flowStep === 'pay' && flowSelectedPerson && (
                                        <div>
                                            {flowLoading && (
                                                <div className="absolute inset-0 flex items-center justify-center bg-white/60 z-50">
                                                    <div className="flex flex-col items-center">
                                                        <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mb-3"></div>
                                                        <div className="text-sm text-gray-600">読み込み中...</div>
                                                    </div>
                                                </div>
                                            )}
                                            <div className="mb-3">
                                                {(() => {
                                                    const entry = summary && summary.perPerson ? summary.perPerson.find(x => (x.email||'').toLowerCase() === (flowSelectedPerson.email||'').toLowerCase()) : null;
                                                    const expected = entry ? Number(entry.expected || 0) : 0;
                                                    const collected = entry ? Number(entry.collected || 0) : 0;
                                                    const diff = expected - collected;
                                                    const label = (!entry || (entry.entries || []).length === 0) ? '請求額' : (diff > 0 ? '不足' : (diff < 0 ? '返金' : '差額'));
                                                    const alert = (label === '不足' || label === '返金');
                                                    return (
                                                        <>
                                                            <div className="flex justify-center mb-3">
                                                                <div className="transform rotate-180 flex flex-col gap-1 bg-black text-white w-full text-center m-5">
                                                                    <div className="text-xl font-bold items-center">{flowSelectedPerson.name || flowSelectedPerson.email} さん</div>
                                                                    <table className="w-full">
                                                                        <tr>
                                                                            <td>
                                                                                <div className={`text-lg ${alert ? 'text-red-600 font-bold' : 'text-white-600'}`}>{label}</div>
                                                                            </td>
                                                                            <td>
                                                                                <div className={`text-3xl font-extrabold ${alert ? 'text-red-600' : ''}`}>{fmt(Math.abs(diff))} 円</div>
                                                                            </td>
                                                                        </tr>
                                                                    </table>
                                                                    <div className="text-sm text-gray-200 mt-1">集金額: {fmt(expected)} 円</div>
                                                                </div>
                                                            </div>
                                                            <br />
                                                            <hr />
                                                            <br />
                                                            <div className="text-center space-y-1 mb-3">
                                                                <div className="text-lg font-bold">{flowSelectedPerson.name || flowSelectedPerson.email}</div>
                                                                <table className="w-full">
                                                                    <tr>
                                                                        <td>
                                                                            <div className={`text-lg ${alert ? 'text-red-600 font-bold' : 'text-gray-600'}`}>{label}</div>
                                                                        </td>
                                                                        <td>
                                                                            <div className={`text-3xl font-extrabold ${alert ? 'text-red-600' : ''}`}>{fmt(Math.abs(diff))} 円</div>
                                                                        </td>
                                                                    </tr>
                                                                </table>
                                                                <div className="text-sm text-gray-600">集金額: {fmt(expected)} 円</div>
                                                            </div>
                                                        </>
                                                    );
                                                })()}
                                                <div className="mb-3">
                                                    <label className="block text-xs text-gray-500 mb-1">受取額</label>
                                                    <div className="flex items-center gap-2">
                                                        <input type="number" inputMode="decimal" value={flowReceiveAmount || 0} onChange={(e)=>setFlowReceiveAmount(Number(e.target.value || 0))} className="w-full border p-2 rounded" />
                                                        <button type="button" onClick={()=>setFlowReceiveAmount(Number(flowReceiveAmount || 0) * -1)} className="px-3 py-2 rounded border bg-gray-100 text-gray-700">±</button>
                                                    </div>
                                                </div>
                                                <div>
                                                    {(() => {
                                                        const entry = summary && summary.perPerson ? summary.perPerson.find(x => (x.email||'').toLowerCase() === (flowSelectedPerson.email||'').toLowerCase()) : null;
                                                        const expected = entry ? Number(entry.expected || 0) : 0;
                                                        const collected = entry ? Number(entry.collected || 0) : 0;
                                                        const diff = expected - collected; // positive: owes money; negative: overpaid
                                                        const received = Number(flowReceiveAmount || 0);

                                                        if (received < 0) {
                                                            return (
                                                                <div className="flex">
                                                                    <button onClick={()=>doFlowReceive('debt','返金')} className="w-full text-lg py-3 bg-gray-200 text-gray-900 rounded">{fmt(Math.abs(received))}円 返金</button>
                                                                </div>
                                                            );
                                                        }

                                                        if (diff >= 0 && received <= diff) {
                                                            return (
                                                                <div className="flex">
                                                                    <button onClick={()=>doFlowReceive('receive')} className="w-full text-lg py-4 bg-blue-600 text-white rounded-lg">{fmt(received)}円 受領</button>
                                                                </div>
                                                            );
                                                        }

                                                        const changeAmt = Math.max(0, received - Number(flowReceiveInitialAmount || 0));
                                                        return (
                                                            <div className="space-y-2">
                                                                <button onClick={()=>doFlowReceive('change')} className="w-full text-lg py-3 bg-gray-200 text-gray-900 rounded">{fmt(changeAmt)}円 おつり</button>
                                                                <button onClick={()=>doFlowReceive('debt','過不足')} className="w-full text-lg py-3 bg-yellow-200 text-gray-900 rounded">{fmt(Math.abs(received - Number(flowReceiveInitialAmount || 0)))}円 過不足扱い（後日精算）</button>
                                                            </div>
                                                        );
                                                    })()}
                                                </div>
                                                <div className="mt-3 text-right"><button onClick={()=>setFlowStep('person')} className="text-xs text-gray-400">戻る</button></div>
                                            </div>
                                        </div>
                                    )}

                                </div>
                            </div>
                        )}

                        {/* Quick single-person Accounting Modal (legacy) */}
                        {accountingOpen && accountTarget && (
                            <div className="fixed inset-0 flex items-center justify-center modal-overlay p-4 z-[235]">
                                <div className="bg-white rounded-xl shadow-2xl w-full max-w-sm p-5">
                                    <div className="flex items-center justify-between mb-3"><h4 className="font-bold">支払い</h4><button onClick={()=>setAccountingOpen(false)} className="text-sm text-gray-500">閉じる</button></div>
                                                {flowLoading && (
                                                                            <div className="absolute inset-0 flex items-center justify-center bg-white/60 z-50">
                                                                                <div className="flex flex-col items-center">
                                                                                    <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mb-3"></div>
                                                                                    <div className="text-sm text-gray-600">処理中…</div>
                                                                                </div>
                                                                            </div>
                                                                        )}
                                                                        <div className="text-center mb-3">
                                                                            {(() => {
                                                                                const diff = Number(accountTarget.expected || 0) - Number(accountTarget.collected || 0);
                                                                                const label = (diff > 0) ? '不足' : (diff < 0 ? '返金' : '差額');
                                                                                const alert = (label === '不足' || label === '返金');
                                                                                return (
                                                                                    <>
                                                                                        <div className="flex justify-center mb-3">
                                                                                            <div className="transform rotate-180 flex flex-col gap-1 bg-black text-white w-full text-center m-5">
                                                                                                <div className="text-xl font-bold items-center">{flowSelectedPerson.name || flowSelectedPerson.email} さん</div>
                                                                                                <table className="w-full">
                                                                                                    <tr>
                                                                                                        <td>
                                                                                                            <div className={`text-l ${alert ? 'text-red-600 font-bold' : 'text-white-600'}`}>{label}</div>
                                                                                                        </td>
                                                                                                        <td>
                                                                                                            <div className={`text-2xl font-bold ${alert ? 'text-red-600' : ''}`}>{fmt(Math.abs(diff))} 円</div>
                                                                                                        </td>
                                                                                                    </tr>
                                                                                                </table>
                                                                                            </div>
                                                                                        </div>
                                                                                        <br />
                                                                                        <hr />
                                                                                        <br />
                                                                                        <div className="space-y-1">
                                                                                            <div className="text-lg font-bold">{(accountTarget && (accountTarget.name || accountTarget.email)) || '-'}</div>
                                                                                            <table className="w-full">
                                                                                                <tr>
                                                                                                    <td>
                                                                                                        <div className={`text-xs ${alert ? 'text-red-600 font-bold' : 'text-gray-600'}`}>{label}</div>
                                                                                                    </td>
                                                                                                    <td>
                                                                                                        <div className={`text-2xl font-bold ${alert ? 'text-red-600' : ''}`}>{fmt(Math.abs(diff))} 円</div>
                                                                                                    </td>
                                                                                                </tr>
                                                                                            </table>
                                                                                        </div>
                                                                                    </>
                                                                                );
                                                                            })()}
                                                                        </div>
                                                                        <div className="mb-3">
                                                                            <label className="block text-xs text-gray-500 mb-1">受取額</label>
                                                                            <div className="flex items-center gap-2">
                                                                                <input id="collect-input-amount" type="number" inputMode="decimal" value={quickReceiveAmount || 0} onChange={(e)=>setQuickReceiveAmount(Number(e.target.value || 0))} className="w-full border p-2 rounded" />
                                                                                <button type="button" onClick={()=>setQuickReceiveAmount(Number(quickReceiveAmount || 0) * -1)} className="px-3 py-2 rounded border bg-gray-100 text-gray-700">±</button>
                                                                            </div>
                                                                        </div>
                                                                        <div className="flex items-center gap-2 justify-end">
                                                                            {(() => {
                                                                                const remaining = Number(accountTarget.expected || 0) - Number(accountTarget.collected || 0);
                                                                                const received = Number(quickReceiveAmount || 0);
                                                                                if (received < 0) {
                                                                                    return (
                                                                                        <button className="px-3 py-1 bg-gray-200 rounded" onClick={()=>doReceive('debt','返金')}>返金</button>
                                                                                    );
                                                                                }
                                                                                if (remaining >= 0 && received <= remaining) {
                                                                                    return (
                                                                                        <button className="px-3 py-1 bg-blue-600 text-white rounded" onClick={()=>doReceive('receive')}>受領</button>
                                                                                    );
                                                                                }
                                                                                return (
                                                                                    <>
                                                                                        <button className="px-3 py-1 bg-gray-200 rounded" onClick={()=>doReceive('change')}>{fmt(Math.max(0, received - Number(quickReceiveInitialAmount || 0)))}円 おつり</button>
                                                                                        <button className="px-3 py-1 bg-yellow-200 rounded" onClick={()=>doReceive('debt','過不足')}>{fmt(Math.abs(received - Number(quickReceiveInitialAmount || 0)))}円 過不足扱い（後日返金）</button>
                                                                                    </>
                                                                                );
                                                                            })()}
                                                                        </div>
                                </div>
                            </div>
                        )}
                        {accountReceiptOpen && (
                            <div className="fixed inset-0 flex items-center justify-center modal-overlay p-4 z-[240]">
                                <div className="bg-white rounded-xl shadow-2xl w-full max-w-md p-5 max-h-[85vh] overflow-auto">
                                    <div className="flex items-center justify-between mb-3">
                                        <h4 className="font-bold">完了</h4>
                                    </div>
                                    {accountReceiptData && (
                                        <>
                                        <div className="text-lg text-white font-bold space-y-3 transform rotate-180 bg-black">
                                            <table className="w-full text-lg">
                                                <tbody>
                                                    <tr>
                                                        <td className="py-1">集金額</td>
                                                        <td className="py-1 text-right">{fmt(accountReceiptData.expected)}円</td>
                                                    </tr>
                                                </tbody>
                                            </table>
                                            <hr />
                                            {accountReceiptData.historyEntries && accountReceiptData.historyEntries.length > 0 && (
                                                <>
                                                    <table className="w-full text-lg">
                                                        <tbody>
                                                            {accountReceiptData.historyEntries.map((h, i) => (
                                                                <tr key={i}>
                                                                    <td className="py-1">{h.type || '-'}</td>
                                                                    <td className="py-1 text-right">{fmt(h.amount)}円</td>
                                                                </tr>
                                                            ))}
                                                        </tbody>
                                                    </table>
                                                    <hr />
                                                </>
                                            )}
                                            <table className="w-full text-lg">
                                                <tbody>
                                                    <tr>
                                                        <td className="py-1">{accountReceiptData.initialReceive < 0 ? '返金額' : '請求額'}</td>
                                                        <td className="py-1 text-right">{fmt(accountReceiptData.initialReceive)}円</td>
                                                    </tr>
                                                    <tr>
                                                        <td className="py-1">受取額</td>
                                                        <td className="py-1 text-right">{fmt(accountReceiptData.received)}円</td>
                                                    </tr>
                                                </tbody>
                                            </table>
                                            <hr />
                                            <table className="w-full text-lg">
                                                <tbody>
                                                    <tr>
                                                        <td className="py-1">{accountReceiptData.changeLabel}</td>
                                                        <td className="py-1 text-right">{fmt(accountReceiptData.changeAmount)}円</td>
                                                    </tr>
                                                </tbody>
                                            </table>
                                        </div>
                                        <br />
                                        <br />
                                        <br />
                                        <div className="text-lg text-gray-700 space-y-3">
                                            <table className="w-full text-lg">
                                                <tbody>
                                                    <tr>
                                                        <td className="py-1">{accountReceiptData.initialReceive < 0 ? '返金額' : '請求額'}</td>
                                                        <td className="py-1 text-right">{fmt(accountReceiptData.initialReceive)}円</td>
                                                    </tr>
                                                    <tr>
                                                        <td className="py-1">受取額</td>
                                                        <td className="py-1 text-right">{fmt(accountReceiptData.received)}円</td>
                                                    </tr>
                                                </tbody>
                                            </table>
                                            <hr />
                                            <table className="w-full text-lg">
                                                <tbody>
                                                    <tr>
                                                        <td className="py-1">{accountReceiptData.changeLabel}</td>
                                                        <td className="py-1 text-right">{fmt(accountReceiptData.changeAmount)}円</td>
                                                    </tr>
                                                </tbody>
                                            </table>
                                        </div>
                                        </>
                                    )}
                                    {!accountReceiptLoading && (
                                        <div className="mt-4 flex justify-end">
                                            <button onClick={() => {
                                                setAccountReceiptOpen(false);
                                                setAccountReceiptData(null);
                                                openAccountFlow();
                                            }} className="px-4 py-2 rounded bg-blue-600 text-white">OK</button>
                                        </div>
                                    )}
                                </div>
                            </div>
                        )}
                    </div>
                </div>
            );
        }

        function RosterTab({ user }) {
            const [isAdmin, setIsAdmin] = useState(false);
            const [isAdminLoading, setIsAdminLoading] = useState(true);
            const [fields] = useState([
                { key: '氏名', label: '氏名', adminOnly: false },
                { key: 'Name', label: 'Name', adminOnly: false },
                { key: '学籍番号', label: '学籍番号', adminOnly: true },
                { key: '学年', label: '学年', adminOnly: false },
                { key: '分野', label: '分野', adminOnly: false },
                { key: 'メールアドレス', label: 'メールアドレス', adminOnly: false },
                { key: '電話番号', label: '電話番号', adminOnly: true },
                { key: '生年月日', label: '生年月日', adminOnly: true },
                { key: '所属局1', label: '所属局1', adminOnly: false },
                { key: '所属部門1', label: '所属部門1', adminOnly: false },
                { key: '役職1', label: '役職1', adminOnly: false },
                // atomic keys for 2-5 must exist so grouping logic can create the grouped checkbox
                { key: '所属局2', label: '所属局2', adminOnly: false },
                { key: '所属部門2', label: '所属部門2', adminOnly: false },
                { key: '役職2', label: '役職2', adminOnly: false },
                { key: '所属局3', label: '所属局3', adminOnly: false },
                { key: '所属部門3', label: '所属部門3', adminOnly: false },
                { key: '役職3', label: '役職3', adminOnly: false },
                { key: '所属局4', label: '所属局4', adminOnly: false },
                { key: '所属部門4', label: '所属部門4', adminOnly: false },
                { key: '役職4', label: '役職4', adminOnly: false },
                { key: '所属局5', label: '所属局5', adminOnly: false },
                { key: '所属部門5', label: '所属部門5', adminOnly: false },
                { key: '役職5', label: '役職5', adminOnly: false },
                { key: '出身校', label: '出身校', adminOnly: true },
                { key: '退局', label: '退局ステータス', adminOnly: true },
                { key: '次年度継続', label: '次年度継続ステータス', adminOnly: true },
                { key: 'Admin', label: '管理者権限', adminOnly: true }
            ]);

            const [selected, setSelected] = useState(new Set());
            const [selectAll, setSelectAll] = useState(false);
            // atomic keys and display grouping for所属2-5
            const availableAtomicKeys = fields.map(f=>f.key);
            const displayFields = (()=>{
                const out = [];
                let skippingGroup = false;
                for (const f of fields) {
                    // 所属2-5グループの開始を検出
                    if (f.key === '所属局2' && !skippingGroup) {
                        out.push({ key: '所属2-5', label: '所属局・部門・役職（2～5）', isGroup: true, adminOnly: false });
                        skippingGroup = true;
                        continue;
                    }
                    // 所属2-5グループに属するアイテムをスキップ
                    if (/^(所属局[2-5]|所属部門[2-5]|役職[2-5])$/.test(f.key)) {
                        continue;
                    }
                    // 他のアイテム（所属1を含む）は通常追加
                    out.push({ ...f, isGroup: false });
                }
                // 重複除去（念のため）
                const seen = new Set();
                return out.filter(x => { if (seen.has(x.key)) return false; seen.add(x.key); return true; });
            })();
            const [modalOpen, setModalOpen] = useState(false);
            const [filename, setFilename] = useState('');
            const [busy, setBusy] = useState(false);
            const [previewData, setPreviewData] = useState(null); // { filename, headers:[], rows:[] }

            // 対象選択（全員 or 局ごと）
            const [targetType, setTargetType] = useState('all'); // 'all' or 'orgs'
            const [orgMatchMode, setOrgMatchMode] = useState('mainOnly'); // 'mainOnly' or 'allAffiliations'
            const [orgList, setOrgList] = useState([]); // [{name, depts: []}]
            const [orgChecked, setOrgChecked] = useState({}); // {orgName: true}
            const [deptChecked, setDeptChecked] = useState({}); // {orgName: {deptName: true}}
            const [gradeOptions, setGradeOptions] = useState([]);
            const [fieldOptions, setFieldOptions] = useState([]);
            const [surveys, setSurveys] = useState([]);
            const [surveysLoading, setSurveysLoading] = useState(false);
            const [surveyFetchError, setSurveyFetchError] = useState('');
            const safeSetSurveyError = (v) => { try { if (typeof setSurveyFetchError === 'function') setSurveyFetchError(v); } catch(e){} };
            const [sourceType, setSourceType] = useState(''); // '' | 'survey' | 'collection'
            const [selectedSurveyRef, setSelectedSurveyRef] = useState('');
            const [collections, setCollections] = useState([]);
            const [collectionsLoading, setCollectionsLoading] = useState(false);
            const [collectionsError, setCollectionsError] = useState('');
            const [selectedCollectionId, setSelectedCollectionId] = useState('');
            const [selectedGrades, setSelectedGrades] = useState([]);
            const [selectedFieldFilters, setSelectedFieldFilters] = useState([]);
            const [statusFilter, setStatusFilter] = useState('active'); // 'active' (在籍), 'retired', 'all'

            useEffect(() => {
                const token = localStorage.getItem('slack_app_session');
                setIsAdminLoading(true);
                runGas('getUserProfile', token)
                    .then(p => { setIsAdmin(!!p.isAdmin); })
                    .catch(()=>{})
                    .finally(()=> setIsAdminLoading(false));

                // デフォルトファイル名（Tokyo timezone）
                const now = new Date();
                const parts = now.toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo', hour12: false });
                const nums = parts.replace(/[^0-9]/g, '');
                const defName = 'list_' + nums.slice(0,14);
                setFilename(defName + '.csv');

                // 取得して局・部門を構築
                runGas('getSearchOptions').then(o => {
                    try {
                        setGradeOptions(o.grades || []);
                        setFieldOptions(o.fields || []);
                        const orgs = (o && o.orgs) ? Array.from(new Set(o.orgs)) : [];
                        const deptMaster = o && o.deptMaster ? o.deptMaster : [];
                        const map = orgs.map(org => ({ name: org, depts: [] }));
                        const m = {};
                        orgs.forEach(org => { m[org] = new Set(); });
                        deptMaster.forEach(d => { if (d.org && d.dept && m[d.org]) m[d.org].add(d.dept); });
                        const final = orgs.map(org => ({ name: org, depts: Array.from(m[org] || []) }));
                        setOrgList(final);
                        const oc = {}; const dc = {};
                        final.forEach(o2 => { oc[o2.name] = false; dc[o2.name] = {}; o2.depts.forEach(d=>dc[o2.name][d]=false); });
                        setOrgChecked(oc); setDeptChecked(dc);
                    } catch (e) { console.warn(e); }
                }).catch(()=>{});

                // fetch surveys for dropdown (with loading state)
                        const fetchSurveys = async () => {
                    safeSetSurveyError('');
                    setSurveysLoading(true);
                    try {
                        const list = await runGas('listSurveys', token);
                        if (Array.isArray(list)) {
                                    const mapped = list.map(s => ({ title: s.title || s.spreadsheetId || s.formUrl || s.spreadsheetUrl, ref: s.spreadsheetId || s.spreadsheetUrl || s.formUrl || '', inChargeOrg: s.inChargeOrg || '', inChargeDept: s.inChargeDept || '' }));
                                    setSurveys(mapped.filter(s=>s.ref));
                        } else if (list && list.message) {
                            setSurveys([]);
                            safeSetSurveyError(String(list.message));
                        } else {
                            setSurveys([]);
                        }
                    } catch (e) {
                        console.warn('listSurveys error', e);
                        setSurveys([]);
                        safeSetSurveyError(e && e.message ? e.message : String(e));
                    } finally {
                        setSurveysLoading(false);
                    }
                };
                fetchSurveys();
            }, []);

            useEffect(() => {
                if (sourceType !== 'collection') return;
                const token = localStorage.getItem('slack_app_session');
                setCollectionsLoading(true);
                setCollectionsError('');
                runGas('listCollections', token)
                    .then(list => { setCollections(list || []); })
                    .catch(e => { setCollections([]); setCollectionsError(e && e.message ? e.message : String(e)); })
                    .finally(() => { setCollectionsLoading(false); });
            }, [sourceType]);

            const toggleField = (k, adminOnly) => {
                if (adminOnly && !isAdmin) return;
                const s = new Set(Array.from(selected));
                // handle grouped key for 所属2-5: toggle all atomic members
                if (k === '所属2-5') {
                    const atomic = [];
                    for (let idx = 2; idx <= 5; idx++) {
                        atomic.push(`所属局${idx}`);
                        atomic.push(`所属部門${idx}`);
                        atomic.push(`役職${idx}`);
                    }
                    const allAtomicSelected = atomic.every(a => s.has(a));
                    if (allAtomicSelected) {
                        atomic.forEach(a => s.delete(a));
                    } else {
                        atomic.forEach(a => s.add(a));
                    }
                } else {
                    if (s.has(k)) s.delete(k); else s.add(k);
                }
                setSelected(s);
                // update selectAll state
                const available = fields.filter(f=>!(f.adminOnly && !isAdmin)).map(f=>f.key);
                const allSelected = available.every(a=>s.has(a));
                setSelectAll(allSelected);
            };

            const toggleGradeFilter = (grade) => {
                setSelectedGrades(prev => {
                    const set = new Set(prev || []);
                    if (set.has(grade)) set.delete(grade); else set.add(grade);
                    return Array.from(set);
                });
            };

            const toggleFieldFilter = (field) => {
                setSelectedFieldFilters(prev => {
                    const set = new Set(prev || []);
                    if (set.has(field)) set.delete(field); else set.add(field);
                    return Array.from(set);
                });
            };

            const clearGradeFilter = () => setSelectedGrades([]);
            const clearFieldFilter = () => setSelectedFieldFilters([]);

            const toggleSelectAll = () => {
                const available = fields.filter(f=>!(f.adminOnly && !isAdmin)).map(f=>f.key);
                if (selectAll) {
                    setSelected(new Set());
                    setSelectAll(false);
                } else {
                    setSelected(new Set(available));
                    setSelectAll(true);
                }
            };

            useEffect(()=>{
                const available = fields.filter(f=>!(f.adminOnly && !isAdmin)).map(f=>f.key);
                const allSelected = available.length>0 && available.every(a=>selected.has(a));
                setSelectAll(allSelected);
            }, [isAdmin, selected]);

            // available keys for UI disabling
            const availableKeys = fields.filter(f=>!(f.adminOnly && !isAdmin)).map(f=>f.key);

            const toggleOrg = (orgName) => {
                const oc = { ...orgChecked, [orgName]: !orgChecked[orgName] };
                setOrgChecked(oc);
                // toggle all depts
                const dc = { ...deptChecked };
                const map = dc[orgName] || {};
                const newMap = {};
                Object.keys(map).forEach(k => { newMap[k] = !orgChecked[orgName]; });
                dc[orgName] = newMap;
                setDeptChecked(dc);
            };

            const toggleDept = (orgName, deptName) => {
                const dc = { ...deptChecked };
                if (!dc[orgName]) dc[orgName] = {};
                dc[orgName][deptName] = !dc[orgName][deptName];
                setDeptChecked(dc);
            };

            const openExportModal = async () => {
                // open modal immediately (shows loading) and generate preview; close modal on failure
                setModalOpen(true);
                const ok = await generatePreview();
                if (!ok) setModalOpen(false);
            };

                const buildExportPayload = () => {
                const arr = Array.from(selected);
                const expanded = [];
                for (const v of arr) {
                    if (v === '所属2-5') {
                        for (let idx = 2; idx <= 5; idx++) {
                            expanded.push(`所属局${idx}`);
                            expanded.push(`所属部門${idx}`);
                            expanded.push(`役職${idx}`);
                        }
                    } else {
                        expanded.push(v);
                    }
                }
                let filter = { type: 'all' };
                if (targetType === 'orgs') {
                    const sels = [];
                    orgList.forEach(o => {
                        const allDepts = (o.depts || []);
                        const checkedDepts = (deptChecked[o.name]) ? Object.keys(deptChecked[o.name]).filter(d=>deptChecked[o.name][d]) : [];
                        if (orgChecked[o.name]) {
                            if (checkedDepts.length === 0 || checkedDepts.length === allDepts.length) {
                                sels.push({ org: o.name, dept: '' });
                            } else { checkedDepts.forEach(d=>sels.push({ org: o.name, dept: d })); }
                        } else {
                            if (checkedDepts.length > 0) checkedDepts.forEach(d=>sels.push({ org: o.name, dept: d }));
                        }
                    });
                    filter = { type: 'orgs', orgMatchMode: orgMatchMode, selections: sels };
                }
                filter.status = statusFilter || 'active';
                if (selectedGrades && selectedGrades.length > 0) filter.grade = selectedGrades.slice();
                if (selectedFieldFilters && selectedFieldFilters.length > 0) filter.field = selectedFieldFilters.slice();

                const now = new Date();
                const fmt = new Intl.DateTimeFormat('ja-JP', { timeZone: 'Asia/Tokyo', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false });
                const parts = fmt.formatToParts(now);
                const get = (type) => (parts.find(p=>p.type===type)||{}).value || '';
                const ts = get('year') + get('month') + get('day') + get('hour') + get('minute') + get('second');
                const filenameAtPress = 'list_' + ts + '.csv';
                return {
                    expanded,
                    filter,
                    filenameAtPress,
                    surveyRef: (sourceType === 'survey' && typeof selectedSurveyRef !== 'undefined') ? selectedSurveyRef : '',
                    collectionId: (sourceType === 'collection' && typeof selectedCollectionId !== 'undefined') ? selectedCollectionId : ''
                };
            };

            const fetchCsvText = async (payload) => {
                const token = localStorage.getItem('slack_app_session');
                    const res = await runGas('createRosterCsv', token, { selectedFields: payload.expanded, filter: payload.filter, filename: payload.filenameAtPress, surveyRef: payload.surveyRef, collectionId: payload.collectionId });
                if (!res || !res.success) throw new Error(res && res.message ? res.message : '作成に失敗しました');
                let csvText = '';
                const fname = payload.filenameAtPress || res.filename || 'list.csv';
                const excelB64 = res.excelBase64 || null;
                if (res.csvBase64) {
                    const b64 = res.csvBase64;
                    const bin = atob(b64);
                    const len = bin.length;
                    const u8 = new Uint8Array(len);
                    for (let i = 0; i < len; i++) u8[i] = bin.charCodeAt(i);
                    try { csvText = (new TextDecoder('utf-8')).decode(u8); } catch(e) { csvText = ''; }
                } else {
                    csvText = res.csv || '';
                }
                return { fname, csvText, excelB64 };
            };

            const generatePreview = async () => {
                if (selected.size === 0) { alert('出力項目を選択してください'); return false; }
                setBusy(true);
                try {
                    const payload = buildExportPayload();
                    const { fname, csvText, excelB64 } = await fetchCsvText(payload);
                    const text = (csvText || '').replace(/^\uFEFF/, '');
                    const parsed = parseCsv(text);
                    // normalize birthday column to YYYY/MM/DD (zero-padded)
                    const bIdx = parsed.headers.findIndex(h => h === '生年月日');
                    if (bIdx >= 0) {
                        parsed.rows = parsed.rows.map(r => {
                            const v = r[bIdx] || '';
                            if (!v) return r;
                            r[bIdx] = formatPreviewBirthday(v);
                            return r;
                        });
                    }
                    setPreviewData({ filename: fname, headers: parsed.headers, rows: parsed.rows, csvText: csvText, excelBase64: excelB64 });
                    return true;
                } catch (e) { alert(e.message || e); setPreviewData(null); return false; }
                finally { setBusy(false); }
            };

            const doExportDownload = async () => {
                // download using the preview csvText if available, otherwise generate then download
                if (!previewData || !previewData.csvText) {
                    await generatePreview();
                    if (!previewData || !previewData.csvText) return;
                }
                try {
                    const fnameRaw = filename || previewData.filename || 'list.csv';
                    // prefer excel if available
                    if (previewData.excelBase64) {
                        const b64 = previewData.excelBase64;
                        const bin = atob(b64);
                        const len = bin.length;
                        const u8 = new Uint8Array(len);
                        for (let i = 0; i < len; i++) u8[i] = bin.charCodeAt(i);
                        const blob = new Blob([u8], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                        let fname = fnameRaw;
                        if (fname.toLowerCase().endsWith('.csv')) fname = fname.replace(/\.csv$/i, '.xlsx');
                        else if (!fname.toLowerCase().endsWith('.xlsx')) fname = fname + '.xlsx';
                        const url = URL.createObjectURL(blob);
                        const a = document.createElement('a'); a.href = url; a.download = fname; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
                        setModalOpen(false);
                        return;
                    }

                    // fallback to CSV download
                    const fname = fnameRaw;
                    const csv = previewData.csvText || '';
                    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
                    const url = URL.createObjectURL(blob);
                    const a = document.createElement('a'); a.href = url; a.download = fname; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
                    setModalOpen(false);
                } catch (e) { alert(e.message || e); }
            };

            return (
                <div className="flex flex-col h-full p-4 md:p-6 bg-gray-50">
                    <div className="max-w-4xl mx-auto bg-white rounded-lg shadow-sm p-6 h-full overflow-y-auto relative">
                            {surveysLoading && (
                                <div className="absolute inset-0 bg-white bg-opacity-80 flex items-center justify-center z-50">
                                    <div className="flex justify-center items-center text-gray-500"><i className="fas fa-spinner fa-spin mr-2"></i>読み込み中...</div>
                                </div>
                            )}
                        <h3 className="text-lg font-bold mb-3">名簿出力</h3>
                        <p className="text-sm text-gray-600 mb-4">出力したい項目を選択し、出力対象を設定して「出力」を押してください。</p>

                        <div className="mb-4">
                            <h4 className="font-medium mb-2">出力項目</h4>
                            <div className="grid grid-cols-1 md:grid-cols-3 gap-2">
                                <label className={`flex items-center p-2 border rounded ${availableAtomicKeys.length===0 ? 'bg-gray-100 text-gray-400 cursor-not-allowed' : 'bg-white'}`} onClick={availableAtomicKeys.length===0 ? undefined : toggleSelectAll} title={availableAtomicKeys.length===0 ? '選択可能な項目がありません' : 'すべて選択'}>
                                    <input type="checkbox" checked={selectAll} disabled={availableAtomicKeys.length===0} readOnly className="mr-3 w-4 h-4" />
                                    <span className="text-sm font-medium">すべて選択</span>
                                </label>
                                {displayFields.map(f => (
                                    f.isGroup ? (
                                        <label key={f.key} className={`flex items-center p-2 border rounded ${f.adminOnly && !isAdmin ? 'bg-gray-100 text-gray-400 cursor-not-allowed' : 'bg-white'}`} onClick={f.adminOnly && !isAdmin ? undefined : () => toggleField(f.key, f.adminOnly, true)} title={f.adminOnly && !isAdmin ? '管理者のみ選択可能' : ''} aria-disabled={f.adminOnly && !isAdmin}>
                                            {/* group checked if all atomic members selected */}
                                            <input type="checkbox" checked={(() => {
                                                const atomic = [];
                                                for (let idx = 2; idx <= 5; idx++) { atomic.push(`所属局${idx}`); atomic.push(`所属部門${idx}`); atomic.push(`役職${idx}`); }
                                                return atomic.every(a => selected.has(a));
                                            })()} disabled={f.adminOnly && !isAdmin} readOnly className="mr-3 w-4 h-4" />
                                            <span className="text-sm">{f.label}</span>
                                        </label>
                                    ) : (
                                        <label key={f.key} className={`flex items-center p-2 border rounded ${f.adminOnly && !isAdmin ? 'bg-gray-100 text-gray-400 cursor-not-allowed' : 'bg-white'}`} onClick={f.adminOnly && !isAdmin ? undefined : () => toggleField(f.key, f.adminOnly)} title={f.adminOnly && !isAdmin ? '管理者のみ選択可能' : ''} aria-disabled={f.adminOnly && !isAdmin}>
                                            <input type="checkbox" checked={selected.has(f.key)} disabled={f.adminOnly && !isAdmin} readOnly className="mr-3 w-4 h-4" />
                                            <span className="text-sm">{f.label}</span>
                                        </label>
                                    )
                                ))}
                            </div>
                        </div>

                        <div className="mb-4">
                            <h4 className="font-medium mb-2">ソース選択</h4>
                            <div className="space-y-3">
                                <select className="w-full border p-2 rounded" value={sourceType} onChange={(e)=>{
                                    const v = e.target.value;
                                    setSourceType(v);
                                    if (v !== 'survey') setSelectedSurveyRef('');
                                    if (v !== 'collection') setSelectedCollectionId('');
                                }}>
                                    <option value="">なし</option>
                                    <option value="survey">アンケート</option>
                                    <option value="collection">集金</option>
                                </select>

                                {sourceType === 'survey' && (
                                    <div>
                                        <select className="w-full border p-2 rounded" value={selectedSurveyRef} onChange={(e)=>setSelectedSurveyRef(e.target.value)}>
                                            <option value="">アンケートを選択（出力時に名簿の後に追加）</option>
                                            {surveys && surveys.length > 0 ? surveys.map(s => <option key={s.ref} value={s.ref}>{s.title + ( (s.inChargeOrg || s.inChargeDept) ? (' — ' + ((s.inChargeOrg || '-') + (s.inChargeDept ? (' / ' + s.inChargeDept) : '')) ) : '' )}</option>) : null}
                                        </select>
                                        {surveys.length === 0 && !surveysLoading && (
                                            <div className="mt-2 text-xs text-gray-500">
                                                {surveyFetchError ? (
                                                    <div>アンケートの取得に失敗しました: <span className="text-red-600">{surveyFetchError}</span> <button className="ml-2 text-blue-600 underline" onClick={() => { const t = localStorage.getItem('slack_app_session'); runGas('listSurveys', t).then(list => { if (Array.isArray(list)) { const mapped = list.map(s => ({ title: s.title || s.spreadsheetId || s.formUrl || s.spreadsheetUrl, ref: s.spreadsheetId || s.spreadsheetUrl || s.formUrl || '' })); setSurveys(mapped.filter(s=>s.ref)); safeSetSurveyError(''); } else if (list && list.message) { safeSetSurveyError(String(list.message)); } }).catch(e=>safeSetSurveyError(e && e.message?e.message:String(e))); }}>再読み込み</button></div>
                                                ) : (
                                                    <div>アンケートが見つかりません。<button className="ml-2 text-blue-600 underline" onClick={() => { const t = localStorage.getItem('slack_app_session'); runGas('listSurveys', t).then(list => { if (Array.isArray(list)) { const mapped = list.map(s => ({ title: s.title || s.spreadsheetId || s.formUrl || s.spreadsheetUrl, ref: s.spreadsheetId || s.spreadsheetUrl || s.formUrl || '' })); setSurveys(mapped.filter(s=>s.ref)); safeSetSurveyError(''); } else if (list && list.message) { safeSetSurveyError(String(list.message)); } }).catch(e=>safeSetSurveyError(e && e.message?e.message:String(e))); }}>再読み込み</button></div>
                                                )}
                                            </div>
                                        )}
                                    </div>
                                )}

                                {sourceType === 'collection' && (
                                    <div>
                                        <select className="w-full border p-2 rounded" value={selectedCollectionId} onChange={(e)=>setSelectedCollectionId(e.target.value)}>
                                            <option value="">集金を選択（出力時に名簿の後に追加）</option>
                                            {collections && collections.length > 0 ? collections.map(c => <option key={c.id} value={c.id}>{c.title || c.spreadsheetUrl || c.id}</option>) : null}
                                        </select>
                                        {collectionsLoading && <div className="mt-2 text-xs text-gray-500">集金一覧を読み込み中...</div>}
                                        {collections.length === 0 && !collectionsLoading && (
                                            <div className="mt-2 text-xs text-gray-500">
                                                {collectionsError ? (
                                                    <div>集金一覧の取得に失敗しました: <span className="text-red-600">{collectionsError}</span></div>
                                                ) : (
                                                    <div>集金が見つかりません。</div>
                                                )}
                                            </div>
                                        )}
                                    </div>
                                )}
                            </div>
                        </div>

                        <div className="mb-4">
                            <h4 className="font-medium mb-2">出力対象</h4>

                            <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mb-3">
                                <div>
                                    <label className="block text-sm mb-1">学年で絞り込む</label>
                                    <div className="flex flex-wrap gap-2">
                                        <button type="button" onClick={clearGradeFilter} className="text-xs px-2 py-1 border rounded bg-gray-50">ALL</button>
                                        {(gradeOptions || []).length === 0 && <div className="text-xs text-gray-500">学年データがありません。</div>}
                                        {(gradeOptions || []).map(g => (
                                            <label key={g} className="inline-flex items-center text-sm px-2 py-1 border rounded">
                                                <input type="checkbox" className="mr-2" checked={selectedGrades.indexOf(g) !== -1} onChange={() => toggleGradeFilter(g)} />{g}
                                            </label>
                                        ))}
                                    </div>
                                </div>
                                <div>
                                    <label className="block text-sm mb-1">分野で絞り込む</label>
                                    <div className="flex flex-wrap gap-2">
                                        <button type="button" onClick={clearFieldFilter} className="text-xs px-2 py-1 border rounded bg-gray-50">ALL</button>
                                        {(fieldOptions || []).length === 0 && <div className="text-xs text-gray-500">分野データがありません。</div>}
                                        {(fieldOptions || []).map(f => (
                                            <label key={f} className="inline-flex items-center text-sm px-2 py-1 border rounded">
                                                <input type="checkbox" className="mr-2" checked={selectedFieldFilters.indexOf(f) !== -1} onChange={() => toggleFieldFilter(f)} />{f}
                                            </label>
                                        ))}
                                    </div>
                                </div>
                            </div>

                            <div className="grid grid-cols-1 md:grid-cols-3 gap-3 mb-3">
                                <div>
                                    <select value={targetType} onChange={(e)=>setTargetType(e.target.value)} className="border p-2 rounded w-full md:w-64">
                                        <option value="all">出力対象：ALL</option>
                                        <option value="orgs">選択した局・部門のみ</option>
                                    </select>
                                </div>
                                <div>
                                    <select value={orgMatchMode} onChange={(e)=>setOrgMatchMode(e.target.value)} disabled={targetType!=='orgs'} className={`border p-2 rounded w-full md:w-64 ${targetType!=='orgs' ? 'opacity-60' : ''}`}>
                                        <option value="mainOnly">兼局: 含まない</option>
                                        <option value="allAffiliations">兼局: 含む</option>
                                    </select>
                                </div>
                                <div>
                                    <select value={statusFilter} onChange={(e)=>setStatusFilter(e.target.value)} className="border p-2 rounded w-full md:w-64">
                                        <option value="active">在籍者のみ</option>
                                        <option value="retired">退局者のみ</option>
                                        <option value="all">在籍状況: ALL</option>
                                    </select>
                                </div>
                            </div>
                            {targetType === 'orgs' && (
                                <div>
                                    <div className="mt-3 grid grid-cols-1 md:grid-cols-2 gap-4">
                                        {orgList.map(o => (
                                            <div key={o.name} className="border rounded p-3">
                                                <label className="flex items-center mb-2">
                                                    <input type="checkbox" className="mr-2" checked={!!orgChecked[o.name]} onChange={()=>toggleOrg(o.name)} />
                                                    <span className="font-medium">{o.name}</span>
                                                </label>
                                                <div className="pl-6 flex flex-wrap gap-2">
                                                    {(o.depts||[]).length===0 ? <div className="text-sm text-gray-500">部門情報なし</div> : o.depts.map(d => (
                                                        <label key={d} className="inline-flex items-center text-sm px-2 py-1 border rounded">
                                                            <input type="checkbox" className="mr-2" checked={deptChecked[o.name] ? !!deptChecked[o.name][d] : false} onChange={()=>toggleDept(o.name, d)} />{d}
                                                        </label>
                                                    ))}
                                                </div>
                                            </div>
                                        ))}
                                    </div>
                                </div>
                            )}
                        </div>

                        <div className="text-right">
                            <button onClick={openExportModal} disabled={busy} className={`bg-indigo-600 text-white px-6 py-3 rounded-lg font-bold shadow hover:bg-indigo-700 ${busy?'opacity-60':'opacity-100'}`}>
                                出力
                            </button>
                        </div>
                    </div>

                    {modalOpen && (
                        <div className="fixed inset-0 flex items-center justify-center modal-overlay p-4 z-[200]">
                            <div className="bg-white rounded-xl shadow-2xl w-full max-w-4xl p-6 flex flex-col animate-fade-in max-h-[80vh] overflow-hidden">
                                {busy || !previewData ? (
                                    <div className="flex flex-col items-center justify-center py-12">
                                        <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mb-4"></div>
                                        <p className="text-gray-600">読み込み中...</p>
                                    </div>
                                ) : (
                                    <>
                                        <div className="px-0 mb-3">
                                            <h4 className="font-bold">出力プレビュー</h4>
                                            <div className="text-sm text-gray-500">{previewData.filename}</div>
                                        </div>

                                        <div className="flex-1 overflow-hidden">
                                            <div className="overflow-x-auto overflow-y-auto w-full border rounded p-2 max-h-[50vh]">
                                                <table className="min-w-max border-collapse table-auto text-sm whitespace-nowrap">
                                                    <thead>
                                                        <tr className="bg-gray-100">
                                                            {previewData.headers.map((h, idx) => <th key={idx} className="border px-2 py-1 text-left">{h}</th>)}
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        {previewData.rows.slice(0,500).map((r, ri) => (
                                                            <tr key={ri} className={`${ri%2===0? 'bg-white':'bg-gray-50'}`}>
                                                                {previewData.headers.map((_, ci) => <td key={ci} className="border px-2 py-1 align-top">{r[ci] || ''}</td>)}
                                                            </tr>
                                                        ))}
                                                    </tbody>
                                                </table>
                                            </div>
                                        </div>

                                        <div className="mt-3 border-t pt-3 flex items-center justify-end gap-3 shrink-0 bg-white">
                                            <input className="border p-2 rounded w-64" value={filename} onChange={e=>setFilename(e.target.value)} />
                                            <button onClick={()=>setModalOpen(false)} className="px-4 py-2 rounded bg-gray-100">キャンセル</button>
                                            <button onClick={async ()=>{
                                                try {
                                                    if (!previewData || !previewData.rows || !previewData.headers) { alert('コピーする内容がありません'); return; }
                                                    // build HTML table
                                                    const headers = previewData.headers;
                                                    const rows = previewData.rows;
                                                    const esc = (s) => String(s === null || typeof s === 'undefined' ? '' : s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
                                                    let html = '<table>';
                                                    html += '<thead><tr>' + headers.map(h => '<th>' + esc(h) + '</th>').join('') + '</tr></thead>';
                                                    html += '<tbody>' + rows.map(r => '<tr>' + headers.map((_,i) => '<td>' + esc(r[i] || '') + '</td>').join('') + '</tr>').join('') + '</tbody>';
                                                    html += '</table>';
                                                    // plain text (tab separated)
                                                    const plain = [headers.join('\t')].concat(rows.map(r => headers.map((_,i) => (r[i] || '')).join('\t'))).join('\n');

                                                    // try Clipboard API with HTML and plain text
                                                    if (navigator.clipboard && navigator.clipboard.write) {
                                                        const blobHtml = new Blob([html], { type: 'text/html' });
                                                        const blobPlain = new Blob([plain], { type: 'text/plain' });
                                                        const item = new ClipboardItem({ 'text/html': blobHtml, 'text/plain': blobPlain });
                                                        await navigator.clipboard.write([item]);
                                                        alert('表形式をクリップボードにコピーしました（貼り付け先で表として貼れます）');
                                                        return;
                                                    }

                                                    // fallback: write plain text only
                                                    if (navigator.clipboard && navigator.clipboard.writeText) {
                                                        await navigator.clipboard.writeText(plain);
                                                        alert('表形式テキストをクリップボードにコピーしました（タブ区切り）');
                                                        return;
                                                    }

                                                    // final fallback: legacy execCommand using a temporary textarea
                                                    const ta = document.createElement('textarea'); ta.value = plain; document.body.appendChild(ta); ta.select(); try { document.execCommand('copy'); alert('表形式テキストをクリップボードにコピーしました（フォールバック）'); } catch (e) { throw e; } finally { ta.remove(); }
                                                } catch (e) { alert('コピーに失敗しました: ' + (e && e.message ? e.message : e)); }
                                            }} className="px-4 py-2 rounded bg-gray-800 text-white">クリップボードにコピー</button>
                                            <button onClick={doExportDownload} disabled={busy} className="px-4 py-2 rounded bg-indigo-600 text-white">ダウンロード</button>
                                        </div>
                                    </>
                                )}
                            </div>
                        </div>
                    )}
                </div>
            );
        }

