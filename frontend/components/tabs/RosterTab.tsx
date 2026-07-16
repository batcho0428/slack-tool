// @ts-nocheck
'use client';
/* eslint-disable */

import { useEffect, useState } from 'react';

        export default function RosterTab({ user, runGas, parseCsv, formatPreviewBirthday }) {
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
                { key: '所属局6', label: '所属局6', adminOnly: false },
                { key: '所属部門6', label: '所属部門6', adminOnly: false },
                { key: '役職6', label: '役職6', adminOnly: false },
                { key: '所属局7', label: '所属局7', adminOnly: false },
                { key: '所属部門7', label: '所属部門7', adminOnly: false },
                { key: '役職7', label: '役職7', adminOnly: false },
                { key: '所属局8', label: '所属局8', adminOnly: false },
                { key: '所属部門8', label: '所属部門8', adminOnly: false },
                { key: '役職8', label: '役職8', adminOnly: false },
                { key: '所属局9', label: '所属局9', adminOnly: false },
                { key: '所属部門9', label: '所属部門9', adminOnly: false },
                { key: '役職9', label: '役職9', adminOnly: false },
                { key: '所属局10', label: '所属局10', adminOnly: false },
                { key: '所属部門10', label: '所属部門10', adminOnly: false },
                { key: '役職10', label: '役職10', adminOnly: false },
                { key: '出身校', label: '出身校', adminOnly: true },
                { key: '退局', label: '退局ステータス', adminOnly: true },
                { key: '次年度継続', label: '次年度継続ステータス', adminOnly: true },
                { key: 'Admin', label: '管理者権限', adminOnly: true }
            ]);

            const [selected, setSelected] = useState(new Set());
            const [selectAll, setSelectAll] = useState(false);
            // atomic keys and display grouping for所属2-10
            const availableAtomicKeys = fields.map(f=>f.key);
            const displayFields = (()=>{
                const out = [];
                let skippingGroup = false;
                for (const f of fields) {
                    // 所属2-10グループの開始を検出
                    if (f.key === '所属局2' && !skippingGroup) {
                        out.push({ key: '所属2-10', label: '所属局・部門・役職（2～10）', isGroup: true, adminOnly: false });
                        skippingGroup = true;
                        continue;
                    }
                    // 所属2-10グループに属するアイテムをスキップ
                    if (/^(所属局([2-9]|10)|所属部門([2-9]|10)|役職([2-9]|10))$/.test(f.key)) {
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
                // handle grouped key for 所属2-10: toggle all atomic members
                if (k === '所属2-10') {
                    const atomic = [];
                    for (let idx = 2; idx <= 10; idx++) {
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
                    if (v === '所属2-10') {
                        for (let idx = 2; idx <= 10; idx++) {
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
                                                for (let idx = 2; idx <= 10; idx++) { atomic.push(`所属局${idx}`); atomic.push(`所属部門${idx}`); atomic.push(`役職${idx}`); }
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
