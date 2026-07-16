// @ts-nocheck
'use client';
/* eslint-disable */

import { useEffect, useMemo, useState } from 'react';
import ResultModal from '../shared/ResultModal';
import DetailsModal from '../shared/DetailsModal';

        export default function SurveyTab({ user, runGas }) {
            const [surveys, setSurveys] = useState([]);
            const [surveysLoading, setSurveysLoading] = useState(false);
            const [surveyFetchError, setSurveyFetchError] = useState('');
            const safeSetSurveyError = (v) => { try { if (typeof setSurveyFetchError === 'function') setSurveyFetchError(v); } catch(e){} };
            const [loading, setLoading] = useState(true);
            const [selected, setSelected] = useState(null);
            const [err, setErr] = useState('');
            const [modalOpen, setModalOpen] = useState(false);
            const [modalLoading, setModalLoading] = useState(false);
            const [modalDetails, setModalDetails] = useState(null);
            const [isAdmin, setIsAdmin] = useState(false);
            const [formsLoading, setFormsLoading] = useState(false);
            const [formsError, setFormsError] = useState('');
            const [formDefs, setFormDefs] = useState([]);
            const [formOptions, setFormOptions] = useState(null);
            const [formSelectOpen, setFormSelectOpen] = useState(false);
            const [formEditOpen, setFormEditOpen] = useState(false);
            const [formSaving, setFormSaving] = useState(false);
            const [formDraft, setFormDraft] = useState({
                rowIndex: null,
                spreadsheetRef: '',
                formUrl: '',
                title: '',
                deadlineDate: '',
                inChargeOrg: '',
                inChargeDept: '',
                collecting: false,
                scoreName: '',
                scoreUnit: ''
            });
            const [remindSurveyModalOpen, setRemindSurveyModalOpen] = useState(false);
            const [remindFilterModalOpen, setRemindFilterModalOpen] = useState(false);
            const [remindRecipientsModalOpen, setRemindRecipientsModalOpen] = useState(false);
            const [remindConfirmModalOpen, setRemindConfirmModalOpen] = useState(false);
            const [remindLoading, setRemindLoading] = useState(false);
            const [remindSending, setRemindSending] = useState(false);
            const [remindError, setRemindError] = useState('');
            const [remindForms, setRemindForms] = useState([]);
            const [remindSurveyChecks, setRemindSurveyChecks] = useState({});
            const [remindStatusData, setRemindStatusData] = useState({ users: [], unansweredByEmail: {}, surveys: [] });
            const [remindSearch, setRemindSearch] = useState('');
            const [remindRecipientSel, setRemindRecipientSel] = useState(new Set());
            const [remindResult, setRemindResult] = useState(null);

            const [targetGrades, setTargetGrades] = useState(new Set());
            const [targetFields, setTargetFields] = useState(new Set());
            const [targetType, setTargetType] = useState('all');
            const [orgMatchMode, setOrgMatchMode] = useState('mainOnly');
            const [statusFilter, setStatusFilter] = useState('active');
            const [targetOrgSelections, setTargetOrgSelections] = useState([]);

            const orgMasterList = Array.isArray(formOptions && formOptions.orgMaster) ? formOptions.orgMaster : [];
            const deptMasterList = Array.isArray(formOptions && formOptions.deptMaster) ? formOptions.deptMaster : [];

            const resolveOrgLabel = (orgCode) => {
                const code = String(orgCode || '').trim();
                if (!code) return '';
                const found = orgMasterList.find(item => String(item.pid || '').trim() === code);
                return (found && found.org) ? found.org : code;
            };

            const resolveDeptLabel = (deptCode) => {
                const code = String(deptCode || '').trim();
                if (!code) return '';
                const found = deptMasterList.find(item => String(item.pid || '').trim() === code);
                return (found && found.dept) ? found.dept : code;
            };

            const formatAffiliationLabel = (orgCode, deptCode) => {
                const orgLabel = resolveOrgLabel(orgCode);
                const deptLabel = resolveDeptLabel(deptCode);
                if (deptLabel && deptLabel !== orgLabel) return (orgLabel || '-') + ' / ' + deptLabel;
                return orgLabel || deptLabel || '-';
            };

            const formIsValid = (formDraft && typeof formDraft === 'object') ? ((formDraft.title || '').toString().trim().length > 0 && (((formDraft.spreadsheetRef||'').toString().trim().length > 0) || ((formDraft.formUrl||'').toString().trim().length > 0))) : false;

            const unansweredCountByEmail = useMemo(() => {
                const out = {};
                const map = remindStatusData && remindStatusData.unansweredByEmail ? remindStatusData.unansweredByEmail : {};
                Object.keys(map).forEach(email => {
                    out[email] = Array.isArray(map[email]) ? map[email].length : 0;
                });
                return out;
            }, [remindStatusData]);

            const normalize = (v) => String(v || '').trim().toLowerCase();

            const matchesTargetFilter = (u) => {
                if (!u) return false;
                const userGrade = String(u.grade || '').trim();
                const userField = String(u.field || '').trim();
                const isRetired = !!u.retired;

                if (statusFilter === 'active' && isRetired) return false;
                if (statusFilter === 'retired' && !isRetired) return false;

                if (targetGrades.size > 0 && userGrade && !targetGrades.has(userGrade)) return false;
                if (targetFields.size > 0 && userField && !targetFields.has(userField)) return false;

                if (targetType === 'all') return true;
                if (!Array.isArray(targetOrgSelections) || targetOrgSelections.length === 0) return true;

                const affiliations = Array.isArray(u.affiliations) ? u.affiliations : [];
                const matchAff = (aff, sel) => {
                    if (!aff || !sel) return false;
                    const org = String(aff.org || '').trim();
                    const dept = String(aff.dept || '').trim();
                    if (!org || org !== String(sel.org || '').trim()) return false;
                    if (!sel.dept) return true;
                    return dept === String(sel.dept || '').trim();
                };

                if (orgMatchMode === 'mainOnly') {
                    const aff0 = affiliations.length > 0 ? affiliations[0] : null;
                    return targetOrgSelections.some(sel => matchAff(aff0, sel));
                }

                return affiliations.some(aff => targetOrgSelections.some(sel => matchAff(aff, sel)));
            };

            const getDefaultRecipientSet = () => {
                const users = Array.isArray(remindStatusData && remindStatusData.users) ? remindStatusData.users : [];
                const s = new Set();
                users.forEach(u => {
                    const email = String(u.email || '').trim().toLowerCase();
                    if (!email) return;
                    const unansweredCount = Number(unansweredCountByEmail[email] || 0);
                    if (matchesTargetFilter(u) && unansweredCount > 0) s.add(email);
                });
                return s;
            };

            useEffect(() => {
                setLoading(true); setErr('');
                const token = localStorage.getItem('slack_app_session');
                runGas('listSurveys', token).then(res => {
                    if (Array.isArray(res)) {
                        setSurveys(res);
                        return;
                    }
                    if (res && res.success === false) {
                        setSurveys([]);
                        setErr(res.message || 'アンケート一覧の取得に失敗しました');
                        return;
                    }
                    setSurveys([]);
                }).catch(e => setErr(e.message || e)).finally(()=>setLoading(false));

                runGas('getUserProfile', token)
                    .then(p => { setIsAdmin(!!p.isAdmin); })
                    .catch(()=>{});
                // fetch org/department options for form editing dropdowns
                runGas('getSearchOptions').then(o=>setFormOptions(o)).catch(()=>{});
            }, []);

            const loadFormDefs = async () => {
                setFormsError('');
                setFormsLoading(true);
                try {
                    const token = localStorage.getItem('slack_app_session');
                    const res = await runGas('listFormDefinitions', token);
                    if (res && res.success) {
                        setFormDefs(res.items || []);
                    } else {
                        setFormDefs([]);
                        setFormsError(res && res.message ? res.message : '取得に失敗しました');
                    }
                } catch (e) {
                    setFormDefs([]);
                    setFormsError(e && e.message ? e.message : String(e));
                } finally {
                    setFormsLoading(false);
                }
            };

            const openNewForm = () => {
                setFormDraft({
                    rowIndex: null,
                    spreadsheetRef: '',
                    formUrl: '',
                    title: '',
                    deadlineDate: '',
                    inChargeOrg: '',
                    inChargeDept: '',
                    collecting: true,
                    scoreName: '',
                    scoreUnit: ''
                });
                setFormEditOpen(true);
            };

            const openEditSelect = async () => {
                setFormSelectOpen(true);
                if (!formsLoading) loadFormDefs();
            };

            const openEditForm = (item) => {
                if (!item) return;
                setFormDraft({
                    rowIndex: item.rowIndex,
                    spreadsheetRef: item.spreadsheetRef || '',
                    formUrl: item.formUrl || '',
                    title: item.title || '',
                    deadlineDate: item.deadlineDate || '',
                    inChargeOrg: item.inChargeOrg || '',
                    inChargeDept: item.inChargeDept || '',
                    collecting: !!item.collecting,
                    scoreName: item.scoreName || '',
                    scoreUnit: item.scoreUnit || ''
                });
                setFormSelectOpen(false);
                setFormEditOpen(true);
            };

            const saveForm = async () => {
                setFormSaving(true);
                try {
                    const token = localStorage.getItem('slack_app_session');
                    const res = await runGas('saveFormDefinition', token, formDraft);
                    if (!res || !res.success) throw new Error(res && res.message ? res.message : '保存に失敗しました');
                    setFormEditOpen(false);
                    loadFormDefs();
                    const list = await runGas('listSurveys', token);
                    if (Array.isArray(list)) {
                        setSurveys(list);
                    } else if (list && list.success === false) {
                        setSurveys([]);
                        setErr(list.message || 'アンケート一覧の取得に失敗しました');
                    } else {
                        setSurveys([]);
                    }
                } catch (e) {
                    alert(e.message || e);
                } finally {
                    setFormSaving(false);
                }
            };

            const openRemindSurveyModal = async () => {
                setRemindError('');
                setRemindLoading(true);
                setRemindSurveyModalOpen(true);
                try {
                    const token = localStorage.getItem('slack_app_session');
                    const [defsRes, opts] = await Promise.all([
                        runGas('listFormDefinitions', token),
                        runGas('getSearchOptions', token)
                    ]);
                    const defs = defsRes && defsRes.success ? (defsRes.items || []) : [];
                    setRemindForms(defs);

                    const checks = {};
                    defs.forEach(d => {
                        const rowIndex = Number(d && d.rowIndex);
                        if (!isNaN(rowIndex) && rowIndex >= 2) checks[rowIndex] = !!d.collecting;
                    });
                    setRemindSurveyChecks(checks);

                    const gradeList = Array.isArray(opts && opts.grades) ? opts.grades : [];
                    const fieldList = Array.isArray(opts && opts.fields) ? opts.fields : [];
                    setTargetGrades(new Set(gradeList));
                    setTargetFields(new Set(fieldList));
                    setTargetType('all');
                    setOrgMatchMode('mainOnly');
                    setStatusFilter('active');
                    setTargetOrgSelections([]);
                } catch (e) {
                    setRemindError(e && e.message ? e.message : String(e));
                } finally {
                    setRemindLoading(false);
                }
            };

            const toggleRemindSurvey = (rowIndex) => {
                setRemindSurveyChecks(prev => ({ ...prev, [rowIndex]: !prev[rowIndex] }));
            };

            const proceedToTargetFilter = async () => {
                const selectedRows = Object.keys(remindSurveyChecks).map(v => Number(v)).filter(v => remindSurveyChecks[v]);
                if (selectedRows.length === 0) {
                    setRemindError('リマインド対象アンケートを1件以上選択してください。');
                    return;
                }
                setRemindLoading(true);
                setRemindError('');
                try {
                    const token = localStorage.getItem('slack_app_session');
                    const statusRes = await runGas('collectSurveyReminderStatus', token, selectedRows);
                    if (!statusRes || statusRes.success === false) {
                        throw new Error(statusRes && statusRes.message ? statusRes.message : '回答状況の取得に失敗しました');
                    }
                    setRemindStatusData({
                        users: statusRes.users || [],
                        unansweredByEmail: statusRes.unansweredByEmail || {},
                        surveys: statusRes.surveys || []
                    });
                    setRemindSurveyModalOpen(false);
                    setRemindFilterModalOpen(true);
                } catch (e) {
                    setRemindError(e && e.message ? e.message : String(e));
                } finally {
                    setRemindLoading(false);
                }
            };

            const toggleSetValue = (setter, prevSet, value) => {
                const next = new Set(prevSet);
                if (next.has(value)) next.delete(value);
                else next.add(value);
                setter(next);
            };

            const toggleOrgSelection = (org, dept) => {
                const orgVal = String(org || '').trim();
                const deptVal = String(dept || '').trim();
                setTargetOrgSelections(prev => {
                    const has = prev.some(x => String(x.org || '') === orgVal && String(x.dept || '') === deptVal);
                    if (has) return prev.filter(x => !(String(x.org || '') === orgVal && String(x.dept || '') === deptVal));
                    if (!deptVal) {
                        return [...prev.filter(x => String(x.org || '') !== orgVal), { org: orgVal, dept: '' }];
                    }
                    const withoutOrgAll = prev.filter(x => !(String(x.org || '') === orgVal && String(x.dept || '') === ''));
                    return [...withoutOrgAll, { org: orgVal, dept: deptVal }];
                });
            };

            const proceedToRecipients = () => {
                setRemindFilterModalOpen(false);
                setRemindRecipientsModalOpen(true);
                setRemindSearch('');
                setRemindRecipientSel(getDefaultRecipientSet());
            };

            const filteredReminderUsers = useMemo(() => {
                const users = Array.isArray(remindStatusData && remindStatusData.users) ? remindStatusData.users : [];
                const q = normalize(remindSearch);
                const sorted = users.slice().sort((a, b) => {
                    const ac = Number(unansweredCountByEmail[String(a.email || '').toLowerCase()] || 0);
                    const bc = Number(unansweredCountByEmail[String(b.email || '').toLowerCase()] || 0);
                    if (bc !== ac) return bc - ac;
                    return String(a.name || '').localeCompare(String(b.name || ''), 'ja');
                });
                if (!q) return sorted;
                return sorted.filter(u => {
                    const text = [u.name, u.email, u.grade, u.field, u.departmentText].map(v => String(v || '')).join(' ').toLowerCase();
                    return text.includes(q);
                });
            }, [remindStatusData, remindSearch, unansweredCountByEmail]);

            const toggleRecipient = (email) => {
                const key = String(email || '').trim().toLowerCase();
                setRemindRecipientSel(prev => {
                    const next = new Set(prev);
                    if (next.has(key)) next.delete(key);
                    else next.add(key);
                    return next;
                });
            };

            const toggleAllRecipients = () => {
                setRemindRecipientSel(prev => {
                    const next = new Set(prev);
                    const allSelected = filteredReminderUsers.length > 0 && filteredReminderUsers.every(u => next.has(String(u.email || '').toLowerCase()));
                    filteredReminderUsers.forEach(u => {
                        const key = String(u.email || '').toLowerCase();
                        if (allSelected) next.delete(key);
                        else next.add(key);
                    });
                    return next;
                });
            };

            const selectedReminderRecipients = useMemo(() => {
                const users = Array.isArray(remindStatusData && remindStatusData.users) ? remindStatusData.users : [];
                return users.filter(u => remindRecipientSel.has(String(u.email || '').toLowerCase()));
            }, [remindStatusData, remindRecipientSel]);

            const openConfirmModal = () => {
                if (selectedReminderRecipients.length === 0) {
                    setRemindError('リマインド対象者を1名以上選択してください。');
                    return;
                }
                setRemindRecipientsModalOpen(false);
                setRemindConfirmModalOpen(true);
            };

            const executeReminderSend = async () => {
                setRemindConfirmModalOpen(false);
                setRemindSending(true);
                setRemindError('');
                try {
                    const token = localStorage.getItem('slack_app_session');
                    const recipientsPayload = selectedReminderRecipients.map(u => ({
                        name: u.name,
                        email: String(u.email || '').toLowerCase()
                    }));
                    const surveyRowIndices = (remindStatusData.surveys || []).map(s => s.rowIndex);
                    const res = await runGas('sendSurveyReminderDMs', token, { recipients: recipientsPayload, surveyRowIndices });
                    setRemindResult({ success: Number(res && res.success ? res.success : 0), failed: Array.isArray(res && res.failed) ? res.failed : [] });
                    setRemindRecipientSel(new Set());
                } catch (e) {
                    setRemindError(e && e.message ? e.message : String(e));
                } finally {
                    setRemindSending(false);
                }
            };

            const selectedUsersWithoutUnanswered = selectedReminderRecipients.filter(u => {
                const email = String(u.email || '').toLowerCase();
                return Number(unansweredCountByEmail[email] || 0) === 0;
            });

            const openDetails = async (survey) => {
                if (!survey) return;
                const ref = survey.spreadsheetUrl || survey.spreadsheetId;
                if (!ref) return;
                setSelected(ref); setErr(''); setModalOpen(true); setModalLoading(true); setModalDetails(null);
                try {
                    const token = localStorage.getItem('slack_app_session');
                    const rowIndex = survey.userLatestRowIndex;
                    const d = await runGas('getSurveyDetails', token, ref, rowIndex);
                    if (!d) { setModalDetails({ error: 'サーバー応答が不正です' }); setModalLoading(false); return; }
                    if (d.success === false) { setModalDetails({ error: d.message || '参照できません' }); setModalLoading(false); return; }
                    setModalDetails(d);
                } catch (e) { setModalDetails({ error: e.message || String(e) }); }
                finally { setModalLoading(false); }
            };

            const openForm = (survey) => {
                if (!survey) return;
                const url = String(survey.formUrl || '').trim();
                if (!url) { setErr('フォームURLが設定されていません。'); return; }
                // open raw URL as-is
                try { window.open(url, '_blank', 'noopener'); } catch (e) { setErr('リンクを開けませんでした'); }
            };

            const isCollectingFlag = (s) => {
                if (!s) return false;
                if (s.collecting === true) return true;
                const sval = String(s.collecting || '').trim().toLowerCase();
                if (sval === 'true' || sval === '1') return true;
                if (typeof s.collecting === 'number' && Number(s.collecting) === 1) return true;
                return false;
            };

            const answered = surveys.filter(s => s && s.available);
            const unanswered = surveys.filter(s => s && !s.available && isCollectingFlag(s));

            const selectedSurvey = surveys.find(s => (s && (s.spreadsheetUrl || s.spreadsheetId)) === selected) || null;

            const formatDateJP = (ts) => {
                if (!ts) return '-';
                try {
                    const d = new Date(Number(ts));
                    if (isNaN(d.getTime())) return '-';
                    const y = d.getFullYear();
                    const mo = ('0' + (d.getMonth() + 1)).slice(-2);
                    const da = ('0' + d.getDate()).slice(-2);
                    const hh = ('0' + d.getHours()).slice(-2);
                    const mm = ('0' + d.getMinutes()).slice(-2);
                    return `${y}/${mo}/${da} ${hh}:${mm}`;
                } catch (e) { return '-'; }
            };

            return (
                <div className="h-full overflow-auto p-4 bg-gray-50">
                    <div className="max-w-6xl mx-auto bg-white rounded-lg shadow-sm p-6 space-y-6">
                        <div className="flex items-start justify-between gap-3">
                            <div>
                                <h3 className="text-2xl font-bold mb-1">アンケート</h3>
                            </div>
                            <div className="flex items-center space-x-2">
                                <button onClick={openNewForm} className="text-xs bg-green-100 text-green-900 px-3 py-1 rounded hover:brightness-95">
                                    新規登録
                                </button>
                                <button onClick={openEditSelect} className="text-xs bg-yellow-100 text-yellow-900 px-3 py-1 rounded hover:brightness-95">
                                    編集
                                </button>
                                <button onClick={openRemindSurveyModal} className="text-xs bg-blue-100 text-blue-900 px-3 py-1 rounded hover:brightness-95">
                                    リマインド送信
                                </button>
                            </div>
                        </div>

                        {loading && <div className="text-gray-500"><i className="fas fa-spinner fa-spin mr-2"></i>読み込み中...</div>}
                        {err && <div className="text-red-600 mb-3">{err}</div>}

                        <div className="flex flex-col gap-4">
                            <div className="border rounded p-4 bg-white">
                                <div className="flex items-center justify-between mb-3">
                                    <div className="font-medium">未回答</div>
                                    <div className="text-xs text-gray-500">収集中のフォーム</div>
                                </div>
                                {unanswered.length === 0 && !loading ? (
                                    <div className="text-sm text-gray-600">未回答のアンケートはありません。</div>
                                ) : (
                                    <div className="space-y-3">
                                        {unanswered.map((s, i) => (
                                                    <div key={i} onClick={() => openForm(s)} role="button" tabIndex={0} className="flex items-start justify-between p-3 border rounded hover:shadow-sm cursor-pointer">
                                                        <div>
                                                            <div className="font-medium">{s.title}</div>
                                                            <div className="text-xs text-gray-600">{formatAffiliationLabel(s.inChargeOrg || s.inChargeOrgCode, s.inChargeDept || s.inChargeDeptCode)}</div>
                                                        </div>
                                                        <div className="flex items-center gap-2">
                                                            {isCollectingFlag(s) && <span className="text-xs bg-green-100 text-green-800 px-2 py-1 rounded">収集中</span>}
                                                        </div>
                                                    </div>
                                                ))}
                                    </div>
                                )}
                            </div>

                            <div className="border rounded p-4 bg-white">
                                <div className="flex items-center justify-between mb-3">
                                    <div className="font-medium">回答済み</div>
                                    <div className="text-xs text-gray-500">最新の回答とスコア</div>
                                </div>
                                <div className="overflow-auto">
                                    <table className="min-w-full text-sm">
                                        <thead className="bg-gray-100">
                                            <tr>
                                                <th className="px-3 py-2 text-left">アンケートタイトル</th>
                                                <th className="px-3 py-2 text-left">担当局 / 部門</th>
                                                <th className="px-3 py-2 text-left">スコア</th>
                                                <th className="px-3 py-2 text-left">回答日</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {answered.map((s, i) => (
                                                <tr key={i} onClick={() => openDetails(s)} role="button" tabIndex={0} className={`${'hover:bg-gray-50 cursor-pointer'} ${selected===(s.spreadsheetUrl || s.spreadsheetId)? 'bg-blue-50':''}`}>
                                                    <td className="px-3 py-2 border-t">{s.title}</td>
                                                    <td className="px-3 py-2 border-t text-sm text-gray-700">{formatAffiliationLabel(s.inChargeOrg || s.inChargeOrgCode, s.inChargeDept || s.inChargeDeptCode)}</td>
                                                    <td className="px-3 py-2 border-t">{s.latestScoreFormatted ? s.latestScoreFormatted : (s.latestScore !== null && typeof s.latestScore !== 'undefined' ? String(s.latestScore) + (s.scoreUnit ? (' ' + s.scoreUnit) : '') : '-')}</td>
                                                    <td className="px-3 py-2 border-t">{s.latestResponseDate ? formatDateJP(s.latestResponseDate) : '-'}</td>
                                                </tr>
                                            ))}
                                            {answered.length===0 && !loading && <tr><td className="px-3 py-4" colSpan={4}>回答済みのアンケートはありません。</td></tr>}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>

                        <DetailsModal isOpen={modalOpen} loading={modalLoading} details={modalDetails} survey={selectedSurvey} onClose={() => { setModalOpen(false); setModalDetails(null); setModalLoading(false); }} />
                        {formSelectOpen && (
                            <div className="fixed inset-0 flex items-center justify-center modal-overlay p-4 z-[220]">
                                <div className="bg-white rounded-xl shadow-2xl w-full max-w-2xl p-5 max-h-[80vh] overflow-auto">
                                    <div className="flex items-center justify-between mb-3">
                                        <h4 className="font-bold">アンケートを選択</h4>
                                        <button onClick={() => setFormSelectOpen(false)} className="text-sm text-gray-500 hover:text-gray-700">閉じる</button>
                                    </div>
                                    {formsLoading && <div className="text-gray-500 text-sm"><i className="fas fa-spinner fa-spin mr-2"></i>読み込み中...</div>}
                                    {formsError && <div className="text-red-600 text-sm mb-2">{formsError}</div>}
                                    {!formsLoading && formDefs.length === 0 && !formsError && (
                                        <div className="text-sm text-gray-500">登録済みのアンケートがありません。</div>
                                    )}
                                    <div className="divide-y">
                                        {[...formDefs].reverse().map((f, i) => (
                                            <div key={i} className="py-3 flex items-center justify-between">
                                                <div className="min-w-0">
                                                    <div className="font-medium truncate">{f.title || f.spreadsheetRef || f.formUrl || '無題'}</div>
                                                    <div className="text-xs text-gray-500 truncate">{formatAffiliationLabel(f.inChargeOrg || f.inChargeOrgCode, f.inChargeDept || f.inChargeDeptCode)}</div>
                                                </div>
                                                <button onClick={() => openEditForm(f)} className="text-sm px-3 py-1 rounded border">編集</button>
                                            </div>
                                        ))}
                                    </div>
                                </div>
                            </div>
                        )}
                        {formEditOpen && (
                            <div className="fixed inset-0 flex items-center justify-center modal-overlay p-4 z-[230]">
                                <div className="bg-white rounded-xl shadow-2xl w-full max-w-2xl p-5 max-h-[85vh] overflow-auto relative">
                                        {formSaving && (
                                            <div className="absolute inset-0 flex items-center justify-center bg-white/60 z-50">
                                                <div className="flex flex-col items-center">
                                                    <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mb-3"></div>
                                                    <div className="text-sm text-gray-600">読み込み中…</div>
                                                </div>
                                            </div>
                                        )}
                                    <div className="flex items-center justify-between mb-3">
                                        <h4 className="font-bold">アンケート登録 / 編集</h4>
                                        <button onClick={() => setFormEditOpen(false)} className="text-sm text-gray-500 hover:text-gray-700">閉じる</button>
                                    </div>
                                    <div className="space-y-3 text-sm">
                                        <p>スプレッドシートにメールアドレス・スコアを登録すれば、フォームに関係なく個人別データの配布が可能です。</p>
                                        <p>回答を求める場合は、スプレッドシートURLとGoogle Forms回答者リンクの両方を設定してください。</p>
                                        <br />
                                        <div>
                                            <label className="block text-xs text-gray-500 mb-1">スプレッドシートURL(共有→リンクをコピー)</label>
                                            <input className="w-full border p-2 rounded" value={formDraft.spreadsheetRef} onChange={e => setFormDraft(prev => ({ ...prev, spreadsheetRef: e.target.value }))} />
                                        </div>
                                        <div>
                                            <label className="block text-xs text-gray-500 mb-1">Google Forms回答者リンク</label>
                                            <input className="w-full border p-2 rounded" value={formDraft.formUrl} onChange={e => setFormDraft(prev => ({ ...prev, formUrl: e.target.value }))} />
                                        </div>
                                        <div>
                                            <label className="block text-xs text-gray-500 mb-1">タイトル</label>
                                            <input className="w-full border p-2 rounded" value={formDraft.title} onChange={e => setFormDraft(prev => ({ ...prev, title: e.target.value }))} />
                                        </div>
                                        <div>
                                            <label className="block text-xs text-gray-500 mb-1">締め切り日（任意）</label>
                                            <input type="date" className="w-full border p-2 rounded" value={formDraft.deadlineDate || ''} onChange={e => setFormDraft(prev => ({ ...prev, deadlineDate: e.target.value }))} />
                                        </div>
                                        <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                                            <div>
                                                <label className="block text-xs text-gray-500 mb-1">担当局</label>
                                                <select className="w-full border p-2 rounded" value={formDraft.inChargeOrg} onChange={e => setFormDraft(prev => ({ ...prev, inChargeOrg: e.target.value, inChargeDept: '' }))}>
                                                    <option value="">選択</option>
                                                    {orgMasterList.map((o2,idx)=>(<option key={o2.pid || idx} value={o2.pid}>{o2.org}</option>))}
                                                </select>
                                            </div>
                                            <div>
                                                <label className="block text-xs text-gray-500 mb-1">担当部門</label>
                                                <select className="w-full border p-2 rounded" value={formDraft.inChargeDept} onChange={e => setFormDraft(prev => ({ ...prev, inChargeDept: e.target.value }))} disabled={!(deptMasterList.length && formDraft.inChargeOrg)}>
                                                    <option value="">選択</option>
                                                    {deptMasterList.filter(d => String(d.orgPid || '').trim() === String(formDraft.inChargeOrg || '').trim()).map((d2,idx)=>(<option key={d2.pid || idx} value={d2.pid}>{d2.dept}</option>))}
                                                </select>
                                            </div>
                                        </div>
                                        <div className="flex items-center gap-3">
                                            <button type="button" onClick={() => setFormDraft(prev => ({ ...prev, collecting: !prev.collecting }))} className={`relative inline-flex flex-shrink-0 h-6 w-11 border-2 border-transparent rounded-full cursor-pointer transition-colors ease-in-out duration-200 ${formDraft.collecting ? 'bg-green-500' : 'bg-gray-300'}`} role="switch" aria-checked={!!formDraft.collecting}>
                                                <span className={`inline-block h-5 w-5 transform bg-white rounded-full shadow ring-0 transition ease-in-out duration-200 ${formDraft.collecting ? 'translate-x-5' : 'translate-x-0'}`} />
                                            </button>
                                            <span className="text-sm">{formDraft.collecting ? '収集中' : '収集終了'}</span>
                                        </div>
                                        <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                                            <div>
                                                <label className="block text-xs text-gray-500 mb-1">スコア名</label>
                                                <input className="w-full border p-2 rounded" value={formDraft.scoreName} onChange={e => setFormDraft(prev => ({ ...prev, scoreName: e.target.value }))} />
                                            </div>
                                            <div>
                                                <label className="block text-xs text-gray-500 mb-1">スコア単位</label>
                                                <input className="w-full border p-2 rounded" value={formDraft.scoreUnit} onChange={e => setFormDraft(prev => ({ ...prev, scoreUnit: e.target.value }))} />
                                            </div>
                                        </div>
                                    </div>
                                        <div className="mt-4 flex items-center justify-end gap-2">
                                        <button onClick={() => setFormEditOpen(false)} className="px-4 py-2 rounded bg-gray-100">キャンセル</button>
                                        <button onClick={saveForm} disabled={formSaving || !formIsValid} className={`px-4 py-2 rounded text-white ${formSaving || !formIsValid ? 'bg-gray-400' : 'bg-blue-600 hover:bg-blue-700'}`}>
                                            {formSaving ? '保存中...' : '保存'}
                                        </button>
                                    </div>
                                </div>
                            </div>
                        )}

                        {remindSurveyModalOpen && (
                            <div className="fixed inset-0 flex items-center justify-center modal-overlay p-4 z-[240]">
                                <div className="bg-white rounded-xl shadow-2xl w-full max-w-3xl p-5 max-h-[85vh] overflow-auto relative">
                                    {remindLoading && (
                                        <div className="absolute inset-0 flex items-center justify-center bg-white/70 z-50">
                                            <div className="flex flex-col items-center">
                                                <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mb-3"></div>
                                                <div className="text-sm text-gray-600">回答状況を取得しています...</div>
                                            </div>
                                        </div>
                                    )}
                                    <div className="flex items-center justify-between mb-3">
                                        <h4 className="font-bold">リマインド対象アンケート選択</h4>
                                        <button onClick={() => setRemindSurveyModalOpen(false)} className="text-sm text-gray-500 hover:text-gray-700">閉じる</button>
                                    </div>
                                    {remindError && <div className="text-sm text-red-600 mb-2">{remindError}</div>}
                                    <div className="divide-y border rounded">
                                        {remindForms.length === 0 && <div className="p-3 text-sm text-gray-500">対象アンケートがありません。</div>}
                                        {remindForms.map((f, idx) => {
                                            const rowIndex = Number(f && f.rowIndex);
                                            const checked = !!remindSurveyChecks[rowIndex];
                                            return (
                                                <label key={rowIndex || idx} className="p-3 flex items-center gap-3 cursor-pointer hover:bg-gray-50">
                                                    <input type="checkbox" checked={checked} onChange={() => toggleRemindSurvey(rowIndex)} className="w-4 h-4" />
                                                    <div className="min-w-0">
                                                        <div className="font-medium truncate">{f.title || f.spreadsheetRef || f.formUrl || '無題'}</div>
                                                        <div className="text-xs text-gray-500 truncate">{formatAffiliationLabel(f.inChargeOrg || f.inChargeOrgCode, f.inChargeDept || f.inChargeDeptCode)}</div>
                                                    </div>
                                                    {f.collecting && <span className="text-xs bg-green-100 text-green-800 px-2 py-1 rounded">収集中</span>}
                                                </label>
                                            );
                                        })}
                                    </div>
                                    <div className="mt-4 flex justify-end gap-2">
                                        <button onClick={() => setRemindSurveyModalOpen(false)} className="px-4 py-2 rounded bg-gray-100">キャンセル</button>
                                        <button onClick={proceedToTargetFilter} className="px-4 py-2 rounded text-white bg-blue-600 hover:bg-blue-700">次へ</button>
                                    </div>
                                </div>
                            </div>
                        )}

                        {remindFilterModalOpen && (
                            <div className="fixed inset-0 flex items-center justify-center modal-overlay p-4 z-[245]">
                                <div className="bg-white rounded-xl shadow-2xl w-full max-w-4xl p-5 max-h-[88vh] overflow-auto">
                                    <div className="flex items-center justify-between mb-3">
                                        <h4 className="font-bold">対象ユーザー区分</h4>
                                        <button onClick={() => setRemindFilterModalOpen(false)} className="text-sm text-gray-500 hover:text-gray-700">閉じる</button>
                                    </div>

                                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                        <div className="border rounded p-3">
                                            <div className="flex items-center justify-between mb-2">
                                                <div className="font-medium text-sm">学年</div>
                                                <button type="button" className="text-xs px-2 py-1 border rounded" onClick={() => setTargetGrades(new Set(Array.isArray(formOptions && formOptions.grades) ? formOptions.grades : []))}>ALL</button>
                                            </div>
                                            <div className="max-h-40 overflow-auto space-y-1">
                                                {(Array.isArray(formOptions && formOptions.grades) ? formOptions.grades : []).map((g, idx) => (
                                                    <label key={g || idx} className="flex items-center gap-2 text-sm">
                                                        <input type="checkbox" checked={targetGrades.has(g)} onChange={() => toggleSetValue(setTargetGrades, targetGrades, g)} className="w-4 h-4" />
                                                        <span>{g}</span>
                                                    </label>
                                                ))}
                                            </div>
                                        </div>

                                        <div className="border rounded p-3">
                                            <div className="flex items-center justify-between mb-2">
                                                <div className="font-medium text-sm">分野</div>
                                                <button type="button" className="text-xs px-2 py-1 border rounded" onClick={() => setTargetFields(new Set(Array.isArray(formOptions && formOptions.fields) ? formOptions.fields : []))}>ALL</button>
                                            </div>
                                            <div className="max-h-40 overflow-auto space-y-1">
                                                {(Array.isArray(formOptions && formOptions.fields) ? formOptions.fields : []).map((f, idx) => (
                                                    <label key={f || idx} className="flex items-center gap-2 text-sm">
                                                        <input type="checkbox" checked={targetFields.has(f)} onChange={() => toggleSetValue(setTargetFields, targetFields, f)} className="w-4 h-4" />
                                                        <span>{f}</span>
                                                    </label>
                                                ))}
                                            </div>
                                        </div>
                                    </div>

                                    <div className="grid grid-cols-1 md:grid-cols-3 gap-3 mt-4">
                                        <div>
                                            <label className="block text-xs text-gray-500 mb-1">出力対象</label>
                                            <select value={targetType} onChange={e => setTargetType(e.target.value)} className="w-full border p-2 rounded text-sm">
                                                <option value="all">出力対象: ALL</option>
                                                <option value="orgs">局 / 部門を指定</option>
                                            </select>
                                        </div>
                                        <div>
                                            <label className="block text-xs text-gray-500 mb-1">兼局</label>
                                            <select value={orgMatchMode} onChange={e => setOrgMatchMode(e.target.value)} className="w-full border p-2 rounded text-sm">
                                                <option value="mainOnly">兼局: 含まない</option>
                                                <option value="allAffiliations">兼局: 含む</option>
                                            </select>
                                        </div>
                                        <div>
                                            <label className="block text-xs text-gray-500 mb-1">在籍状況</label>
                                            <select value={statusFilter} onChange={e => setStatusFilter(e.target.value)} className="w-full border p-2 rounded text-sm">
                                                <option value="active">在籍者のみ</option>
                                                <option value="retired">退局者のみ</option>
                                                <option value="all">在籍状況: ALL</option>
                                            </select>
                                        </div>
                                    </div>

                                    {targetType === 'orgs' && (
                                        <div className="mt-4 border rounded p-3">
                                            <div className="font-medium text-sm mb-2">局 / 部門</div>
                                            <div className="max-h-52 overflow-auto space-y-2">
                                                {orgMasterList.map((org, idx) => {
                                                    const orgName = String(org && org.org ? org.org : '');
                                                    const orgChecked = targetOrgSelections.some(s => String(s.org || '') === orgName && !s.dept);
                                                    const depts = deptMasterList.filter(d => String(d.org || '') === orgName).map(d => d.dept);
                                                    return (
                                                        <div key={org.pid || idx} className="border rounded p-2">
                                                            <label className="flex items-center gap-2 text-sm font-medium">
                                                                <input type="checkbox" checked={orgChecked} onChange={() => toggleOrgSelection(orgName, '')} className="w-4 h-4" />
                                                                <span>{orgName}</span>
                                                            </label>
                                                            <div className="mt-1 ml-5 grid grid-cols-1 md:grid-cols-2 gap-1">
                                                                {depts.map((dept, dIdx) => {
                                                                    const checked = targetOrgSelections.some(s => String(s.org || '') === orgName && String(s.dept || '') === String(dept || ''));
                                                                    return (
                                                                        <label key={String(dept || '') + dIdx} className="flex items-center gap-2 text-xs">
                                                                            <input type="checkbox" checked={checked} onChange={() => toggleOrgSelection(orgName, dept)} className="w-3.5 h-3.5" />
                                                                            <span>{dept}</span>
                                                                        </label>
                                                                    );
                                                                })}
                                                            </div>
                                                        </div>
                                                    );
                                                })}
                                            </div>
                                        </div>
                                    )}

                                    <div className="mt-4 flex justify-end gap-2">
                                        <button onClick={() => { setRemindFilterModalOpen(false); setRemindSurveyModalOpen(true); }} className="px-4 py-2 rounded bg-gray-100">戻る</button>
                                        <button onClick={proceedToRecipients} className="px-4 py-2 rounded text-white bg-blue-600 hover:bg-blue-700">次へ</button>
                                    </div>
                                </div>
                            </div>
                        )}

                        {remindRecipientsModalOpen && (
                            <div className="fixed inset-0 flex items-center justify-center modal-overlay p-3 z-[246]">
                                <div className="bg-white w-full max-w-4xl rounded-xl shadow-2xl flex flex-col h-[92dvh] overflow-hidden">
                                    <div className="p-3 border-b">
                                        <div className="flex items-center justify-between gap-2 mb-2">
                                            <h4 className="font-bold">リマインド対象者</h4>
                                            <button onClick={toggleAllRecipients} className="text-xs px-3 py-1 rounded border bg-gray-50">全選択 / 解除</button>
                                        </div>
                                        <input value={remindSearch} onChange={e => setRemindSearch(e.target.value)} className="w-full border p-2 rounded text-sm" placeholder="名前やメールで検索" />
                                    </div>
                                    <div className="px-3 py-2 bg-blue-600 text-white text-xs font-bold">表示: {filteredReminderUsers.length} / 選択: {remindRecipientSel.size}</div>
                                    <div className="flex-1 overflow-auto bg-gray-100 p-2 space-y-1">
                                        {filteredReminderUsers.map((u, idx) => {
                                            const email = String(u.email || '').toLowerCase();
                                            const checked = remindRecipientSel.has(email);
                                            const cnt = Number(unansweredCountByEmail[email] || 0);
                                            const countClass = cnt > 0 ? 'bg-orange-100 text-orange-700' : 'bg-gray-200 text-gray-600';
                                            return (
                                                <div key={email || idx} onClick={() => toggleRecipient(email)} className={`p-2 rounded-lg border transition-all flex items-center bg-white cursor-pointer ${checked ? 'border-blue-500 bg-blue-50 shadow-sm' : 'border-gray-200'}`}>
                                                    <div className={`w-5 h-5 min-w-[20px] rounded-full border mr-3 flex items-center justify-center ${checked ? 'bg-blue-500 border-blue-500' : 'border-gray-300'}`}>
                                                        {checked && <i className="fas fa-check text-white text-[10px]"></i>}
                                                    </div>
                                                    <div className="overflow-hidden w-full">
                                                        <div className="flex justify-between items-center">
                                                            <div className="font-bold text-gray-800 text-sm truncate">{u.name}</div>
                                                            <div className="text-[10px] bg-white border border-blue-200 text-blue-700 px-1.5 py-0.5 rounded ml-2 whitespace-nowrap flex items-center">{String(u.field || '-')} {String(u.grade || '-')}</div>
                                                        </div>
                                                        <div className="mt-0.5 flex items-center justify-between gap-2">
                                                            <div className="text-xs text-gray-500 whitespace-normal break-words">{u.departmentText || '所属なし'}</div>
                                                            <div className={`text-xs px-2 py-0.5 rounded whitespace-nowrap ${countClass}`}>{cnt}件</div>
                                                        </div>
                                                    </div>
                                                </div>
                                            );
                                        })}
                                    </div>
                                    <div className="p-3 border-t flex gap-2 bg-white">
                                        <button onClick={() => { setRemindRecipientsModalOpen(false); setRemindFilterModalOpen(true); }} className="flex-1 py-3 rounded-lg text-sm font-medium text-gray-600 bg-gray-100">戻る</button>
                                        <button onClick={openConfirmModal} className="flex-[2] bg-blue-600 text-white py-3 rounded-lg font-bold text-sm">確認 ({remindRecipientSel.size})</button>
                                    </div>
                                </div>
                            </div>
                        )}

                        {remindConfirmModalOpen && (
                            <div className="fixed inset-0 flex items-center justify-center modal-overlay p-4 z-[247]">
                                <div className="bg-white rounded-xl shadow-2xl w-full max-w-lg p-6">
                                    <h4 className="text-lg font-bold mb-3">確認</h4>
                                    <p className="text-sm text-gray-700 mb-3">{selectedReminderRecipients.length}名のユーザーへbotから通知を送信します。よろしいですか？</p>
                                    {selectedUsersWithoutUnanswered.length > 0 && (
                                        <p className="text-sm text-red-600 font-bold mb-3">未回答アンケートのないユーザーが通知対象者になっています。本当によろしいですか？</p>
                                    )}
                                    <div className="flex justify-end gap-2">
                                        <button onClick={() => { setRemindConfirmModalOpen(false); setRemindRecipientsModalOpen(true); }} className="px-4 py-2 rounded bg-gray-100">キャンセル</button>
                                        <button onClick={executeReminderSend} className="px-4 py-2 rounded text-white bg-blue-600 hover:bg-blue-700">OK</button>
                                    </div>
                                </div>
                            </div>
                        )}

                        {remindSending && (
                            <div className="fixed inset-0 flex items-center justify-center modal-overlay p-4 z-[248]">
                                <div className="bg-white rounded-xl shadow-2xl w-full max-w-md p-6 flex flex-col items-center">
                                    <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mb-4"></div>
                                    <p className="text-gray-700 text-sm mb-2">送信中...</p>
                                    <p className="text-red-600 font-bold text-sm text-center">送信中は絶対にページを閉じないでください</p>
                                </div>
                            </div>
                        )}

                        {remindResult && <ResultModal result={remindResult} onClose={() => setRemindResult(null)} />}
                        {remindError && !remindSurveyModalOpen && !remindFilterModalOpen && !remindRecipientsModalOpen && !remindConfirmModalOpen && (
                            <div className="text-red-600 text-sm">{remindError}</div>
                        )}
                    </div>
                </div>
            );
        }
