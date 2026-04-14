// @ts-nocheck
'use client';
/* eslint-disable */

import { useEffect, useState } from 'react';

        export default function SurveyTab({ user, runGas, DetailsModal }) {
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
                inChargeOrg: '',
                inChargeDept: '',
                collecting: false,
                scoreName: '',
                scoreUnit: ''
            });

            const formIsValid = (formDraft && typeof formDraft === 'object') ? ((formDraft.title || '').toString().trim().length > 0 && (((formDraft.spreadsheetRef||'').toString().trim().length > 0) || ((formDraft.formUrl||'').toString().trim().length > 0))) : false;

            useEffect(() => {
                setLoading(true); setErr('');
                const token = localStorage.getItem('slack_app_session');
                runGas('listSurveys', token).then(res => {
                    setSurveys(res || []);
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
                    if (Array.isArray(list)) setSurveys(list);
                } catch (e) {
                    alert(e.message || e);
                } finally {
                    setFormSaving(false);
                }
            };

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
                                                            <div className="text-xs text-gray-600">{(s.inChargeOrg || '-') + (s.inChargeDept ? (' / ' + s.inChargeDept) : '')}</div>
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
                                                    <td className="px-3 py-2 border-t text-sm text-gray-700">{(s.inChargeOrg || '-') + (s.inChargeDept ? (' / ' + s.inChargeDept) : '')}</td>
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
                                                    <div className="text-xs text-gray-500 truncate">{(f.inChargeOrg || '-') + (f.inChargeDept ? (' / ' + f.inChargeDept) : '')}</div>
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
                                        <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                                            <div>
                                                <label className="block text-xs text-gray-500 mb-1">担当局</label>
                                                <select className="w-full border p-2 rounded" value={formDraft.inChargeOrg} onChange={e => setFormDraft(prev => ({ ...prev, inChargeOrg: e.target.value, inChargeDept: '' }))}>
                                                    <option value="">選択</option>
                                                    {(formOptions && formOptions.orgs ? Array.from(new Set(formOptions.orgs)) : []).map((o2,idx)=>(<option key={idx} value={o2}>{o2}</option>))}
                                                </select>
                                            </div>
                                            <div>
                                                <label className="block text-xs text-gray-500 mb-1">担当部門</label>
                                                <select className="w-full border p-2 rounded" value={formDraft.inChargeDept} onChange={e => setFormDraft(prev => ({ ...prev, inChargeDept: e.target.value }))} disabled={!(formOptions && formOptions.deptMaster && formDraft.inChargeOrg)}>
                                                    <option value="">選択</option>
                                                    {(formOptions && formOptions.deptMaster ? Array.from(new Set(formOptions.deptMaster.filter(d=>d.org===formDraft.inChargeOrg).map(d=>d.dept))) : []).map((d2,idx)=>(<option key={idx} value={d2}>{d2}</option>))}
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
                    </div>
                </div>
            );
        }

