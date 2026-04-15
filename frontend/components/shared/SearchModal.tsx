// @ts-nocheck
'use client';
/* eslint-disable */


import { useEffect, useState, useMemo } from 'react';

        export default function SearchModal({ onClose, onAdd, currentUserEmail, singleSelect = false, confirmLabel = '追加', runGas }) {

            // 初期値はstatus: 'active'（在籍者のみ表示）
            const [criteria, setCriteria] = useState({ query: '', grade: '', field: '', org: '', dept: '', role: '', status: 'active' });
            const [options, setOptions] = useState({ grades:[], fields:[], orgs:[], roles:[], deptMaster:[] });
            const [availableDepts, setAvailableDepts] = useState([]);
            const [users, setUsers] = useState([]); // サーバーでフィルタ済み
            const [sel, setSel] = useState(new Set());
            const [loading, setLoading] = useState(true);
            const [showFilters, setShowFilters] = useState(false);


            useEffect(() => {
                runGas('getSearchOptions').then(setOptions);
            }, []);

            // criteriaが変わるたびにサーバー検索
            useEffect(() => {
                setLoading(true);
                runGas('searchRecipients', { ...criteria }).then(data => {
                    setUsers((data || []).filter(r => r.email !== currentUserEmail));
                }).finally(() => setLoading(false));
            }, [criteria, currentUserEmail]);

            useEffect(() => {
                if (criteria.org) {
                    const filtered = options.deptMaster.filter(d => d.org === criteria.org).map(d => d.dept);
                    setAvailableDepts([...new Set(filtered)].sort());
                } else {
                    const allDepts = options.deptMaster.map(d => d.dept);
                    setAvailableDepts([...new Set(allDepts)].sort());
                }
            }, [criteria.org, options.deptMaster]);


            const handleFilterChange = (key, value) => {
                let newCriteria = { ...criteria, [key]: value };
                if (key === 'org') newCriteria.dept = '';
                setCriteria(newCriteria);
            };


            // サーバーでフィルタ済み
            const filteredUsers = users;

            const toggle = (m) => {
                if (singleSelect) {
                    const s = new Set();
                    if (!sel.has(m)) s.add(m);
                    setSel(s);
                    return;
                }
                const s = new Set(sel);
                if(s.has(m)) s.delete(m); else s.add(m);
                setSel(s);
            };

            const toggleAll = () => {
                const s = new Set(sel);
                const all = res.length > 0 && res.every(r => sel.has(r.email));
                res.forEach(r => all ? s.delete(r.email) : s.add(r.email));
                setSel(s);
            };

            return (
                <div className="fixed inset-0 flex items-center justify-center modal-overlay p-2 z-[100]">
                    <div className="bg-white w-full max-w-4xl rounded-xl shadow-2xl flex flex-col h-[95dvh] md:h-[90vh] overflow-hidden">
                        <div className="p-3 md:p-4 border-b bg-white shrink-0">
                            <div className="flex justify-between items-center mb-2">
                                <h3 className="font-bold text-gray-700 text-sm md:text-base">ユーザー検索</h3>
                                <button onClick={() => setShowFilters(!showFilters)} className="md:hidden text-blue-600 text-xs font-bold bg-blue-50 px-3 py-1.5 rounded active:bg-blue-100">
                                    <i className={`fas fa-filter mr-1`}></i>
                                    {showFilters ? 'フィルターを閉じる' : 'フィルター設定'}
                                </button>
                            </div>
                            <div className={`${showFilters ? 'grid' : 'hidden'} md:grid grid-cols-2 md:grid-cols-3 gap-2 mb-2 animate-fade-in`}>
                                {[
                                    { k: 'grade', l: '学年', opt: options.grades },
                                    { k: 'field', l: '分野', opt: options.fields },
                                    { k: 'org', l: '局', opt: options.orgs },
                                    { k: 'dept', l: '部門', opt: availableDepts },
                                    { k: 'role', l: '役職', opt: options.roles }
                                ].map(f => (
                                    <select key={f.k} className="border p-2 rounded text-xs md:text-sm bg-gray-50 outline-none" value={criteria[f.k]} onChange={e=>handleFilterChange(f.k, e.target.value)}>
                                        <option value="">{f.l}: ALL</option>
                                        {f.opt.map(o=><option key={o} value={o}>{o}</option>)}
                                    </select>
                                ))}
                                {/* 在籍状況フィルタ（デフォルト: 在籍） */}
                                <select key="status" className="border p-2 rounded text-xs md:text-sm bg-gray-50 outline-none" value={criteria.status} onChange={e=>handleFilterChange('status', e.target.value)}>
                                    <option value="active">在籍者のみ</option>
                                    <option value="retired">退局者のみ</option>
                                    <option value="all">在籍状況: ALL</option>
                                </select>
                            </div>
                            <input className="w-full border p-2 md:p-2.5 rounded text-sm md:text-base focus:ring-2 focus:ring-blue-400 outline-none bg-gray-50" value={criteria.query} onChange={e=>handleFilterChange('query', e.target.value)} placeholder="名前やメールで検索..." />
                        </div>
                        <div className="px-3 py-2 bg-blue-600 flex justify-between text-xs font-bold text-white uppercase tracking-wider shrink-0 items-center">
                            <span>検索結果: {filteredUsers.length}</span>
                            <button onClick={toggleAll} className="bg-white/20 px-3 py-1 rounded hover:bg-white/30 transition active:bg-white/40">全選択 / 解除</button>
                        </div>
                        <div className="flex-1 overflow-y-auto p-2 space-y-1 bg-gray-100">
                            {loading ? (
                                <div className="text-center py-10 text-gray-500"><i className="fas fa-spinner fa-spin mr-2"></i></div>
                            ) : filteredUsers.map(r => (
                                <div key={r.email} onClick={()=>toggle(r.email)} className={`p-2 rounded-lg border transition-all flex items-center bg-white ${sel.has(r.email)?'border-blue-500 bg-blue-50 shadow-sm':'border-gray-200'} active:bg-gray-50`}>
                                    <div className={`w-5 h-5 min-w-[20px] rounded-full border mr-3 flex items-center justify-center ${sel.has(r.email)?'bg-blue-500 border-blue-500':'border-gray-300'}`}>
                                        {sel.has(r.email) && <i className="fas fa-check text-white text-[10px]"></i>}
                                    </div>
                                    <div className="overflow-hidden w-full">
                                        <div className="flex justify-between items-center">
                                            <div className="font-bold text-gray-800 text-sm truncate">{r.name}</div>
                                            <div className="text-[10px] bg-white border border-blue-200 text-blue-700 px-1.5 py-0.5 rounded ml-2 whitespace-nowrap flex items-center">
                                                <i className="fas fa-school mr-1"></i>{r.grade} {r.field}
                                            </div>
                                        </div>
                                        <div className="text-xs text-gray-500 mt-0.5 whitespace-normal break-words">
                                            {r.departmentText || [
                                                Array.isArray(r.org) ? r.org.join(', ') : r.org,
                                                Array.isArray(r.department) ? r.department.join(', ') : r.department,
                                                Array.isArray(r.role) ? r.role.join(', ') : r.role
                                            ].filter(Boolean).join(' / ')}
                                        </div>
                                    </div>
                                </div>
                            ))}
                            {!loading && filteredUsers.length === 0 && <div className="text-center py-10 text-gray-400 text-xs">該当なし</div>}
                        </div>
                        <div className="p-3 md:p-4 border-t flex space-x-2 bg-white shadow-[0_-2px_10px_rgba(0,0,0,0.05)] shrink-0 pb-safe">
                            <button onClick={onClose} className="flex-1 py-3 rounded-lg text-sm font-medium text-gray-600 transition-colors bg-gray-100 active:bg-gray-200">キャンセル</button>
                            <button onClick={()=>onAdd(filteredUsers.filter(r=>sel.has(r.email)))} className="flex-[2] bg-blue-600 text-white py-3 rounded-lg font-bold shadow-lg hover:bg-blue-700 text-sm transition-all active:scale-95 active:bg-blue-800">{confirmLabel} ({sel.size})</button>
                        </div>
                    </div>
                </div>
            );
        }
