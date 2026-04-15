// @ts-nocheck
'use client';
/* eslint-disable */

import { useEffect, useRef, useState } from 'react';
import SearchModal from '../shared/SearchModal';

        export default function MyPageTab({ runGas, DialogModal }) {
            const [profile, setProfile] = useState(null);
            const [savedProfile, setSavedProfile] = useState(null);
            const [options, setOptions] = useState(null);
            const [loading, setLoading] = useState(true);
            const [continueSwitchEnabled, setContinueSwitchEnabled] = useState(false);
            const [saving, setSaving] = useState(false);
            const [saveDialog, setSaveDialog] = useState({ isOpen: false, type: '', message: '', onOk: null, onCancel: null });
            const [msg, setMsg] = useState({ type: '', text: '' });
            const [userEditOpen, setUserEditOpen] = useState(false);
            const [creating, setCreating] = useState(false);
            const [editingTarget, setEditingTarget] = useState(null);

            const formRef = useRef(null);
            const nameRef = useRef(null);
            const saveBtnRef = useRef(null);

            const normalizeOrgs = (arr) => {
                const src = Array.isArray(arr) ? arr.slice() : [];
                while (src.length < 10) src.push({ org: '', dept: '', role: '' });
                return src.slice(0, 10);
            };

            useEffect(() => {
                Promise.all([
                    runGas('getUserProfile', localStorage.getItem('slack_app_session')),
                    runGas('getSearchOptions')
                ]).then(([p, o]) => {
                    const normalized = { ...p, orgs: normalizeOrgs(p && p.orgs) };
                    setProfile(normalized);
                    setSavedProfile(normalized);
                    setOptions(o);
                    runGas('isContinueSwitchEnabled').then(r=>setContinueSwitchEnabled(!!r)).catch(()=>{});
                }).catch(e => setMsg({ type: 'error', text: e.message }))
                .finally(() => setLoading(false));
            }, []);

            const handleFormKeyDown = (e) => {
                if (e.key !== 'Enter') return;
                const root = formRef.current;
                if (!root) return;
                const tag = e.target && e.target.tagName && e.target.tagName.toLowerCase();
                if (!['input','select','textarea'].includes(tag)) return;
                e.preventDefault();
                const focusables = Array.from(root.querySelectorAll('input:not([disabled]), select:not([disabled]), textarea:not([disabled])'))
                    .filter(el => el.type !== 'hidden');
                const idx = focusables.indexOf(e.target);
                if (idx === -1) return;
                const next = focusables[idx + 1];
                if (next) next.focus();
            };

            const handleNameEnBlur = (e) => {
                if (!profile) return;
                let val = (e.target.value || '').toString();
                // 全角スペースを半角に
                val = val.replace(/\u3000/g, ' ').replace(/\s+/g, ' ').trim();
                // update normalized name
                setProfile(prev => ({ ...prev, nameEn: val }));

                // only auto-create when in creating mode
                if (!creating) return;
                // do not overwrite existing email
                if (profile.email) return;
                if (!val) return;

                const parts = val.split(' ');
                // expect at least surname and given name: "姓 名"
                if (parts.length < 2) return;
                const surname = parts[0];
                const given = parts[1];
                const alpha = /^[A-Za-z-']+$/;
                if (!alpha.test(surname) || !alpha.test(given)) return;

                const yy = String(new Date().getFullYear()).slice(-2);
                const initial = given[0].toLowerCase();
                const cleanSurname = surname.toLowerCase().replace(/[^a-z]/g, '');
                const email = `${yy}.${initial}.${cleanSurname}.nutfes@gmail.com`;
                setProfile(prev => ({ ...prev, email }));
            };

            const parseBirthday = (b) => {
                if (!b) return { y: '', m: '', d: '' };
                const parts = String(b).split('-');
                return { y: parts[0] || '', m: parts[1] || '', d: parts[2] || '' };
            };

            const handleBirthdayPartChange = (part, value) => {
                const { y, m, d } = parseBirthday(profile.birthday || '');
                const ny = part === 'y' ? value.replace(/\D/g,'').slice(0,4) : y;
                const nm = part === 'm' ? value.replace(/\D/g,'').slice(0,2) : m;
                const nd = part === 'd' ? value.replace(/\D/g,'').slice(0,2) : d;
                // Compose without automatic left-zero padding so partial inputs remain as-typed
                const parts = [];
                if (ny) parts.push(ny);
                if (nm) parts.push(nm);
                if (nd) parts.push(nd);
                const composed = parts.join('-');
                setProfile(prev => ({ ...prev, birthday: composed }));
            };

            const handleStudentIdInput = (e) => {
                const v = (e.target.value || '').replace(/\D/g, '').slice(0,8);
                setProfile(prev => ({ ...prev, studentId: v }));
            };

            const validatePhoneNumber = (phone) => {
                if (!phone) return true; // allow empty
                const digits = (phone || '').replace(/\D/g, '');
                return digits.length >= 10 && digits.length <= 11;
            };

            const handleChange = (key, val) => setProfile(prev => ({ ...prev, [key]: val }));

            const handleOrgChange = (idx, key, val) => {
                const newOrgs = [...profile.orgs];
                newOrgs[idx] = { ...newOrgs[idx], [key]: val };
                setProfile(prev => ({ ...prev, orgs: newOrgs }));
            };

            const orgMaster = (options && Array.isArray(options.orgMaster) && options.orgMaster.length > 0)
                ? options.orgMaster
                : ((options && Array.isArray(options.orgs)) ? options.orgs.map(o => ({ org: o, notMain: false })) : []);
            const roleMaster = (options && Array.isArray(options.roleMaster) && options.roleMaster.length > 0)
                ? options.roleMaster
                : ((options && Array.isArray(options.roles)) ? options.roles.map(r => ({ role: r, notMain: false })) : []);

            const EMPTY_NEW_PROFILE = () => ({
                name: '', nameEn: '', email: '', studentId: '', grade: '', field: '', phone: '', birthday: '', almaMater: '', carOwner: false, isAdmin: false,
                orgs: [
                    {org:'',dept:'',role:''},{org:'',dept:'',role:''},{org:'',dept:'',role:''},{org:'',dept:'',role:''},{org:'',dept:'',role:''},
                    {org:'',dept:'',role:''},{org:'',dept:'',role:''},{org:'',dept:'',role:''},{org:'',dept:'',role:''},{org:'',dept:'',role:''}
                ],
                canEditNameEmail: true
            });

            const handleSave = async () => {
                setSaving(true); setMsg({ type: '', text: '' });
                // show loading modal
                setSaveDialog({ isOpen: true, type: 'loading', message: creating ? '作成中...' : '保存中...', onOk: null, onCancel: null });
                try {
                    const token = localStorage.getItem('slack_app_session');
                    if (creating) {
                        // create new user
                        const payload = { ...profile };
                        payload.orgs = payload.orgs.filter(o => o.org || o.dept || o.role);
                        await runGas('createUser', token, payload);
                        // refresh options in case admin lists changed
                        runGas('getSearchOptions').then(o=>setOptions(o)).catch(()=>{});
                        // reset form for next new registration
                        setProfile(EMPTY_NEW_PROFILE());
                        setEditingTarget('__new__');
                        setCreating(true);
                        // close loading, show success modal and focus name for next entry
                        setSaveDialog({ isOpen: false, type: '', message: '' });
                        setSaveDialog({ isOpen: true, type: 'notify', message: '新規ユーザーを作成しました', onOk: () => { setSaveDialog({ isOpen: false }); if (nameRef.current) nameRef.current.focus(); }, onCancel: null });
                    } else {
                        // validations for edit/save
                        const sid = (profile.studentId || '').replace(/\D/g, '');
                        if (sid && sid.length !== 8) {
                            setSaveDialog({ isOpen: true, type: 'notify', message: '学籍番号は8桁の数字で入力してください', onOk: () => setSaveDialog({ isOpen: false }), onCancel: null });
                            setSaving(false); return;
                        }
                        if (!validatePhoneNumber(profile.phone)) {
                            setSaveDialog({ isOpen: true, type: 'notify', message: '電話番号は数字10〜11桁で入力してください', onOk: () => setSaveDialog({ isOpen: false }), onCancel: null });
                            setSaving(false); return;
                        }
                        // birthday validation
                        if (profile.birthday) {
                            const parts = (profile.birthday || '').split('-');
                            if (parts.length === 3) {
                                const y = parseInt(parts[0],10), m = parseInt(parts[1],10), d = parseInt(parts[2],10);
                                const dt = new Date(y,m-1,d);
                                if (dt.getFullYear() !== y || dt.getMonth() !== m-1 || dt.getDate() !== d) {
                                    setSaveDialog({ isOpen: true, type: 'notify', message: '生年月日の形式が正しくありません', onOk: () => setSaveDialog({ isOpen: false }), onCancel: null });
                                    setSaving(false); return;
                                }
                            }
                        }

                        await runGas('updateUserProfile', token, profile, editingTarget);
                        // 保存済みプロフィールを更新（制限判定に使用）
                        setSavedProfile(profile);
                        // close loading, show success modal
                        setSaveDialog({ isOpen: false, type: '', message: '' });
                        setSaveDialog({ isOpen: true, type: 'notify', message: '保存しました', onOk: () => setSaveDialog({ isOpen: false }), onCancel: null });
                    }
                } catch(e) {
                    setSaveDialog({ isOpen: false, type: '', message: '' });
                    setSaveDialog({ isOpen: true, type: 'notify', message: (creating ? '作成失敗: ' : '保存失敗: ') + e.message, onOk: () => setSaveDialog({ isOpen: false }), onCancel: null });
                } finally { setSaving(false); }
            };

            const openUserEdit = () => setUserEditOpen(true);
            const loadUserForEdit = async (selectedEmail) => {
                try {
                    // モーダルを閉じてタブで読み込み表示を行う
                    setUserEditOpen(false);
                    setLoading(true);
                    const token = localStorage.getItem('slack_app_session');
                    const p = await runGas('getUserProfile', token, selectedEmail);
                    const normalized = { ...p, orgs: normalizeOrgs(p && p.orgs) };
                    setProfile(normalized);
                    setSavedProfile(normalized);
                    setEditingTarget(selectedEmail);
                    setCreating(false);
                } catch(e) { setMsg({ type: 'error', text: '読み込み失敗: ' + e.message }); }
                finally { setLoading(false); }
            };

            const restoreSelf = async () => {
                try {
                    setLoading(true);
                    const token = localStorage.getItem('slack_app_session');
                    const p = await runGas('getUserProfile', token);
                    const normalized = { ...p, orgs: normalizeOrgs(p && p.orgs) };
                    setProfile(normalized);
                    setSavedProfile(normalized);
                    setEditingTarget(null);
                    setCreating(false);
                } catch(e) { setMsg({ type: 'error', text: '読み込み失敗: ' + e.message }); }
                finally { setLoading(false); }
            };

            if (loading) return <div className="flex justify-center items-center h-full text-gray-500"><i className="fas fa-spinner fa-spin mr-2"></i>読み込み中...</div>;
            if (!profile || !options) return <div className="p-4 text-red-500">データ読み込みエラー</div>;

            return (
                <div className="h-full overflow-y-auto p-4 md:p-6 bg-gray-50">
                    <div ref={formRef} onKeyDown={handleFormKeyDown} className="max-w-3xl mx-auto bg-white rounded-lg shadow-sm p-6">
                        <div className="flex items-center justify-between mb-6 border-b pb-2">
                            <h2 className="text-xl font-bold text-gray-800"><i className="fas fa-user-edit mr-2"></i>プロフィール編集</h2>
                            <div className="flex items-center space-x-2">
                                {profile && profile.canEditNameEmail && (
                                    <>
                                        {editingTarget && <button onClick={restoreSelf} className="text-xs text-gray-600 bg-gray-100 px-3 py-1 rounded">自分に戻す</button>}
                                        <button onClick={()=>{ setCreating(true); setEditingTarget('__new__'); setProfile(EMPTY_NEW_PROFILE()); }} className="text-xs bg-green-100 text-green-900 px-3 py-1 rounded hover:brightness-95">新規登録</button>
                                        <button onClick={openUserEdit} className="text-xs bg-yellow-100 text-yellow-900 px-3 py-1 rounded hover:brightness-95">ユーザー編集</button>
                                    </>
                                )}
                            </div>
                        </div>

                        {msg.text && (
                            <div className={`mb-4 p-3 rounded text-sm ${msg.type==='success'?'bg-green-50 text-green-700':'bg-red-50 text-red-700'}`}>
                                {msg.text}
                            </div>
                        )}

                        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-6">
                            <div>
                                <label className="block text-xs font-bold text-gray-500 mb-1">氏名</label>
                                <input ref={nameRef} type="text" value={profile.name} onChange={e=>handleChange('name', e.target.value)} disabled={!profile.canEditNameEmail} placeholder="山田 太郎" className={`w-full border p-2 rounded text-sm ${profile.canEditNameEmail? 'focus:ring-2 focus:ring-blue-400' : 'bg-gray-100 text-gray-500'}`} />
                            </div>
                            <div>
                                <label className="block text-xs font-bold text-gray-700 mb-1">英語の氏名（姓 半角スペース 名）</label>
                                <input type="text" value={profile.nameEn} onChange={e=>handleChange('nameEn', e.target.value)} onBlur={handleNameEnBlur} placeholder="Yamada Taro" className="w-full border p-2 rounded text-sm focus:ring-2 focus:ring-blue-400" />
                            </div>

                            <div>
                                <label className="block text-xs font-bold text-gray-500 mb-1">メールアドレス</label>
                                <input type="text" value={profile.email} onChange={e=>handleChange('email', e.target.value)} disabled={!profile.canEditNameEmail} placeholder="example@ex.com" className={`w-full border p-2 rounded text-sm ${profile.canEditNameEmail? 'focus:ring-2 focus:ring-blue-400' : 'bg-gray-100 text-gray-500'}`} />
                            </div>

                            <div>
                                <label className="block text-xs font-bold text-gray-500 mb-1">学籍番号</label>
                                <input inputMode="numeric" pattern="\\d{8}" maxLength={8} value={profile.studentId} onInput={handleStudentIdInput} placeholder="12345678" className="w-full border p-2 rounded text-sm focus:ring-2 focus:ring-blue-400" />
                            </div>
                            <div>
                                <label className="block text-xs font-bold text-gray-700 mb-1">電話番号</label>
                                <input type="tel" value={profile.phone} onChange={e=>handleChange('phone', e.target.value)} placeholder="09012345678" className="w-full border p-2 rounded text-sm focus:ring-2 focus:ring-blue-400" />
                            </div>

                            <div className="grid grid-cols-2 gap-2">
                                <div>
                                    <label className="block text-xs font-bold text-gray-700 mb-1">学年</label>
                                    <select value={profile.grade} onChange={e=>handleChange('grade', e.target.value)} className="w-full border p-2 rounded text-sm bg-white">
                                        <option value="">選択</option>
                                        {options.grades.map(o=><option key={o} value={o}>{o}</option>)}
                                    </select>
                                </div>
                                <div>
                                    <label className="block text-xs font-bold text-gray-700 mb-1">分野</label>
                                    <select value={profile.field} onChange={e=>handleChange('field', e.target.value)} className="w-full border p-2 rounded text-sm bg-white">
                                        <option value="">選択</option>
                                        {options.fields.map(o=><option key={o} value={o}>{o}</option>)}
                                    </select>
                                </div>
                            </div>

                            <div>
                                <label className="block text-xs font-bold text-gray-700 mb-1">生年月日</label>
                                <div className="flex space-x-2">
                                    <input inputMode="numeric" maxLength={4} placeholder="YYYY" value={parseBirthday(profile.birthday).y} onChange={e=>handleBirthdayPartChange('y', e.target.value)} className="w-1/3 border p-2 rounded text-sm focus:ring-2 focus:ring-blue-400" />
                                    <input inputMode="numeric" maxLength={2} placeholder="MM" value={parseBirthday(profile.birthday).m} onChange={e=>handleBirthdayPartChange('m', e.target.value)} className="w-1/4 border p-2 rounded text-sm focus:ring-2 focus:ring-blue-400" />
                                    <input inputMode="numeric" maxLength={2} placeholder="DD" value={parseBirthday(profile.birthday).d} onChange={e=>handleBirthdayPartChange('d', e.target.value)} className="w-1/4 border p-2 rounded text-sm focus:ring-2 focus:ring-blue-400" />
                                </div>
                            </div>
                             <div>
                                <label className="block text-xs font-bold text-gray-700 mb-1">出身校</label>
                                <input type="text" value={profile.almaMater} onChange={e=>handleChange('almaMater', e.target.value)} placeholder="○○工業高等専門学校△△キャンパス" className="w-full border p-2 rounded text-sm focus:ring-2 focus:ring-blue-400" />
                            </div>

                            <div>
                                <label className="block text-xs font-bold text-gray-700 mb-1">車所有</label>
                                <div className="flex items-center space-x-3 mt-2">
                                    <button
                                        type="button"
                                        onClick={() => handleChange('carOwner', !profile.carOwner)}
                                        className={`relative inline-flex h-8 w-14 items-center rounded-full transition-colors ${
                                            profile.carOwner ? 'bg-green-600' : 'bg-gray-300'
                                        }`}
                                    >
                                        <span
                                            className={`inline-block h-6 w-6 transform rounded-full bg-white transition-transform ${
                                                profile.carOwner ? 'translate-x-7' : 'translate-x-1'
                                            }`}
                                        />
                                    </button>
                                    <span className="text-sm text-gray-700 font-medium">{profile.carOwner ? '所有' : 'なし'}</span>
                                </div>
                            </div>
                        </div>

                        <div className="mt-3">
                            <label className="block text-xs font-bold text-gray-700 mb-1">退局</label>
                            <div className="flex items-center space-x-3 mt-2">
                                <button
                                    type="button"
                                    onClick={() => {
                                            if (!profile) return;
                                            // 管理者は両方向で変更可能
                                            if (profile.canEditNameEmail) {
                                                handleChange('retired', !profile.retired);
                                                return;
                                            }
                                            // 非管理者: 保存済みの値が在籍(false) の場合のみ保存まで編集可能
                                            if (!savedProfile || savedProfile.retired === false) {
                                                handleChange('retired', !profile.retired);
                                            }
                                        }}
                                        // enabled when editorIsAdmin OR savedProfile indicates not retired (so non-admin can set to retired until saved)
                                        disabled={!(profile && (profile.canEditNameEmail || !savedProfile || savedProfile.retired === false))}
                                    className={`relative inline-flex h-8 w-14 items-center rounded-full transition-colors ${
                                        profile.retired ? 'bg-red-600' : 'bg-gray-300'
                                    }`}
                                >
                                    <span className={`inline-block h-6 w-6 transform rounded-full bg-white transition-transform ${profile.retired ? 'translate-x-7' : 'translate-x-1'}`} />
                                </button>
                                <span className="text-sm text-gray-700 font-medium">{profile.retired ? '退局' : '在籍'}</span>
                            </div>
                            <div className="text-xs text-gray-500 mt-1">(一度有効化すると管理者以外変更できません。)</div>
                        </div>

                        <div className="mt-3">
                            <label className="block text-xs font-bold text-gray-700 mb-1">次年度も継続</label>
                            <div className="flex items-center space-x-3 mt-2">
                                <button
                                    type="button"
                                    onClick={() => {
                                        if (!continueSwitchEnabled) return;
                                        if (!profile || !profile.canEditNameEmail) return;
                                        handleChange('continueNext', !profile.continueNext);
                                    }}
                                    disabled={!continueSwitchEnabled || !(profile && profile.canEditNameEmail)}
                                    className={`relative inline-flex h-8 w-14 items-center rounded-full transition-colors ${
                                        profile.continueNext ? 'bg-green-600' : 'bg-gray-300'
                                    }`}
                                >
                                    <span className={`inline-block h-6 w-6 transform rounded-full bg-white transition-transform ${profile.continueNext ? 'translate-x-7' : 'translate-x-1'}`} />
                                </button>
                                <span className="text-sm text-gray-700 font-medium">{profile.continueNext ? '継続する' : '未設定'}</span>
                            </div>
                            <div className="text-xs text-gray-500 mt-1">(期間外は操作不可・一度有効化すると管理者以外変更できません。)</div>
                        </div>

                        {profile && profile.canEditNameEmail && (
                            <div className="mt-3">
                                <label className="block text-xs font-bold text-gray-700 mb-1">管理者権限</label>
                                <div className="flex items-center space-x-3 mt-2">
                                    <button
                                        type="button"
                                        onClick={() => handleChange('isAdmin', !profile.isAdmin)}
                                        className={`relative inline-flex h-8 w-14 items-center rounded-full transition-colors ${
                                            profile.isAdmin ? 'bg-green-600' : 'bg-gray-300'
                                        }`}
                                    >
                                        <span
                                            className={`inline-block h-6 w-6 transform rounded-full bg-white transition-transform ${
                                                profile.isAdmin ? 'translate-x-7' : 'translate-x-1'
                                            }`}
                                        />
                                    </button>
                                    <span className="text-sm text-gray-700 font-medium">{profile.isAdmin ? '管理者' : 'なし'}</span>
                                </div>
                            </div>
                        )}
                        <br/><br/>
                        <h3 className="text-sm font-bold text-gray-700 mb-3 border-b pb-1">所属情報 (最大10つ)</h3>
                        <div className="text-xs text-gray-500 mt-1">(1つ目の所属先は主となる所属先です。主所属局は管理者以外変更できません。<br/>兼局先も登録してください。<br/>執行部は兼局先として登録してください。)</div>
                        <div className="space-y-3 mb-6">
                            {profile.orgs.map((orgData, idx) => (
                                <div key={idx} className="flex flex-col md:flex-row gap-2 p-3 bg-gray-50 rounded border">
                                    <div className="flex-1">
                                        <label className="text-[10px] text-gray-500">{idx === 0 ? (profile && profile.canEditNameEmail ? '所属局' : '所属局 (変更不可)') : '所属局'}</label>
                                        <select value={orgData.org} onChange={e=>handleOrgChange(idx, 'org', e.target.value)} disabled={idx === 0 && !(profile && profile.canEditNameEmail)} className={`w-full border p-1 rounded text-sm ${(idx === 0 && !(profile && profile.canEditNameEmail)) ? 'bg-gray-100 text-gray-500' : 'bg-white'}`}>
                                            <option value="">なし</option>
                                            {orgMaster.filter(o => idx !== 0 || !o.notMain || o.org === orgData.org).map(o=><option key={o.org} value={o.org}>{o.org}</option>)}
                                        </select>
                                    </div>
                                    <div className="flex-1">
                                        <label className="text-[10px] text-gray-500">所属部門</label>
                                        <select value={orgData.dept} onChange={e=>handleOrgChange(idx, 'dept', e.target.value)} className="w-full border p-1 rounded text-sm bg-white">
                                            <option value="">なし</option>
                                            {options.deptMaster.filter(d=>(!orgData.org || d.org===orgData.org) && (idx !== 0 || !d.notMain || d.pid === orgData.dept)).map(d=><option key={d.pid} value={d.pid}>{d.dept}</option>)}
                                        </select>
                                    </div>
                                    <div className="flex-1">
                                        <label className="text-[10px] text-gray-500">役職</label>
                                        <select value={orgData.role} onChange={e=>handleOrgChange(idx, 'role', e.target.value)} className="w-full border p-1 rounded text-sm bg-white">
                                            <option value="">なし</option>
                                            {roleMaster.filter(r => idx !== 0 || !r.notMain || r.role === orgData.role).map(r=><option key={r.role} value={r.role}>{r.role}</option>)}
                                        </select>
                                    </div>
                                </div>
                            ))}
                        </div>

                        <div className="text-right">
                            <button ref={saveBtnRef} data-save-button onClick={handleSave} disabled={saving} className="bg-blue-600 text-white px-8 py-3 rounded-lg font-bold shadow-lg hover:bg-blue-700 transition active:scale-95 disabled:bg-gray-400">
                                {saving ? <i className="fas fa-spinner fa-spin"></i> : '保存する'}
                            </button>
                        </div>
                        {userEditOpen && (
                            <SearchModal runGas={runGas} currentUserEmail={profile.email} singleSelect={true} confirmLabel={"選択"} onClose={()=>setUserEditOpen(false)} onAdd={(ls)=>{
                                if (ls && ls.length > 0) loadUserForEdit(ls[0].email);
                                else setUserEditOpen(false);
                            }} />
                        )}
                        {/* 新規登録はマイページ内の作成モードで実装するためモーダルは削除しました */}
                        <DialogModal isOpen={saveDialog.isOpen} type={saveDialog.type === 'notify' ? 'notify' : saveDialog.type} message={saveDialog.message} onOk={saveDialog.onOk} onCancel={saveDialog.onCancel} />
                    </div>
                </div>
            );
        }
