'use client';

import { useEffect, useMemo, useState } from 'react';

const API_TIMEOUT_MS = 60000;

type OrgItem = { pid: string; org: string; status?: string; not_main_org?: boolean; gen?: number | string };
type DeptItem = { pid: string; dept: string; orgPid: string; org?: string; status?: string; not_main_dept?: boolean };
type RoleItem = { pid: string; role: string; gen?: number | string; not_main_role?: boolean };

const isActiveStatus = (status?: string) => {
  const value = String(status || '').trim();
  return value === '' || value === '0' || value.toLowerCase() === 'false';
};

const statusLabel = (status?: string) => (isActiveStatus(status) ? '有効' : '無効');

const padOrgIdForDisplay = (pid?: string) => {
  const s = String(pid || '').trim();
  return /^\d$/.test(s) ? `0${s}` : s;
};

const padDeptIdForDisplay = (pid?: string) => {
  const s = String(pid || '').trim();
  if (!/^\d+$/.test(s)) return s;
  if (s.length >= 4) return s;
  return s.padStart(4, '0');
};

function runGas(funcName: string, ...args: any[]) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), API_TIMEOUT_MS);
  let payload: any = {};

  switch (funcName) {
    case 'getLoginUser':
    case 'getUserProfile':
    case 'listAffiliationMasters':
      payload.sessionToken = args[0];
      break;
    case 'saveOrgMaster':
    case 'saveDeptMaster':
    case 'saveRoleMaster':
      payload.sessionToken = args[0];
      payload.payload = args[1] || {};
      break;
    default:
      payload.__args = args;
  }

  return fetch('/api/gas', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ action: funcName, payload }),
    cache: 'no-store',
    signal: controller.signal
  })
    .then(async (res) => {
      const json = await res.json();
      if (!json || !json.ok) throw new Error((json && json.error) || 'API request failed');
      return json.result;
    })
    .finally(() => clearTimeout(timer));
}

export default function AdminPage() {
  const [loading, setLoading] = useState(true);
  const [authorized, setAuthorized] = useState(false);
  const [savingKind, setSavingKind] = useState<'org' | 'dept' | 'role' | null>(null);
  const [activeMenu, setActiveMenu] = useState<'org' | 'dept' | 'role'>('org');
  const [orgs, setOrgs] = useState<OrgItem[]>([]);
  const [depts, setDepts] = useState<DeptItem[]>([]);
  const [roles, setRoles] = useState<RoleItem[]>([]);
  const [error, setError] = useState('');

  const [orgModal, setOrgModal] = useState<{ open: boolean; item: OrgItem | null }>({ open: false, item: null });
  const [deptModal, setDeptModal] = useState<{ open: boolean; item: DeptItem | null }>({ open: false, item: null });
  const [roleModal, setRoleModal] = useState<{ open: boolean; item: RoleItem | null }>({ open: false, item: null });
  const [orgNotMainChecked, setOrgNotMainChecked] = useState(false);
  const [orgGenValue, setOrgGenValue] = useState('');
  const [deptOrgPid, setDeptOrgPid] = useState('');
  const [deptNotMainChecked, setDeptNotMainChecked] = useState(false);

  const token = useMemo(() => {
    if (typeof window === 'undefined') return '';
    return localStorage.getItem('slack_app_session') || '';
  }, []);

  const reload = async () => {
    const result = await runGas('listAffiliationMasters', token);
    setOrgs(Array.isArray(result?.orgs) ? result.orgs : []);
    setDepts(Array.isArray(result?.depts) ? result.depts : []);
    setRoles(Array.isArray(result?.roles) ? result.roles : []);
  };

  const sortByPid = (items: any[]) => {
    return [...items].sort((a, b) => {
      const aPid = parseInt(String(a.pid || '0'), 10);
      const bPid = parseInt(String(b.pid || '0'), 10);
      return aPid - bPid;
    });
  };

  const isSaving = (kind: 'org' | 'dept' | 'role') => savingKind === kind;

  const submitAndReload = async (kind: 'org' | 'dept' | 'role', action: string, payload: any, closeModal: () => void) => {
    setSavingKind(kind);
    setError('');
    try {
      await runGas(action, token, payload);
      await reload();
      closeModal();
    } catch (err: any) {
      setError(err?.message || String(err));
    } finally {
      setSavingKind(null);
    }
  };

  useEffect(() => {
    (async () => {
      try {
        const login = await runGas('getLoginUser', token);
        if (!login || login.status !== 'authorized') {
          setAuthorized(false);
          setLoading(false);
          return;
        }
        const profile = await runGas('getUserProfile', token);
        if (!profile || !profile.isAdmin) {
          setAuthorized(false);
          setLoading(false);
          return;
        }
        setAuthorized(true);
        await reload();
      } catch (e: any) {
        setError(e?.message || String(e));
      } finally {
        setLoading(false);
      }
    })();
  }, [token]);

  useEffect(() => {
    if (!orgModal.open) return;
    const checked = !!orgModal.item?.not_main_org;
    setOrgNotMainChecked(checked);
    setOrgGenValue(checked ? String(orgModal.item?.gen ?? '') : '');
  }, [orgModal.open, orgModal.item]);

  useEffect(() => {
    if (!deptModal.open) return;
    const initialOrgPid = String(deptModal.item?.orgPid || '');
    const initialNotMain = !!deptModal.item?.not_main_dept;
    const selectedOrg = orgs.find((o) => String(o.pid) === initialOrgPid);
    const forceNotMain = !!selectedOrg?.not_main_org;
    setDeptOrgPid(initialOrgPid);
    setDeptNotMainChecked(forceNotMain ? true : initialNotMain);
  }, [deptModal.open, deptModal.item, orgs]);

  const saveOrg = async (e: any) => {
    e.preventDefault();
    const f = new FormData(e.currentTarget);
    setError('');
    const pid = String(f.get('pid') || '');
    const orgName = String(f.get('org') || '').trim();
    const status = String(f.get('status') || '');
    const notMain = !!f.get('not_main_org');
    const genRaw = String(f.get('gen') || '').trim();

    // 入力値の検証
    if (!orgName) {
      setError('局名を入力してください');
      return;
    }
    if (notMain && genRaw === '') {
      setError('主所属局外 (特命局) の場合、gen は必須です');
      return;
    }

    const gen = genRaw === '' ? undefined : Number(genRaw);
    await submitAndReload('org', 'saveOrgMaster', {
      pid: pid,
      org: orgName,
      gen: gen,
      status: status,
      not_main_org: notMain
    }, () => setOrgModal({ open: false, item: null }));
  };

  const saveDept = async (e: any) => {
    e.preventDefault();
    const f = new FormData(e.currentTarget);
    const deptName = String(f.get('dept') || '').trim();
    const orgPid = String(f.get('orgPid') || '').trim();
    const selectedOrg = orgs.find((o) => String(o.pid) === orgPid);
    const notMainDept = selectedOrg?.not_main_org ? true : !!f.get('not_main_dept');

    // 入力値の検証
    if (!deptName) {
      setError('部門名を入力してください');
      return;
    }
    if (!orgPid) {
      setError('所属局を選択してください');
      return;
    }
    setError('');

    await submitAndReload('dept', 'saveDeptMaster', {
      pid: String(f.get('pid') || ''),
      dept: deptName,
      orgPid: orgPid,
      status: String(f.get('status') || ''),
      not_main_dept: notMainDept
    }, () => setDeptModal({ open: false, item: null }));
  };

  const saveRole = async (e: any) => {
    e.preventDefault();
    const f = new FormData(e.currentTarget);
    const roleName = String(f.get('role') || '').trim();
    const genRaw = String(f.get('gen') || '').trim();

    // 入力値の検証
    if (!roleName) {
      setError('役職名を入力してください');
      return;
    }
    if (!genRaw) {
      setError('genを入力してください');
      return;
    }
    setError('');

    const gen = Number(genRaw);
    await submitAndReload('role', 'saveRoleMaster', {
      pid: String(f.get('pid') || ''),
      role: roleName,
      gen: gen,
      not_main_role: !!f.get('not_main_role')
    }, () => setRoleModal({ open: false, item: null }));
  };

  if (loading) return <div className="p-8">読み込み中...</div>;
  if (!authorized) return <div className="p-8 text-red-600">管理者のみアクセス可能です</div>;

  return (
    <div className="min-h-screen bg-gray-100 p-4">
      <div className="mx-auto max-w-7xl bg-white shadow rounded-xl overflow-hidden flex min-h-[80vh]">
        <aside className="w-64 border-r bg-gray-50 p-4">
          <div className="text-sm text-gray-500 mb-3">所属管理</div>
          <button onClick={() => setActiveMenu('org')} className={`w-full text-left px-3 py-2 rounded mb-2 ${activeMenu === 'org' ? 'bg-blue-600 text-white' : 'hover:bg-gray-200'}`}>局</button>
          <button onClick={() => setActiveMenu('dept')} className={`w-full text-left px-3 py-2 rounded mb-2 ${activeMenu === 'dept' ? 'bg-blue-600 text-white' : 'hover:bg-gray-200'}`}>部門</button>
          <button onClick={() => setActiveMenu('role')} className={`w-full text-left px-3 py-2 rounded mb-4 ${activeMenu === 'role' ? 'bg-blue-600 text-white' : 'hover:bg-gray-200'}`}>役職</button>
          <a href="/" className="block text-center px-3 py-2 rounded border hover:bg-white">ユーザー画面に戻る</a>
        </aside>

        <main className="flex-1 p-6 overflow-auto">
          {error && <div className="mb-3 text-sm text-red-600">{error}</div>}

          {activeMenu === 'org' && (
            <div>
              <div className="flex items-center justify-between mb-3">
                <h1 className="text-xl font-bold">局</h1>
                <button onClick={() => setOrgModal({ open: true, item: null })} className="px-3 py-2 bg-blue-600 text-white rounded">新規作成</button>
              </div>
              <table className="w-full text-sm border">
                <thead className="bg-gray-100"><tr><th className="p-2 text-left">id</th><th className="p-2 text-left">局名</th><th className="p-2 text-left">gen</th><th className="p-2 text-left">状態</th><th className="p-2 text-left">所属区分</th><th className="p-2"></th></tr></thead>
                <tbody>
                  {sortByPid(orgs).map(o => <tr key={o.pid} className="border-t"><td className="p-2">{padOrgIdForDisplay(o.pid)}</td><td className="p-2">{o.org}</td><td className="p-2">{o.gen || ''}</td><td className="p-2">{statusLabel(o.status)}</td><td className="p-2">{o.not_main_org ? '特命局' : '主所属局'}</td><td className="p-2 text-right"><button className="px-2 py-1 border rounded" onClick={() => setOrgModal({ open: true, item: o })}>編集</button></td></tr>)}
                </tbody>
              </table>
            </div>
          )}

          {activeMenu === 'dept' && (
            <div>
              <div className="flex items-center justify-between mb-3">
                <h1 className="text-xl font-bold">部門</h1>
                <button onClick={() => setDeptModal({ open: true, item: null })} className="px-3 py-2 bg-blue-600 text-white rounded">新規作成</button>
              </div>
              <table className="w-full text-sm border">
                <thead className="bg-gray-100"><tr><th className="p-2 text-left">id</th><th className="p-2 text-left">部門名</th><th className="p-2 text-left">局</th><th className="p-2 text-left">状態</th><th className="p-2 text-left">所属区分</th><th className="p-2"></th></tr></thead>
                <tbody>
                  {sortByPid(depts).map(d => <tr key={d.pid} className="border-t"><td className="p-2">{padDeptIdForDisplay(d.pid)}</td><td className="p-2">{d.dept}</td><td className="p-2">{d.org || d.orgPid}</td><td className="p-2">{statusLabel(d.status)}</td><td className="p-2">{d.not_main_dept ? '特命部門' : '主所属部門'}</td><td className="p-2 text-right"><button className="px-2 py-1 border rounded" onClick={() => setDeptModal({ open: true, item: d })}>編集</button></td></tr>)}
                </tbody>
              </table>
            </div>
          )}

          {activeMenu === 'role' && (
            <div>
              <div className="flex items-center justify-between mb-3">
                <h1 className="text-xl font-bold">役職</h1>
                <button onClick={() => setRoleModal({ open: true, item: null })} className="px-3 py-2 bg-blue-600 text-white rounded">新規作成</button>
              </div>
              <table className="w-full text-sm border">
                <thead className="bg-gray-100"><tr><th className="p-2 text-left">id</th><th className="p-2 text-left">役職</th><th className="p-2 text-left">gen</th><th className="p-2 text-left">役職区分</th><th className="p-2"></th></tr></thead>
                <tbody>
                  {sortByPid(roles).map(r => <tr key={r.pid} className="border-t"><td className="p-2">{r.pid}</td><td className="p-2">{r.role}</td><td className="p-2">{r.gen || ''}</td><td className="p-2">{r.not_main_role ? '特命役職' : '主所属役職'}</td><td className="p-2 text-right"><button className="px-2 py-1 border rounded" onClick={() => setRoleModal({ open: true, item: r })}>編集</button></td></tr>)}
                </tbody>
              </table>
            </div>
          )}
        </main>
      </div>

      {orgModal.open && (
        <div className="fixed inset-0 bg-black/40 flex items-center justify-center p-4">
          <form onSubmit={saveOrg} className="relative bg-white rounded-xl p-5 w-full max-w-lg space-y-3">
            {isSaving('org') && (
              <div className="absolute inset-0 z-10 flex items-center justify-center rounded-xl bg-white/80 backdrop-blur-[1px]">
                <div className="flex items-center gap-3 rounded-full bg-gray-900 px-4 py-2 text-sm text-white shadow-lg">
                  <i className="fas fa-spinner fa-spin" />
                  保存中...
                </div>
              </div>
            )}
            <h2 className="text-lg font-bold">局の編集</h2>
            {error && (
              <div className="rounded border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-700" role="alert">
                {error}
              </div>
            )}
            <div className="space-y-1">
              <label className="text-xs text-gray-500">pid</label>
              <input type="hidden" name="pid" value={orgModal.item?.pid || ''} />
              <input value={padOrgIdForDisplay(orgModal.item?.pid || '')} className="w-full border p-2 rounded bg-gray-100" placeholder="id(新規時は自動採番)" disabled />
            </div>
            <div className="space-y-1">
              <label className="text-xs text-gray-500">局名称</label>
              <input name="org" defaultValue={orgModal.item?.org || ''} className="w-full border p-2 rounded" placeholder="局名 *必須" required disabled={isSaving('org')} />
            </div>
            <div className="space-y-1">
              <label className="text-xs text-gray-500">開催回数</label>
              <input
                type="number"
                name="gen"
                value={orgGenValue}
                onChange={(e) => setOrgGenValue(e.target.value)}
                className="w-full border p-2 rounded"
                placeholder="gen (not_main_org の場合必須)"
                required={orgNotMainChecked}
                disabled={!orgNotMainChecked || isSaving('org')}
              />
            </div>
            <div className="space-y-1">
              <label className="text-xs text-gray-500">ステータス</label>
              <select name="status" defaultValue={isActiveStatus(orgModal.item?.status) ? '' : '1'} className="w-full border p-2 rounded" disabled={isSaving('org')}>
                <option value="">有効</option><option value="1">無効</option>
              </select>
            </div>
            <div className="space-y-1">
              <label className="flex items-center gap-2 text-sm">
                <input
                  type="checkbox"
                  name="not_main_org"
                  checked={orgNotMainChecked}
                  onChange={(e) => {
                    const checked = e.target.checked;
                    setOrgNotMainChecked(checked);
                    if (!checked) setOrgGenValue('');
                  }}
                  disabled={isSaving('org')}
                />
                主所属局外
              </label>
            </div>
            <div className="text-xs text-gray-500">note: 主所属局外 (特命局) の場合、`gen` を2桁で入力してください。局 pid は `gen(2桁)+通し番号(1桁)` の3桁になります。</div>
            <div className="flex justify-end gap-2"><button type="button" onClick={() => setOrgModal({ open: false, item: null })} className="px-3 py-2 rounded bg-gray-100" disabled={isSaving('org')}>キャンセル</button><button className="px-3 py-2 rounded bg-blue-600 text-white" disabled={isSaving('org') || (orgNotMainChecked && orgGenValue.trim() === '')}>保存</button></div>
          </form>
        </div>
      )}

      {deptModal.open && (
        <div className="fixed inset-0 bg-black/40 flex items-center justify-center p-4">
          <form onSubmit={saveDept} className="relative bg-white rounded-xl p-5 w-full max-w-lg space-y-3">
            {isSaving('dept') && (
              <div className="absolute inset-0 z-10 flex items-center justify-center rounded-xl bg-white/80 backdrop-blur-[1px]">
                <div className="flex items-center gap-3 rounded-full bg-gray-900 px-4 py-2 text-sm text-white shadow-lg">
                  <i className="fas fa-spinner fa-spin" />
                  保存中...
                </div>
              </div>
            )}
            <h2 className="text-lg font-bold">部門の編集</h2>
            {error && (
              <div className="rounded border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-700" role="alert">
                {error}
              </div>
            )}
            <div className="space-y-1">
              <label className="text-xs text-gray-500">pid</label>
              <input type="hidden" name="pid" value={deptModal.item?.pid || ''} />
              <input value={padDeptIdForDisplay(deptModal.item?.pid || '')} className="w-full border p-2 rounded bg-gray-100" placeholder="id(新規時は自動採番)" disabled />
            </div>
            <div className="space-y-1">
              <label className="text-xs text-gray-500">部門名称</label>
              <input name="dept" defaultValue={deptModal.item?.dept || ''} className="w-full border p-2 rounded" placeholder="部門名 *必須" required disabled={isSaving('dept')} />
            </div>
            <div className="space-y-1">
              <label className="text-xs text-gray-500">所属局</label>
              <select
                name="orgPid"
                value={deptOrgPid}
                onChange={(e) => {
                  const nextOrgPid = e.target.value;
                  const selectedOrg = orgs.find((o) => String(o.pid) === nextOrgPid);
                  const forceNotMain = !!selectedOrg?.not_main_org;
                  setDeptOrgPid(nextOrgPid);
                  if (forceNotMain) {
                    setDeptNotMainChecked(true);
                  }
                }}
                className="w-full border p-2 rounded"
                required
                disabled={isSaving('dept')}
              >
                <option value="">所属局を選択 *必須</option>
                {sortByPid(orgs).map(o => <option key={o.pid} value={o.pid}>{o.org} ({padOrgIdForDisplay(o.pid)})</option>)}
              </select>
            </div>
            <div className="space-y-1">
              <label className="text-xs text-gray-500">ステータス</label>
              <select name="status" defaultValue={isActiveStatus(deptModal.item?.status) ? '' : '1'} className="w-full border p-2 rounded" disabled={isSaving('dept')}>
                <option value="">有効</option><option value="1">無効</option>
              </select>
            </div>
            <div className="space-y-1">
              <label className="flex items-center gap-2 text-sm"><input type="checkbox" name="not_main_dept" checked={deptNotMainChecked} onChange={(e) => setDeptNotMainChecked(e.target.checked)} disabled={isSaving('dept') || !!orgs.find((o) => String(o.pid) === deptOrgPid)?.not_main_org} /> 主所属部門外</label>
            </div>
            <div className="flex justify-end gap-2"><button type="button" onClick={() => setDeptModal({ open: false, item: null })} className="px-3 py-2 rounded bg-gray-100" disabled={isSaving('dept')}>キャンセル</button><button className="px-3 py-2 rounded bg-blue-600 text-white" disabled={isSaving('dept')}>保存</button></div>
          </form>
        </div>
      )}

      {roleModal.open && (
        <div className="fixed inset-0 bg-black/40 flex items-center justify-center p-4">
          <form onSubmit={saveRole} className="relative bg-white rounded-xl p-5 w-full max-w-lg space-y-3">
            {isSaving('role') && (
              <div className="absolute inset-0 z-10 flex items-center justify-center rounded-xl bg-white/80 backdrop-blur-[1px]">
                <div className="flex items-center gap-3 rounded-full bg-gray-900 px-4 py-2 text-sm text-white shadow-lg">
                  <i className="fas fa-spinner fa-spin" />
                  保存中...
                </div>
              </div>
            )}
            <h2 className="text-lg font-bold">役職の編集</h2>
            {error && (
              <div className="rounded border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-700" role="alert">
                {error}
              </div>
            )}
            <div className="space-y-1">
              <label className="text-xs text-gray-500">pid</label>
              <input type="hidden" name="pid" value={roleModal.item?.pid || ''} />
              <input value={roleModal.item?.pid || ''} className="w-full border p-2 rounded bg-gray-100" placeholder="id(新規時は自動採番)" disabled />
            </div>
            <div className="space-y-1">
              <label className="text-xs text-gray-500">役職名称</label>
              <input name="role" defaultValue={roleModal.item?.role || ''} className="w-full border p-2 rounded" placeholder="役職名 *必須" required disabled={isSaving('role')} />
            </div>
            <div className="space-y-1">
              <label className="text-xs text-gray-500">開催回数</label>
              <input type="number" name="gen" defaultValue={String(roleModal.item?.gen || '')} className="w-full border p-2 rounded" placeholder="gen *必須" required disabled={isSaving('role')} />
            </div>
            <div className="space-y-1">
              <label className="flex items-center gap-2 text-sm"><input type="checkbox" name="not_main_role" defaultChecked={!!roleModal.item?.not_main_role} disabled={isSaving('role')} /> 主所属役職外</label>
            </div>
            <div className="flex justify-end gap-2"><button type="button" onClick={() => setRoleModal({ open: false, item: null })} className="px-3 py-2 rounded bg-gray-100" disabled={isSaving('role')}>キャンセル</button><button className="px-3 py-2 rounded bg-blue-600 text-white" disabled={isSaving('role')}>保存</button></div>
          </form>
        </div>
      )}
    </div>
  );
}
