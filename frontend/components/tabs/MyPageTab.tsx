'use client';

import { useEffect, useState } from 'react';
import { gasApi } from '../../lib/api';
import type { UserProfile } from '../../lib/types';

type Props = { sessionToken: string };

export function MyPageTab({ sessionToken }: Props) {
  const [profile, setProfile] = useState<UserProfile | null>(null);
  const [saving, setSaving] = useState(false);
  const [status, setStatus] = useState('');

  useEffect(() => {
    gasApi
      .getUserProfile(sessionToken)
      .then((res) => setProfile(res as UserProfile))
      .catch((e) => setStatus(e instanceof Error ? e.message : String(e)));
  }, [sessionToken]);

  const save = async () => {
    if (!profile) return;
    setSaving(true);
    setStatus('');
    try {
      await gasApi.updateUserProfile(sessionToken, {
        name: profile.name,
        nameEn: profile.nameEn,
        studentId: profile.studentId,
        grade: profile.grade,
        field: profile.field,
        phone: profile.phone,
        birthday: profile.birthday,
        almaMater: profile.almaMater,
        carOwner: profile.carOwner,
        retired: profile.retired,
        continueNext: profile.continueNext,
        isAdmin: profile.isAdmin,
        orgs: []
      });
      setStatus('保存しました');
    } catch (e) {
      setStatus(e instanceof Error ? e.message : String(e));
    } finally {
      setSaving(false);
    }
  };

  if (!profile) return <p>読み込み中...</p>;

  return (
    <div className="card" style={{ padding: 16 }}>
      <h3 style={{ marginTop: 0 }}>プロフィール</h3>
      <div style={{ display: 'grid', gap: 10 }}>
        <label>
          <div>氏名</div>
          <input style={input} value={profile.name || ''} onChange={(e) => setProfile({ ...profile, name: e.target.value })} />
        </label>
        <label>
          <div>英語名</div>
          <input style={input} value={profile.nameEn || ''} onChange={(e) => setProfile({ ...profile, nameEn: e.target.value })} />
        </label>
        <label>
          <div>電話番号</div>
          <input style={input} value={profile.phone || ''} onChange={(e) => setProfile({ ...profile, phone: e.target.value })} />
        </label>
      </div>
      <div style={{ marginTop: 14 }}>
        <button style={primaryBtn} onClick={save} disabled={saving}>保存</button>
      </div>
      {status ? <p>{status}</p> : null}
    </div>
  );
}

const input: any = {
  width: '100%',
  border: '1px solid #d1d5db',
  borderRadius: 8,
  padding: '10px 12px',
  marginTop: 6
};

const primaryBtn: any = {
  border: 'none',
  background: '#1a237e',
  color: '#fff',
  borderRadius: 8,
  padding: '10px 14px',
  cursor: 'pointer'
};
