'use client';

import { useMemo, useState } from 'react';
import { gasApi } from '../../lib/api';
import type { Recipient } from '../../lib/types';

type Props = { sessionToken: string };

export function DmTab({ sessionToken }: Props) {
  const [message, setMessage] = useState('');
  const [query, setQuery] = useState('');
  const [searching, setSearching] = useState(false);
  const [results, setResults] = useState<Recipient[]>([]);
  const [recipients, setRecipients] = useState<Recipient[]>([]);
  const [sending, setSending] = useState(false);
  const [status, setStatus] = useState('');

  const recipientIds = useMemo(() => new Set(recipients.map((r) => r.email)), [recipients]);

  const search = async () => {
    setSearching(true);
    setStatus('');
    try {
      const list = (await gasApi.searchRecipients({ query, status: 'active' })) as Recipient[];
      setResults(list || []);
    } catch (e) {
      setStatus(e instanceof Error ? e.message : String(e));
    } finally {
      setSearching(false);
    }
  };

  const addRecipient = (item: Recipient) => {
    if (recipientIds.has(item.email)) return;
    setRecipients((prev) => [...prev, item]);
  };

  const removeRecipient = (email: string) => {
    setRecipients((prev) => prev.filter((r) => r.email !== email));
  };

  const send = async () => {
    setSending(true);
    setStatus('');
    try {
      const res = (await gasApi.sendDMs(sessionToken, message, recipients)) as {
        success: number;
        failed: Array<{ email: string; error: string }>;
      };
      setStatus(`送信完了: 成功 ${res.success} 件 / 失敗 ${(res.failed || []).length} 件`);
    } catch (e) {
      setStatus(e instanceof Error ? e.message : String(e));
    } finally {
      setSending(false);
    }
  };

  return (
    <div style={{ display: 'grid', gap: 16 }}>
      <section className="card" style={{ padding: 16 }}>
        <h3 style={{ marginTop: 0 }}>一斉DM</h3>
        <textarea
          style={{ width: '100%', minHeight: 120, border: '1px solid #d1d5db', borderRadius: 8, padding: 10 }}
          value={message}
          onChange={(e) => setMessage(e.target.value)}
          placeholder="送信するメッセージ"
        />
      </section>

      <section className="card" style={{ padding: 16 }}>
        <h4 style={{ marginTop: 0 }}>送信先検索</h4>
        <div className="row" style={{ alignItems: 'stretch' }}>
          <input
            style={{ flex: 1, border: '1px solid #d1d5db', borderRadius: 8, padding: '10px 12px' }}
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="名前・メールで検索"
          />
          <button onClick={search} disabled={searching} style={btn}>検索</button>
        </div>
        <div style={{ marginTop: 10, display: 'grid', gap: 6 }}>
          {results.slice(0, 20).map((r) => (
            <button key={r.email} onClick={() => addRecipient(r)} style={resultBtn}>
              {r.name} ({r.email})
            </button>
          ))}
        </div>
      </section>

      <section className="card" style={{ padding: 16 }}>
        <h4 style={{ marginTop: 0 }}>選択済み ({recipients.length})</h4>
        <div style={{ display: 'grid', gap: 6 }}>
          {recipients.map((r) => (
            <div key={r.email} style={recipientRow}>
              <span>{r.name} ({r.email})</span>
              <button onClick={() => removeRecipient(r.email)} style={removeBtn}>削除</button>
            </div>
          ))}
        </div>
      </section>

      <div>
        <button onClick={send} disabled={sending || !message || recipients.length === 0} style={primaryBtn}>送信</button>
      </div>
      {status ? <p style={{ margin: 0, color: '#374151' }}>{status}</p> : null}
    </div>
  );
}

const btn: any = {
  border: '1px solid #d1d5db',
  background: '#fff',
  borderRadius: 8,
  padding: '10px 12px',
  cursor: 'pointer'
};

const primaryBtn: any = {
  border: 'none',
  background: '#1a237e',
  color: '#fff',
  borderRadius: 8,
  padding: '10px 14px',
  cursor: 'pointer'
};

const resultBtn: any = {
  textAlign: 'left',
  border: '1px solid #e5e7eb',
  background: '#fff',
  borderRadius: 8,
  padding: '8px 10px',
  cursor: 'pointer'
};

const recipientRow: any = {
  display: 'flex',
  justifyContent: 'space-between',
  alignItems: 'center',
  border: '1px solid #e5e7eb',
  borderRadius: 8,
  padding: '8px 10px'
};

const removeBtn: any = {
  border: '1px solid #fecaca',
  color: '#b91c1c',
  background: '#fff5f5',
  borderRadius: 8,
  padding: '4px 8px',
  cursor: 'pointer'
};
