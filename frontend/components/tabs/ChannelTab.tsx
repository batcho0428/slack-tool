'use client';

import { useEffect, useState } from 'react';
import { gasApi } from '../../lib/api';
import type { Channel, Recipient } from '../../lib/types';

type Props = { sessionToken: string };

export function ChannelTab({ sessionToken }: Props) {
  const [channels, setChannels] = useState<Channel[]>([]);
  const [channelId, setChannelId] = useState('');
  const [query, setQuery] = useState('');
  const [results, setResults] = useState<Recipient[]>([]);
  const [recipients, setRecipients] = useState<Recipient[]>([]);
  const [status, setStatus] = useState('');

  useEffect(() => {
    gasApi
      .getChannels(sessionToken)
      .then((list) => setChannels((list as Channel[]) || []))
      .catch((e) => setStatus(e instanceof Error ? e.message : String(e)));
  }, [sessionToken]);

  const search = async () => {
    try {
      const list = (await gasApi.searchRecipients({ query, status: 'active' })) as Recipient[];
      setResults(list || []);
    } catch (e) {
      setStatus(e instanceof Error ? e.message : String(e));
    }
  };

  const invite = async () => {
    try {
      const res = (await gasApi.inviteToChannel(sessionToken, channelId, recipients)) as {
        success: number;
        failed: Array<{ email: string; error: string }>;
      };
      setStatus(`招待完了: 成功 ${res.success} 件 / 失敗 ${(res.failed || []).length} 件`);
    } catch (e) {
      setStatus(e instanceof Error ? e.message : String(e));
    }
  };

  return (
    <div style={{ display: 'grid', gap: 16 }}>
      <section className="card" style={{ padding: 16 }}>
        <h3 style={{ marginTop: 0 }}>チャンネル招待</h3>
        <select style={input} value={channelId} onChange={(e) => setChannelId(e.target.value)}>
          <option value="">チャンネルを選択</option>
          {channels.map((c) => (
            <option key={c.id} value={c.id}>{c.name}</option>
          ))}
        </select>
      </section>

      <section className="card" style={{ padding: 16 }}>
        <div className="row" style={{ alignItems: 'stretch' }}>
          <input style={{ ...input, flex: 1 }} value={query} onChange={(e) => setQuery(e.target.value)} placeholder="ユーザー検索" />
          <button style={btn} onClick={search}>検索</button>
        </div>
        <div style={{ marginTop: 10, display: 'grid', gap: 6 }}>
          {results.slice(0, 20).map((r) => (
            <button key={r.email} onClick={() => setRecipients((prev) => prev.some((x) => x.email === r.email) ? prev : [...prev, r])} style={resultBtn}>
              {r.name} ({r.email})
            </button>
          ))}
        </div>
      </section>

      <section className="card" style={{ padding: 16 }}>
        <h4 style={{ marginTop: 0 }}>対象者 ({recipients.length})</h4>
        <ul style={{ margin: 0, paddingInlineStart: 20 }}>
          {recipients.map((r) => <li key={r.email}>{r.name} ({r.email})</li>)}
        </ul>
      </section>

      <button style={primaryBtn} onClick={invite} disabled={!channelId || recipients.length === 0}>招待実行</button>
      {status ? <p style={{ margin: 0 }}>{status}</p> : null}
    </div>
  );
}

const input: any = {
  border: '1px solid #d1d5db',
  borderRadius: 8,
  padding: '10px 12px',
  width: '100%'
};

const btn: any = {
  border: '1px solid #d1d5db',
  background: '#fff',
  borderRadius: 8,
  padding: '10px 12px',
  cursor: 'pointer'
};

const primaryBtn: any = {
  border: 'none',
  background: '#14532d',
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
