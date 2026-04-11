'use client';

import { useCallback, useEffect, useMemo, useState, type CSSProperties, type ReactNode } from 'react';
import { gasApi } from '../../lib/api';
import type { CollectionItem, CollectionSummary } from '../../lib/types';

type Props = { sessionToken: string };

export function CollectionsTab({ sessionToken }: Props) {
  const [collections, setCollections] = useState<CollectionItem[]>([]);
  const [selectedId, setSelectedId] = useState('');
  const [summary, setSummary] = useState<CollectionSummary | null>(null);
  const [loading, setLoading] = useState(false);
  const [status, setStatus] = useState('');

  const [title, setTitle] = useState('');
  const [spreadsheetUrl, setSpreadsheetUrl] = useState('');
  const [inChargeOrg, setInChargeOrg] = useState('');
  const [inChargeDept, setInChargeDept] = useState('');

  const [recipientEmail, setRecipientEmail] = useState('');
  const [amount, setAmount] = useState('');
  const [handlerEmail, setHandlerEmail] = useState('');

  const selected = useMemo(
    () => collections.find((c) => c.id === selectedId) || null,
    [collections, selectedId]
  );

  const loadCollections = useCallback(async () => {
    setLoading(true);
    setStatus('');
    try {
      const list = (await gasApi.listCollections(sessionToken)) as CollectionItem[];
      setCollections(Array.isArray(list) ? list : []);
      if (!selectedId && Array.isArray(list) && list[0]?.id) setSelectedId(list[0].id);
    } catch (e) {
      setStatus(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  }, [selectedId, sessionToken]);

  const loadSummary = useCallback(async (collectionId: string) => {
    if (!collectionId) return;
    setStatus('');
    setSummary(null);
    try {
      const s = (await gasApi.fetchCollectionSummary(sessionToken, collectionId)) as CollectionSummary;
      setSummary(s);
      if (!s.success) setStatus(s.message || '集金サマリの取得に失敗しました。');
    } catch (e) {
      setStatus(e instanceof Error ? e.message : String(e));
    }
  }, [sessionToken]);

  useEffect(() => {
    loadCollections();
  }, [loadCollections]);

  useEffect(() => {
    if (selectedId) loadSummary(selectedId);
  }, [loadSummary, selectedId]);

  const create = async () => {
    setStatus('');
    try {
      const res = (await gasApi.createCollection(sessionToken, {
        title,
        spreadsheetUrl,
        inChargeOrg,
        inChargeDept
      })) as { success: boolean; message?: string };
      if (!res.success) throw new Error(res.message || '作成に失敗しました。');
      setTitle('');
      setSpreadsheetUrl('');
      await loadCollections();
    } catch (e) {
      setStatus(e instanceof Error ? e.message : String(e));
    }
  };

  const record = async (useChange: boolean) => {
    if (!selectedId) return;
    const numericAmount = Number(amount);
    if (!Number.isFinite(numericAmount)) {
      setStatus('金額は数値で入力してください。');
      return;
    }
    setStatus('');
    try {
      const res = useChange
        ? ((await gasApi.recordPaymentWithChange(
            sessionToken,
            selectedId,
            recipientEmail,
            numericAmount,
            summary?.perPerson?.find((p) => p.email === recipientEmail)?.expected || 0,
            handlerEmail
          )) as { success: boolean; message?: string })
        : ((await gasApi.recordPayment(
            sessionToken,
            selectedId,
            recipientEmail,
            numericAmount,
            '受領',
            handlerEmail
          )) as { success: boolean; message?: string });
      if (!res.success) throw new Error(res.message || '記録に失敗しました。');
      setAmount('');
      await loadSummary(selectedId);
    } catch (e) {
      setStatus(e instanceof Error ? e.message : String(e));
    }
  };

  return (
    <div style={{ display: 'grid', gap: 16 }}>
      <section className="card" style={{ padding: 16 }}>
        <h3 style={{ marginTop: 0 }}>集金マスタ</h3>
        <div className="row" style={{ flexWrap: 'wrap' }}>
          <input style={input} value={title} onChange={(e) => setTitle(e.target.value)} placeholder="タイトル" />
          <input
            style={{ ...input, minWidth: 280 }}
            value={spreadsheetUrl}
            onChange={(e) => setSpreadsheetUrl(e.target.value)}
            placeholder="スプレッドシートURL"
          />
          <input style={input} value={inChargeOrg} onChange={(e) => setInChargeOrg(e.target.value)} placeholder="担当局" />
          <input style={input} value={inChargeDept} onChange={(e) => setInChargeDept(e.target.value)} placeholder="担当部門" />
          <button onClick={create} style={primaryBtn} disabled={!title || !spreadsheetUrl}>
            新規作成
          </button>
        </div>
      </section>

      <section className="card" style={{ padding: 16 }}>
        <div className="row" style={{ justifyContent: 'space-between' }}>
          <h4 style={{ margin: 0 }}>対象集金</h4>
          <button onClick={loadCollections} style={btn} disabled={loading}>
            再読み込み
          </button>
        </div>
        <select style={{ ...input, width: '100%', marginTop: 10 }} value={selectedId} onChange={(e) => setSelectedId(e.target.value)}>
          <option value="">選択してください</option>
          {collections.map((c) => (
            <option key={c.id} value={c.id}>
              {c.title} ({c.inChargeOrg}/{c.inChargeDept})
            </option>
          ))}
        </select>
        {selected?.spreadsheetUrl ? (
          <p style={{ marginBottom: 0 }}>
            <a href={selected.spreadsheetUrl} target="_blank" rel="noreferrer">
              元データを開く
            </a>
          </p>
        ) : null}
      </section>

      {summary?.success ? (
        <section className="card" style={{ padding: 16, display: 'grid', gap: 12 }}>
          <h4 style={{ margin: 0 }}>集金サマリ</h4>
          <div className="row" style={{ flexWrap: 'wrap' }}>
            <span className="badge">請求合計 {Number(summary.expectedTotal || 0).toLocaleString()} 円</span>
            <span className="badge">受領合計 {Number(summary.collectedTotal || 0).toLocaleString()} 円</span>
            <span className="badge">対象人数 {Number(summary.expectedCount || 0)} 名</span>
            <span className="badge">受領人数 {Number(summary.collectedCount || 0)} 名</span>
          </div>

          <div className="row" style={{ flexWrap: 'wrap' }}>
            <input
              style={{ ...input, minWidth: 260 }}
              value={recipientEmail}
              onChange={(e) => setRecipientEmail(e.target.value)}
              placeholder="対象メールアドレス"
              list="collection-emails"
            />
            <datalist id="collection-emails">
              {(summary.perPerson || []).map((p) => (
                <option key={p.email} value={p.email} />
              ))}
            </datalist>
            <input style={input} value={amount} onChange={(e) => setAmount(e.target.value)} placeholder="受領金額" />
            <input
              style={{ ...input, minWidth: 240 }}
              value={handlerEmail}
              onChange={(e) => setHandlerEmail(e.target.value)}
              placeholder="担当者メール"
            />
            <button style={btn} onClick={() => record(false)} disabled={!recipientEmail || !amount}>
              受領記録
            </button>
            <button style={primaryBtn} onClick={() => record(true)} disabled={!recipientEmail || !amount}>
              おつり込み記録
            </button>
          </div>

          <div style={{ maxHeight: 320, overflow: 'auto', border: '1px solid #e5e7eb', borderRadius: 10 }}>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr>
                  <Th>メール</Th>
                  <Th>請求額</Th>
                  <Th>受領額</Th>
                  <Th>状態</Th>
                </tr>
              </thead>
              <tbody>
                {(summary.perPerson || []).map((p) => (
                  <tr key={p.email}>
                    <Td>{p.email}</Td>
                    <Td>{Number(p.expected || 0).toLocaleString()}</Td>
                    <Td>{Number(p.collected || 0).toLocaleString()}</Td>
                    <Td>{p.status}</Td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      ) : null}

      {status ? <p style={{ margin: 0, color: '#b91c1c' }}>{status}</p> : null}
    </div>
  );
}

function Th({ children }: { children: ReactNode }) {
  return (
    <th
      style={{
        textAlign: 'left',
        position: 'sticky',
        top: 0,
        background: '#f9fafb',
        borderBottom: '1px solid #e5e7eb',
        padding: '8px 10px'
      }}
    >
      {children}
    </th>
  );
}

function Td({ children }: { children: ReactNode }) {
  return <td style={{ padding: '8px 10px', borderBottom: '1px solid #f3f4f6' }}>{children}</td>;
}

const input: CSSProperties = {
  border: '1px solid #d1d5db',
  borderRadius: 8,
  padding: '10px 12px'
};

const btn: CSSProperties = {
  border: '1px solid #d1d5db',
  background: '#fff',
  borderRadius: 8,
  padding: '8px 10px',
  cursor: 'pointer'
};

const primaryBtn: CSSProperties = {
  border: 'none',
  background: '#14532d',
  color: '#fff',
  borderRadius: 8,
  padding: '8px 10px',
  cursor: 'pointer'
};
