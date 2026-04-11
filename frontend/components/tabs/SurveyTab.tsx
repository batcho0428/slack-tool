'use client';

import { useCallback, useEffect, useState, type CSSProperties } from 'react';
import { gasApi } from '../../lib/api';
import type { SurveyDetailResponse, SurveyItem } from '../../lib/types';

type Props = { sessionToken: string };

export function SurveyTab({ sessionToken }: Props) {
  const [surveys, setSurveys] = useState<SurveyItem[]>([]);
  const [loading, setLoading] = useState(false);
  const [selected, setSelected] = useState<SurveyItem | null>(null);
  const [details, setDetails] = useState<SurveyDetailResponse | null>(null);
  const [status, setStatus] = useState('');

  const load = useCallback(async () => {
    setLoading(true);
    setStatus('');
    try {
      const res = await gasApi.listSurveys(sessionToken);
      if (Array.isArray(res)) {
        setSurveys(res as SurveyItem[]);
      } else {
        const obj = res as { success?: boolean; message?: string };
        if (obj.success === false) {
          setStatus(obj.message || 'アンケート一覧の取得に失敗しました。');
          setSurveys([]);
        } else {
          setSurveys([]);
        }
      }
    } catch (e) {
      setStatus(e instanceof Error ? e.message : String(e));
      setSurveys([]);
    } finally {
      setLoading(false);
    }
  }, [sessionToken]);

  useEffect(() => {
    load();
  }, [load]);

  const openDetails = async (item: SurveyItem) => {
    if (!item.spreadsheetUrl || !item.userLatestRowIndex) {
      setStatus('このアンケートには表示可能な回答がありません。');
      return;
    }
    setStatus('');
    setSelected(item);
    setDetails(null);
    try {
      const res = (await gasApi.getSurveyDetails(
        sessionToken,
        item.spreadsheetUrl,
        item.userLatestRowIndex
      )) as SurveyDetailResponse;
      setDetails(res);
      if (!res.success) setStatus(res.message || '回答詳細の取得に失敗しました。');
    } catch (e) {
      setStatus(e instanceof Error ? e.message : String(e));
    }
  };

  return (
    <div style={{ display: 'grid', gap: 16 }}>
      <section className="card" style={{ padding: 16 }}>
        <div className="row" style={{ justifyContent: 'space-between' }}>
          <h3 style={{ margin: 0 }}>アンケート一覧</h3>
          <button onClick={load} style={btn} disabled={loading}>
            再読み込み
          </button>
        </div>

        {loading ? <p>読み込み中...</p> : null}
        <div style={{ display: 'grid', gap: 8 }}>
          {surveys.map((s) => (
            <div key={`${s.title}-${s.spreadsheetId || 'na'}`} style={itemBox}>
              <div style={{ display: 'grid', gap: 4 }}>
                <strong>{s.title}</strong>
                <div style={{ color: '#4b5563', fontSize: 13 }}>
                  担当: {s.inChargeOrg || '-'} / {s.inChargeDept || '-'}
                </div>
                <div style={{ color: '#4b5563', fontSize: 13 }}>
                  最新スコア: {s.latestScoreFormatted || '-'}
                </div>
              </div>
              <div style={{ display: 'flex', gap: 8 }}>
                {s.formUrl ? (
                  <a href={s.formUrl} target="_blank" rel="noreferrer" style={linkBtn}>
                    フォーム
                  </a>
                ) : null}
                <button onClick={() => openDetails(s)} style={primaryBtn} disabled={!s.available}>
                  回答を表示
                </button>
              </div>
            </div>
          ))}
          {!loading && surveys.length === 0 ? <p style={{ marginBottom: 0 }}>表示できるアンケートがありません。</p> : null}
        </div>
      </section>

      {selected ? (
        <section className="card" style={{ padding: 16 }}>
          <h4 style={{ marginTop: 0 }}>{selected.title} 回答詳細</h4>
          {!details ? <p>読み込み中...</p> : null}
          {details?.success && details.response ? (
            <div style={{ display: 'grid', gap: 8 }}>
              <div style={{ color: '#374151' }}>
                {details.scoreName || 'スコア'}: {details.response.scoreFormatted || details.response.score || '-'}
              </div>
              <div style={detailGrid}>
                {Object.entries(details.response.answers || {}).map(([k, v]) => (
                  <div key={k} style={detailRow}>
                    <div style={{ color: '#374151', fontWeight: 700 }}>{k}</div>
                    <div>{String(v ?? '')}</div>
                  </div>
                ))}
              </div>
            </div>
          ) : null}
        </section>
      ) : null}

      {status ? <p style={{ margin: 0, color: '#b91c1c' }}>{status}</p> : null}
    </div>
  );
}

const itemBox: CSSProperties = {
  border: '1px solid #e5e7eb',
  borderRadius: 10,
  padding: 12,
  display: 'flex',
  justifyContent: 'space-between',
  gap: 12,
  alignItems: 'center'
};

const detailGrid: CSSProperties = {
  display: 'grid',
  gap: 8,
  borderTop: '1px solid #e5e7eb',
  paddingTop: 8
};

const detailRow: CSSProperties = {
  border: '1px solid #e5e7eb',
  borderRadius: 8,
  padding: '8px 10px'
};

const btn: CSSProperties = {
  border: '1px solid #d1d5db',
  background: '#fff',
  borderRadius: 8,
  padding: '8px 10px',
  cursor: 'pointer'
};

const linkBtn: CSSProperties = {
  ...btn,
  textDecoration: 'none',
  color: '#374151',
  display: 'inline-flex',
  alignItems: 'center'
};

const primaryBtn: CSSProperties = {
  border: 'none',
  background: '#1a237e',
  color: '#fff',
  borderRadius: 8,
  padding: '8px 10px',
  cursor: 'pointer'
};
