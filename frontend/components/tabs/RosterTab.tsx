'use client';

import { useEffect, useMemo, useState, type CSSProperties } from 'react';
import { gasApi } from '../../lib/api';
import type { CollectionItem, RosterCsvResponse, SurveyItem } from '../../lib/types';

type Props = { sessionToken: string };

type SearchOptions = {
  grades: string[];
  fields: string[];
};

const DEFAULT_FIELDS = ['氏名', 'Name', '学年', '分野', 'メールアドレス', '在籍', '次年度継続'];
const OPTIONAL_FIELDS = [
  '学籍番号',
  '電話番号',
  '生年月日',
  '出身校',
  '所属局1',
  '所属部門1',
  '役職1',
  '所属局2',
  '所属部門2',
  '役職2',
  '所属局3',
  '所属部門3',
  '役職3',
  '所属局4',
  '所属部門4',
  '役職4',
  '所属局5',
  '所属部門5',
  '役職5',
  '車所有',
  'Admin'
];

export function RosterTab({ sessionToken }: Props) {
  const [options, setOptions] = useState<SearchOptions>({ grades: [], fields: [] });
  const [selectedFields, setSelectedFields] = useState<string[]>(DEFAULT_FIELDS);
  const [statusFilter, setStatusFilter] = useState<'active' | 'retired' | 'all'>('active');
  const [grade, setGrade] = useState('');
  const [field, setField] = useState('');
  const [surveyRef, setSurveyRef] = useState('');
  const [collectionId, setCollectionId] = useState('');
  const [filename, setFilename] = useState('');
  const [status, setStatus] = useState('');
  const [working, setWorking] = useState(false);

  const [surveys, setSurveys] = useState<SurveyItem[]>([]);
  const [collections, setCollections] = useState<CollectionItem[]>([]);

  const allFields = useMemo(() => [...DEFAULT_FIELDS, ...OPTIONAL_FIELDS], []);

  useEffect(() => {
    gasApi
      .getSearchOptions()
      .then((o) => {
        const r = (o || {}) as SearchOptions;
        setOptions({ grades: r.grades || [], fields: r.fields || [] });
      })
      .catch((e) => setStatus(e instanceof Error ? e.message : String(e)));

    gasApi
      .listSurveys(sessionToken)
      .then((list) => {
        if (Array.isArray(list)) setSurveys(list as SurveyItem[]);
      })
      .catch(() => {
        // optional feature; ignore load failure
      });

    gasApi
      .listCollections(sessionToken)
      .then((list) => {
        if (Array.isArray(list)) setCollections(list as CollectionItem[]);
      })
      .catch(() => {
        // optional feature; ignore load failure
      });
  }, [sessionToken]);

  const toggleField = (f: string) => {
    setSelectedFields((prev) =>
      prev.includes(f) ? prev.filter((x) => x !== f) : [...prev, f]
    );
  };

  const download = async (excel: boolean) => {
    setWorking(true);
    setStatus('');
    try {
      const params: Record<string, unknown> = {
        selectedFields,
        filter: {
          type: 'all',
          status: statusFilter,
          grade: grade || null,
          field: field || null
        }
      };
      if (surveyRef) params.surveyRef = surveyRef;
      if (collectionId) params.collectionId = collectionId;
      if (filename) params.filename = filename;

      const res = (await gasApi.createRosterCsv(sessionToken, params)) as RosterCsvResponse;
      if (!res.success) throw new Error(res.message || '名簿出力に失敗しました。');

      const name = res.filename || 'list.csv';
      const buffer = excel ? decodeBase64(res.excelBase64 || '') : decodeBase64(res.csvBase64 || '');
      if (!buffer) {
        // fallback when base64 is not available
        const text = res.csv || '';
        const blob = new Blob([text], { type: 'text/csv;charset=utf-8' });
        triggerDownload(blob, name);
        return;
      }
      const blob = new Blob([buffer], {
        type: excel
          ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
          : 'text/csv;charset=utf-8'
      });
      triggerDownload(blob, excel ? name.replace(/\.csv$/i, '.xlsx') : name);
    } catch (e) {
      setStatus(e instanceof Error ? e.message : String(e));
    } finally {
      setWorking(false);
    }
  };

  return (
    <div style={{ display: 'grid', gap: 16 }}>
      <section className="card" style={{ padding: 16 }}>
        <h3 style={{ marginTop: 0 }}>名簿出力</h3>

        <div style={{ display: 'grid', gap: 10 }}>
          <div className="row" style={{ flexWrap: 'wrap' }}>
            <label>
              在籍区分
              <select style={input} value={statusFilter} onChange={(e) => setStatusFilter(e.target.value as 'active' | 'retired' | 'all')}>
                <option value="active">在籍のみ</option>
                <option value="retired">退局のみ</option>
                <option value="all">全員</option>
              </select>
            </label>
            <label>
              学年
              <select style={input} value={grade} onChange={(e) => setGrade(e.target.value)}>
                <option value="">指定なし</option>
                {options.grades.map((g) => (
                  <option key={g} value={g}>
                    {g}
                  </option>
                ))}
              </select>
            </label>
            <label>
              分野
              <select style={input} value={field} onChange={(e) => setField(e.target.value)}>
                <option value="">指定なし</option>
                {options.fields.map((f) => (
                  <option key={f} value={f}>
                    {f}
                  </option>
                ))}
              </select>
            </label>
            <label>
              ファイル名
              <input
                style={input}
                value={filename}
                onChange={(e) => setFilename(e.target.value)}
                placeholder="list_YYYYMMDD.csv"
              />
            </label>
          </div>

          <div className="row" style={{ flexWrap: 'wrap' }}>
            <label style={{ minWidth: 320 }}>
              アンケート追加列
              <select style={{ ...input, width: '100%' }} value={surveyRef} onChange={(e) => setSurveyRef(e.target.value)}>
                <option value="">なし</option>
                {surveys.map((s) => (
                  <option key={`${s.title}-${s.spreadsheetUrl || 'na'}`} value={s.spreadsheetUrl || ''}>
                    {s.title}
                  </option>
                ))}
              </select>
            </label>
            <label style={{ minWidth: 320 }}>
              集金追加列
              <select style={{ ...input, width: '100%' }} value={collectionId} onChange={(e) => setCollectionId(e.target.value)}>
                <option value="">なし</option>
                {collections.map((c) => (
                  <option key={c.id} value={c.id}>
                    {c.title}
                  </option>
                ))}
              </select>
            </label>
          </div>

          <div>
            <div style={{ marginBottom: 8, fontWeight: 700 }}>出力項目</div>
            <div style={fieldGrid}>
              {allFields.map((f) => (
                <label key={f} style={fieldLabel}>
                  <input type="checkbox" checked={selectedFields.includes(f)} onChange={() => toggleField(f)} />
                  <span>{f}</span>
                </label>
              ))}
            </div>
          </div>

          <div className="row">
            <button style={primaryBtn} onClick={() => download(false)} disabled={working || selectedFields.length === 0}>
              CSVダウンロード
            </button>
            <button style={btn} onClick={() => download(true)} disabled={working || selectedFields.length === 0}>
              Excelダウンロード
            </button>
          </div>
        </div>
      </section>

      {status ? <p style={{ margin: 0, color: '#b91c1c' }}>{status}</p> : null}
    </div>
  );
}

function decodeBase64(b64: string): ArrayBuffer | null {
  if (!b64) return null;
  try {
    const binary = atob(b64);
    const buffer = new ArrayBuffer(binary.length);
    const bytes = new Uint8Array(buffer);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    return buffer;
  } catch {
    return null;
  }
}

function triggerDownload(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

const input: CSSProperties = {
  border: '1px solid #d1d5db',
  borderRadius: 8,
  padding: '8px 10px',
  marginLeft: 8,
  minWidth: 160
};

const fieldGrid: CSSProperties = {
  display: 'grid',
  gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))',
  gap: 8,
  border: '1px solid #e5e7eb',
  borderRadius: 10,
  padding: 10
};

const fieldLabel: CSSProperties = {
  display: 'flex',
  gap: 8,
  alignItems: 'center'
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
  background: '#1a237e',
  color: '#fff',
  borderRadius: 8,
  padding: '8px 10px',
  cursor: 'pointer'
};
