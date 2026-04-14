'use client';

import { useEffect, useMemo, useState } from 'react';
import { useRouter, useSearchParams } from 'next/navigation';

type CallbackState = 'processing' | 'success' | 'error';

function buildNetworkErrorMessage(err: unknown): string {
  const msg = err instanceof Error ? err.message : String(err);
  if (/fetch failed|failed to fetch|networkerror/i.test(msg)) {
    return '通信に失敗しました。サーバーが起動しているか、ネットワーク設定と /api/gas の疎通を確認してください。';
  }
  return msg;
}

export function SlackCallbackClient() {
  const router = useRouter();
  const searchParams = useSearchParams();
  const [state, setState] = useState<CallbackState>('processing');
  const [message, setMessage] = useState('認証を処理しています...');

  const code = useMemo(() => searchParams.get('code') || '', [searchParams]);
  const oauthError = useMemo(() => searchParams.get('error') || '', [searchParams]);
  const oauthErrorDescription = useMemo(
    () => searchParams.get('error_description') || '',
    [searchParams]
  );

  useEffect(() => {
    const run = async () => {
      // callback URL を直接開いたケースではトップへ戻す
      if (!code && !oauthError) {
        setState('processing');
        setMessage('トップへ移動しています...');
        router.replace('/');
        return;
      }

      if (oauthError) {
        setState('error');
        setMessage(`Slack認証エラー: ${oauthErrorDescription || oauthError}`);
        return;
      }

      if (!code) {
        setState('error');
        setMessage('認証コードが見つかりません。');
        return;
      }

      try {
        const redirectUri = `${window.location.origin}/auth/slack/callback`;
        let res: Response;
        try {
          res = await fetch('/api/gas', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              action: 'handleSlackOAuthCode',
              payload: { code, redirectUri }
            })
          });
        } catch (e) {
          throw new Error(buildNetworkErrorMessage(e));
        }

        const text = await res.text();
        let json: any;
        try {
          json = JSON.parse(text);
        } catch {
          const snippet = String(text || '').slice(0, 120);
          throw new Error(`API応答がJSONではありません: ${snippet}`);
        }
        if (!json || !json.ok) {
          throw new Error((json && json.error) || 'OAuth交換に失敗しました');
        }

        const result = json.result || {};
        if (!result.success || !result.sessionToken) {
          throw new Error(result.message || 'セッション作成に失敗しました');
        }

        localStorage.setItem('slack_app_session', String(result.sessionToken));
        setState('success');
        setMessage('認証が完了しました。トップへ移動します...');
        router.replace('/');
      } catch (e) {
        setState('error');
        setMessage(buildNetworkErrorMessage(e));
      }
    };

    run();
  }, [code, oauthError, oauthErrorDescription, router]);

  return (
    <main style={{ minHeight: '100vh', display: 'grid', placeItems: 'center', padding: 16 }}>
      <div style={{ width: 'min(520px, 100%)', border: '1px solid #e5e7eb', borderRadius: 12, padding: 20 }}>
        <h1 style={{ marginTop: 0, fontSize: 20 }}>Slack 認証</h1>
        <p style={{ marginBottom: 0, color: state === 'error' ? '#b91c1c' : '#374151' }}>{message}</p>
      </div>
    </main>
  );
}
