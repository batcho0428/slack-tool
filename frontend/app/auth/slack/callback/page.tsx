import { Suspense } from 'react';
import { SlackCallbackClient } from './slack-callback-client';

export default function SlackCallbackPage() {
  return (
    <Suspense
      fallback={
        <main style={{ minHeight: '100vh', display: 'grid', placeItems: 'center', padding: 16 }}>
          <div style={{ width: 'min(520px, 100%)', border: '1px solid #e5e7eb', borderRadius: 12, padding: 20 }}>
            <h1 style={{ marginTop: 0, fontSize: 20 }}>Slack 認証</h1>
            <p style={{ marginBottom: 0, color: '#374151' }}>認証を処理しています...</p>
          </div>
        </main>
      }
    >
      <SlackCallbackClient />
    </Suspense>
  );
}
