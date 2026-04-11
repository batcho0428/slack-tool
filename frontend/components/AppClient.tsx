'use client';

import { useEffect, useState } from 'react';
import { gasApi } from '../lib/api';
import { AppHeader } from './layout/AppHeader';
import { LoginPanel } from './auth/LoginPanel';
import { DmTab } from './tabs/DmTab';
import { ChannelTab } from './tabs/ChannelTab';
import { MyPageTab } from './tabs/MyPageTab';
import { RosterTab } from './tabs/RosterTab';
import { SurveyTab } from './tabs/SurveyTab';
import { CollectionsTab } from './tabs/CollectionsTab';
import type { LoginUserResponse } from '../lib/types';

const SESSION_KEY = 'slack_app_session';

type TabKey = 'dm' | 'channel' | 'mypage' | 'roster' | 'survey' | 'collections';

export function AppClient() {
  const [login, setLogin] = useState<LoginUserResponse | null>(null);
  const [booting, setBooting] = useState(true);
  const [tab, setTab] = useState<TabKey>('dm');
  const [message, setMessage] = useState('');

  const sessionToken = typeof window === 'undefined' ? '' : (localStorage.getItem(SESSION_KEY) || '');

  const refreshLogin = async () => {
    try {
      const token = localStorage.getItem(SESSION_KEY) || '';
      const res = (await gasApi.getLoginUser(token)) as LoginUserResponse;
      setLogin(res);
    } catch (e) {
      setLogin({ status: 'error', message: e instanceof Error ? e.message : String(e) });
    } finally {
      setBooting(false);
    }
  };

  useEffect(() => {
    const readTokenFromUrl = (): string => {
      const hash = window.location.hash || '';
      const hashText = hash.startsWith('#') ? hash.slice(1) : hash;
      const hashParams = new URLSearchParams(hashText);
      const fromHash = hashParams.get('sessionToken') || hashParams.get('session_token') || '';
      if (fromHash) return fromHash;

      const qs = new URLSearchParams(window.location.search || '');
      return qs.get('sessionToken') || qs.get('session_token') || '';
    };

    const tokenFromUrl = readTokenFromUrl();
    if (tokenFromUrl) {
      localStorage.setItem(SESSION_KEY, tokenFromUrl);
      if (window.location.hash || window.location.search) {
        window.history.replaceState({}, document.title, window.location.pathname);
      }
    }
    refreshLogin();
  }, []);

  const onLoginSuccess = (token: string) => {
    localStorage.setItem(SESSION_KEY, token);
    refreshLogin();
  };

  const logout = () => {
    localStorage.removeItem(SESSION_KEY);
    setLogin({ status: 'guest' });
  };

  return (
    <main>
      <AppHeader appName="45th NUTFES 実行委員マスタ" userName={login?.user?.name} onLogout={logout} />
      <div className="container" style={{ padding: '18px 0 28px' }}>
        {booting ? <p>初期化中...</p> : null}

        {!booting && login?.status !== 'authorized' ? (
          <div style={{ display: 'grid', gap: 12 }}>
            <LoginPanel onLoginSuccess={onLoginSuccess} />
            {login?.message ? <p style={{ color: '#b91c1c' }}>{login.message}</p> : null}
          </div>
        ) : null}

        {!booting && login?.status === 'authorized' ? (
          <div style={{ display: 'grid', gap: 12 }}>
            <nav className="card" style={{ padding: 10, display: 'flex', gap: 8, flexWrap: 'wrap' }}>
              {[
                ['dm', '一斉送信'],
                ['channel', '招待'],
                ['mypage', 'ユーザー管理'],
                ['roster', '名簿'],
                ['survey', 'アンケート'],
                ['collections', '集金']
              ].map(([key, label]) => (
                <button
                  key={key}
                  onClick={() => setTab(key as TabKey)}
                  style={{
                    border: tab === key ? '1px solid #1a237e' : '1px solid #d1d5db',
                    color: tab === key ? '#1a237e' : '#374151',
                    background: tab === key ? '#eef2ff' : '#fff',
                    borderRadius: 8,
                    padding: '8px 10px',
                    cursor: 'pointer'
                  }}
                >
                  {label}
                </button>
              ))}
            </nav>

            {tab === 'dm' ? <DmTab sessionToken={sessionToken} /> : null}
            {tab === 'channel' ? <ChannelTab sessionToken={sessionToken} /> : null}
            {tab === 'mypage' ? <MyPageTab sessionToken={sessionToken} /> : null}
            {tab === 'roster' ? <RosterTab sessionToken={sessionToken} /> : null}
            {tab === 'survey' ? <SurveyTab sessionToken={sessionToken} /> : null}
            {tab === 'collections' ? <CollectionsTab sessionToken={sessionToken} /> : null}
            {message ? <p>{message}</p> : null}
          </div>
        ) : null}
      </div>
    </main>
  );
}
