'use client';

import { useState } from 'react';
import { gasApi } from '../../lib/api';

type Props = {
  onLoginSuccess: (token: string) => void;
};

export function LoginPanel({ onLoginSuccess }: Props) {
  const [step, setStep] = useState<1 | 2>(1);
  const [email, setEmail] = useState('');
  const [code, setCode] = useState('');
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState('');

  const requestOtp = async () => {
    setLoading(true);
    setMessage('');
    try {
      const res = (await gasApi.requestLoginOtp(email)) as { success: boolean; message?: string };
      if (res.success) {
        setStep(2);
      } else {
        setMessage(res.message || 'OTP送信に失敗しました');
      }
    } catch (e) {
      setMessage(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  };

  const verifyOtp = async () => {
    setLoading(true);
    setMessage('');
    try {
      const res = (await gasApi.verifyLoginOtp(email, code)) as {
        success: boolean;
        token?: string;
        message?: string;
      };
      if (res.success && res.token) {
        onLoginSuccess(res.token);
      } else {
        setMessage(res.message || '認証に失敗しました');
      }
    } catch (e) {
      setMessage(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="card" style={{ padding: 20 }}>
      <h2 style={{ marginTop: 0 }}>ログイン</h2>
      {step === 1 ? (
        <div style={{ display: 'grid', gap: 10 }}>
          <label>
            <div>メールアドレス</div>
            <input
              style={input}
              type="email"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              placeholder="example@example.com"
            />
          </label>
          <button style={primaryBtn} onClick={requestOtp} disabled={loading || !email}>認証コードを送信</button>
        </div>
      ) : (
        <div style={{ display: 'grid', gap: 10 }}>
          <label>
            <div>認証コード</div>
            <input
              style={input}
              value={code}
              onChange={(e) => setCode(e.target.value)}
              placeholder="123456"
            />
          </label>
          <div className="row">
            <button style={secondaryBtn} onClick={() => setStep(1)} disabled={loading}>戻る</button>
            <button style={primaryBtn} onClick={verifyOtp} disabled={loading || !code}>ログイン</button>
          </div>
        </div>
      )}
      {message ? <p style={{ color: '#dc2626' }}>{message}</p> : null}
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

const secondaryBtn: any = {
  border: '1px solid #d1d5db',
  background: '#fff',
  color: '#111827',
  borderRadius: 8,
  padding: '10px 14px',
  cursor: 'pointer'
};
