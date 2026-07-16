// @ts-nocheck
'use client';
/* eslint-disable */

import { useEffect, useState } from 'react';
import DmTab from './tabs/DmTab';
import ChannelTab from './tabs/ChannelTab';
import MyPageTab from './tabs/MyPageTab';
import SurveyTab from './tabs/SurveyTab';
import CollectTab from './tabs/CollectTab';
import RosterTab from './tabs/RosterTab';
import DialogModal from './shared/DialogModal';

declare global {
    interface Window {
        APP_NAME: string;
        APP_HEADER_COLOR: string;
    }
}

const APP_NAME = process.env.NEXT_PUBLIC_APP_NAME || 'Slack送信ツール';
const APP_HEADER_COLOR = process.env.NEXT_PUBLIC_APP_HEADER_COLOR || '#1a237e';
const API_TIMEOUT_MS = 60000;
const API_TIMEOUT_OVERRIDES = {
    collectSurveyReminderStatus: 180000,
    sendSurveyReminderDMs: 300000
};

if (typeof window !== 'undefined') {
    window.APP_NAME = APP_NAME;
    window.APP_HEADER_COLOR = APP_HEADER_COLOR;
}

const runGas = (funcName, ...args) => {
    const timeoutMs = API_TIMEOUT_OVERRIDES[funcName] || API_TIMEOUT_MS;
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    const toFriendlyError = (err) => {
        const msg = err && err.message ? String(err.message) : String(err);
        if (/aborted|timeout/i.test(msg)) {
            return '通信がタイムアウトしました。しばらくしてから再試行してください。';
        }
        if (/fetch failed|failed to fetch|networkerror/i.test(msg)) {
            return '通信に失敗しました。サーバーが起動しているか、ネットワーク設定と /api/gas の疎通を確認してください。';
        }
        return msg;
    };

    // APIごとにpayloadを組み立てる
    let payload = {};
    switch (funcName) {
        case 'getLoginUser':
        case 'getUserProfile':
        case 'listSurveys':
        case 'getAuthUrl':
        case 'listFormDefinitions':
        case 'listCollections':
        case 'getChannels':
        case 'getSearchOptions':
            payload.sessionToken = args[0];
            if (funcName === 'getUserProfile' && args[1]) payload.targetEmail = args[1];
            break;
        case 'inviteToChannel':
            payload.sessionToken = args[0];
            payload.channelId = args[1];
            payload.recipients = args[2] || [];
            break;
        case 'sendDMs':
            payload.sessionToken = args[0];
            payload.message = args[1];
            payload.recipients = args[2] || [];
            break;
        case 'createUser':
            payload.sessionToken = args[0];
            payload.userObj = args[1] || {};
            break;
        case 'updateUserProfile':
            payload.sessionToken = args[0];
            payload.formData = args[1] || {};
            payload.targetEmail = args[2];
            break;
        case 'getSurveyDetails':
            payload.sessionToken = args[0];
            payload.spreadsheetRef = args[1];
            payload.rowIndex = args[2];
            break;
        case 'createCollection':
        case 'updateCollection':
        case 'saveFormDefinition':
            payload.sessionToken = args[0];
            if (funcName === 'updateCollection') {
                payload.collectionId = args[1];
                payload.payload = args[2] || {};
            } else {
                payload.payload = args[1] || {};
            }
            break;
        case 'collectSurveyReminderStatus':
            payload.sessionToken = args[0];
            payload.surveyRowIndices = args[1] || [];
            break;
        case 'sendSurveyReminderDMs':
            payload.sessionToken = args[0];
            payload.payload = args[1] || {};
            break;
        case 'deleteCollection':
            payload.sessionToken = args[0];
            payload.collectionId = args[1];
            break;
        case 'fetchCollectionSummary':
            payload.sessionToken = args[0];
            payload.collectionId = args[1];
            break;
        case 'getCollectionRowDetails':
            payload.sessionToken = args[0];
            payload.collectionId = args[1];
            payload.recipientEmail = args[2];
            break;
        case 'recordPayment':
            payload.sessionToken = args[0];
            payload.collectionId = args[1];
            payload.recipientEmail = args[2];
            payload.amount = args[3];
            payload.type = args[4];
            payload.handlerEmail = args[5];
            break;
        case 'recordPaymentWithChange':
            payload.sessionToken = args[0];
            payload.collectionId = args[1];
            payload.recipientEmail = args[2];
            payload.receivedAmount = args[3];
            payload.expectedAmount = args[4];
            payload.handlerEmail = args[5];
            break;
        case 'createRosterCsv':
            payload.sessionToken = args[0];
            payload.params = args[1] || {};
            break;
        case 'requestLoginOtp':
            payload.email = args[0];
            break;
        case 'verifyLoginOtp':
            payload.email = args[0];
            payload.code = args[1];
            break;
        case 'handleSlackOAuthCode':
            payload.code = args[0];
            payload.redirectUri = args[1];
            break;
        default:
            // fallback: legacy互換
            payload.__args = args;
    }

    return fetch('/api/gas', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: funcName, payload }),
        cache: 'no-store',
        signal: controller.signal
    })
    .then(async (res) => {
            const text = await res.text();
            let json = null;
            try {
                json = JSON.parse(text);
            } catch (e) {
                const snippet = String(text || '').slice(0, 120);
                throw new Error(`API応答がJSONではありません: ${snippet}`);
            }
      if (!json || !json.ok) {
        throw new Error((json && json.error) || 'API request failed');
      }
      return json.result;
        })
        .catch((e) => {
            throw new Error(toFriendlyError(e));
        })
        .finally(() => {
            clearTimeout(timer);
        });
};

const extractAuthUrl = (value) => {
    if (typeof value === 'string') return value.trim();
    if (value && typeof value === 'object' && value.url) return String(value.url).trim();
    return '';
};

const fetchAuthUrl = async () => {
    const raw = await runGas('getAuthUrl');
    const url = extractAuthUrl(raw);
    if (!url) {
        throw new Error('Slack連携URLを取得できませんでした。設定を確認してください。');
    }
    return url;
};


        // CSV parser (handles quoted fields).
        const parseCsv = (text) => {
            if (!text) return { headers: [], rows: [] };
            const rows = [];
            let cur = '';
            let inQuotes = false;
            const row = [];
            for (let i = 0; i < text.length; i++) {
                const ch = text[i];
                const next = text[i+1];
                if (inQuotes) {
                    if (ch === '"') {
                        if (next === '"') { cur += '"'; i++; }
                        else { inQuotes = false; }
                    } else {
                        cur += ch;
                    }
                } else {
                    if (ch === '"') { inQuotes = true; }
                    else if (ch === ',') { row.push(cur); cur = ''; }
                    else if (ch === '\r') { continue; }
                    else if (ch === '\n') { row.push(cur); rows.push(row.slice()); row.length = 0; cur = ''; }
                    else { cur += ch; }
                }
            }
            if (cur !== '' || row.length > 0) { row.push(cur); rows.push(row.slice()); }
            if (rows.length === 0) return { headers: [], rows: [] };
            const headers = rows[0];
            const body = rows.slice(1).filter(r => r.length > 0 && !(r.length === 1 && r[0] === ''));
            return { headers, rows: body };
        };

        // Normalize birthday string to YYYY/MM/DD (zero-padded)
        const formatPreviewBirthday = (s) => {
            if (!s) return '';
            const str = String(s).trim();
            // try to match YYYY MM DD with any separator
            let m = str.match(/(\d{4})[^0-9]?(\d{1,2})[^0-9]?(\d{1,2})/);
            if (m) {
                const y = m[1];
                const mo = ('0' + m[2]).slice(-2);
                const d = ('0' + m[3]).slice(-2);
                return `${y}/${mo}/${d}`;
            }
            // try to match YY MM DD and assume 20xx for reasonable range
            m = str.match(/^(\d{2})[^0-9]?(\d{1,2})[^0-9]?(\d{1,2})$/);
            if (m) {
                const yy = parseInt(m[1], 10);
                const y = (yy <= 30) ? ('20' + m[1]) : ('19' + m[1]);
                const mo = ('0' + m[2]).slice(-2);
                const d = ('0' + m[3]).slice(-2);
                return `${y}/${mo}/${d}`;
            }
            // fallback: return original trimmed
            // try Date.parse fallback (handles full Date string like 'Wed Apr 28 2004 ...')
            const parsed = Date.parse(str);
            if (!isNaN(parsed)) {
                const d = new Date(parsed);
                const y = d.getFullYear();
                const mo = ('0' + (d.getMonth() + 1)).slice(-2);
                const da = ('0' + d.getDate()).slice(-2);
                return `${y}/${mo}/${da}`;
            }
            return str;
        };

        const processInBatches = async (items, batchSize, processFunction, onProgress) => {
            let successTotal = 0;
            let failedListTotal = [];
            for (let i = 0; i < items.length; i += batchSize) {
                const batch = items.slice(i, i + batchSize);
                try {
                    const res = await processFunction(batch);
                    successTotal += res.success;
                    failedListTotal = [...failedListTotal, ...res.failed];
                } catch (e) {
                    const batchFailures = batch.map(b => ({ email: b.email, error: e.message || "Batch Error" }));
                    failedListTotal = [...failedListTotal, ...batchFailures];
                }
                if (onProgress) onProgress(Math.min(100, Math.round(((i + batch.length) / items.length) * 100)));
            }
            return { success: successTotal, failed: failedListTotal };
        };

        function App() {
            const [view, setView] = useState('loading');
            const [user, setUser] = useState(null);
            const [authUrl, setAuthUrl] = useState('');
            const [errorMsg, setErrorMsg] = useState('');
            const [logoutDialog, setLogoutDialog] = useState(false);

            useEffect(() => {
                const readTokenFromUrl = () => {
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
                    localStorage.setItem('slack_app_session', tokenFromUrl);
                    if (window.location.hash || window.location.search) {
                        window.history.replaceState({}, document.title, window.location.pathname);
                    }
                }
                const token = localStorage.getItem('slack_app_session');
                checkLogin(token);
            }, []);

            const checkLogin = (token) => {
                runGas('getLoginUser', token).then(res => {
                    if (res.status === 'authorized') {
                        runGas('getUserProfile', token)
                            .then((profile) => {
                                setUser({ ...res.user, hasToken: !!res.hasToken, isAdmin: !!(profile && profile.isAdmin) });
                                setView('main');
                                if (!res.hasToken) {
                                    fetchAuthUrl().then(setAuthUrl).catch((e) => setErrorMsg(e.message || String(e)));
                                }
                            })
                            .catch(() => {
                                setUser({ ...res.user, hasToken: !!res.hasToken, isAdmin: false });
                                setView('main');
                            });
                    }

                    else if (res.status === 'guest') {
                        fetchAuthUrl()
                            .then((url) => {
                                setAuthUrl(url);
                                setView('guest');
                            })
                            .catch((e) => {
                                setErrorMsg(e.message || String(e));
                                setView('error');
                            });
                    }
                    else { throw new Error(res.message); }
                }).catch(e => { setErrorMsg(e.message); setView('error'); });
            };




            const executeLogout = () => {
                setLogoutDialog(false);
                localStorage.removeItem('slack_app_session');
                setUser(null);
                setAuthUrl('');
                fetchAuthUrl().then(setAuthUrl).catch(() => {});
                setView('guest');
            };

            return (
                <div className="min-h-screen flex flex-col items-center pt-2 px-2 pb-2 md:pt-6 md:px-4 md:pb-10 bg-[#f3f4f6]">
                    <div className="w-full max-w-4xl bg-white rounded-lg md:rounded-xl shadow-xl overflow-hidden flex flex-col h-[calc(100dvh-1rem)] md:h-auto md:min-h-[700px]">
                        <div style={{ backgroundColor: APP_HEADER_COLOR }} className="p-3 md:p-4 text-white flex justify-between items-center shadow-md z-10 shrink-0">
                            <h1 className="font-bold text-base md:text-lg tracking-wide flex items-center">
                                <i className="fab fa-slack mr-2 text-xl"></i> {APP_NAME}
                            </h1>
                            {user && (
                                <div className="flex items-center text-sm bg-white/10 px-2 py-1 md:px-3 rounded-full">
                                    <span className="mr-1 md:mr-2 truncate max-w-[100px] md:max-w-[150px] text-xs md:text-sm">{user.name}</span>
                                    {user.isAdmin && (
                                        <button onClick={() => { window.location.href = '/admin'; }} className="hover:text-yellow-200 ml-1 md:ml-2 p-1" title="管理者ページ">
                                            <i className="fas fa-user-shield"></i>
                                        </button>
                                    )}
                                    <button onClick={()=>setLogoutDialog(true)} className="hover:text-red-200 ml-1 md:ml-2 p-1" title="ログアウト">
                                        <i className="fas fa-sign-out-alt"></i>
                                    </button>
                                </div>
                            )}
                        </div>

                        <div className="flex-1 relative bg-white flex flex-col overflow-hidden">
                            {view === 'loading' && <div className="absolute inset-0 flex flex-col items-center justify-center text-gray-500"><div className="animate-spin rounded-full h-10 w-10 border-b-2 border-gray-900 mb-4"></div><p>読み込み中...</p></div>}
                            {view === 'error' && (
                                <div className="p-4 md:p-8 flex flex-col items-center justify-center h-full">
                                    <div className="bg-red-50 text-red-700 p-6 rounded-lg text-center max-w-md w-full border border-red-200">
                                        <h3 className="font-bold text-lg mb-2">エラー</h3>
                                        <p className="text-sm mb-4">{errorMsg}</p>
                                        <button onClick={()=>window.location.reload()} className="bg-red-600 text-white px-6 py-2 rounded hover:bg-red-700">再読み込み</button>
                                    </div>
                                </div>
                            )}
                            {view === 'guest' && <GuestScreen authUrl={authUrl} onLoginSuccess={checkLogin} />}
                            {view === 'main' && user && <MainScreen user={user} authUrl={authUrl} />}
                        </div>
                    </div>
                    <DialogModal isOpen={logoutDialog} type="confirm" message="ログアウトしますか？" onOk={executeLogout} onCancel={()=>setLogoutDialog(false)} />
                </div>
            );
        }

        // ゲスト画面 (Bot OTP認証対応)
        function GuestScreen({ authUrl, onLoginSuccess }) {
            const [step, setStep] = useState(1); // 1:Email, 2:Code
            const [email, setEmail] = useState('');
            const [code, setCode] = useState('');
            const [loading, setLoading] = useState(false);
            const [error, setError] = useState('');

            const openSlackLogin = async () => {
                try {
                    const url = authUrl || await fetchAuthUrl();
                    window.location.href = url;
                } catch (e) {
                    setError(e instanceof Error ? e.message : String(e));
                }
            };

            const handleRequestOtp = async (e) => {
                e.preventDefault();
                setLoading(true); setError('');
                try {
                    const res = await runGas('requestLoginOtp', email);
                    if (res.success) setStep(2);
                    else setError(res.message);
                } catch(e) { setError(e.message); }
                finally { setLoading(false); }
            };

            const handleVerifyOtp = async (e) => {
                e.preventDefault();
                setLoading(true); setError('');
                try {
                    const res = await runGas('verifyLoginOtp', email, code);
                    if (res.success) {
                        localStorage.setItem('slack_app_session', res.token);
                        onLoginSuccess(res.token);
                    } else setError(res.message);
                } catch(e) { setError(e.message); }
                finally { setLoading(false); }
            };

            return (
                <div className="flex flex-col items-center justify-start h-full p-4 md:p-6 overflow-y-auto">
                    <h2 className="text-xl md:text-2xl font-bold text-gray-800 mb-4 md:mb-6 mt-2">ログイン</h2>

                    <div className="w-full max-w-md bg-white border border-gray-200 rounded-lg p-6 shadow-sm mb-6">
                        <div className="flex items-center mb-4 text-gray-700 font-bold border-b pb-2">
                            <i className="fas fa-lock mr-2 text-blue-600"></i>
                            認証コードでログイン
                        </div>

                        {step === 1 ? (
                            <form onSubmit={handleRequestOtp} className="space-y-4">
                                <p className="text-sm text-gray-600">
                                    登録済みのメールアドレスを入力してください。<br/>Slack Botから認証コードが送信されます。
                                </p>
                                <div>
                                    <label className="block text-xs font-bold text-gray-500 mb-1">メールアドレス</label>
                                    <input required type="email" className="w-full border p-3 rounded text-sm focus:ring-2 focus:ring-blue-400 outline-none" value={email} onChange={e=>setEmail(e.target.value)} placeholder="example@ex.com" />
                                </div>
                                <button type="submit" disabled={loading} className={`w-full py-3 rounded-lg font-bold text-sm text-white shadow transition ${loading ? 'bg-gray-400' : 'bg-blue-600 hover:bg-blue-700'}`}>
                                    {loading ? <i className="fas fa-spinner fa-spin"></i> : '認証コードを送信'}
                                </button>
                            </form>
                        ) : (
                            <form onSubmit={handleVerifyOtp} className="space-y-4">
                                <p className="text-sm text-gray-600">
                                    Slack Botから送信された6桁のコードを入力してください。
                                </p>
                                <div>
                                    <label className="block text-xs font-bold text-gray-500 mb-1">認証コード</label>
                                    <input required type="text" className="w-full border p-3 rounded text-lg tracking-widest text-center focus:ring-2 focus:ring-blue-400 outline-none" value={code} onChange={e=>setCode(e.target.value)} placeholder="123456" maxLength={6} />
                                </div>
                                {error && <div className="text-red-600 text-xs bg-red-50 p-2 rounded">{error}</div>}
                                <button type="submit" disabled={loading} className={`w-full py-3 rounded-lg font-bold text-sm text-white shadow transition ${loading ? 'bg-gray-400' : 'bg-green-600 hover:bg-green-700'}`}>
                                    {loading ? <i className="fas fa-spinner fa-spin"></i> : 'ログイン'}
                                </button>
                                <button type="button" onClick={()=>setStep(1)} className="w-full text-xs text-gray-500 hover:underline">メールアドレス入力に戻る</button>
                            </form>
                        )}
                        {step === 1 && error && <div className="mt-3 text-red-600 text-xs bg-red-50 p-2 rounded">{error}</div>}
                    </div>

                    <div className="w-full max-w-md text-center">
                        <p className="text-xs text-gray-500 mb-2">- または -</p>
                        <button onClick={openSlackLogin} className="w-full bg-[#4A154B] text-white font-bold py-3 rounded-lg shadow hover:bg-[#381039] transition flex items-center justify-center text-sm">
                            <i className="fab fa-slack text-lg mr-2"></i> Slackアカウントでログイン (PC推奨)
                        </button>
                    </div>
                </div>
            );
        }

        function MainScreen({ user, authUrl }) {
            const [tab, setTab] = useState('dm');
            return (
                <div className="flex flex-col h-full">
                    <div className="grid grid-cols-3 md:flex border-b border-gray-200 shrink-0 bg-gray-50">
                        <button onClick={() => setTab('dm')} className={`w-full md:flex-1 py-3 text-center font-bold text-sm border-b-2 transition ${tab === 'dm' ? 'border-blue-600 text-blue-600 bg-white' : 'border-transparent text-gray-500 hover:bg-gray-100'}`}>
                            <i className="far fa-paper-plane mr-2"></i> 一斉送信
                        </button>
                        <button onClick={() => setTab('channel')} className={`w-full md:flex-1 py-3 text-center font-bold text-sm border-b-2 transition ${tab === 'channel' ? 'border-green-600 text-green-600 bg-white' : 'border-transparent text-gray-500 hover:bg-gray-100'}`}>
                            <i className="fas fa-user-plus mr-2"></i> 招待
                        </button>
                        <button onClick={() => setTab('mypage')} className={`w-full md:flex-1 py-3 text-center font-bold text-sm border-b-2 transition ${tab === 'mypage' ? 'border-purple-600 text-purple-600 bg-white' : 'border-transparent text-gray-500 hover:bg-gray-100'}`}>
                            <i className="fas fa-id-card mr-2"></i> ユーザー管理
                        </button>
                        <button onClick={() => setTab('roster')} className={`w-full md:flex-1 py-3 text-center font-bold text-sm border-b-2 transition ${tab === 'roster' ? 'border-indigo-600 text-indigo-600 bg-white' : 'border-transparent text-gray-500 hover:bg-gray-100'}`}>
                            <i className="fas fa-list mr-2"></i> 名簿
                        </button>
                        <button onClick={() => setTab('survey')} className={`w-full md:flex-1 py-3 text-center font-bold text-sm border-b-2 transition ${tab === 'survey' ? 'border-rose-600 text-rose-600 bg-white' : 'border-transparent text-gray-500 hover:bg-gray-100'}`}>
                            <i className="fas fa-poll mr-2"></i> アンケート
                        </button>
                        <button onClick={() => setTab('collections')} className={`w-full md:flex-1 py-3 text-center font-bold text-sm border-b-2 transition ${tab === 'collections' ? 'border-amber-600 text-amber-600 bg-white' : 'border-transparent text-gray-500 hover:bg-gray-100'}`}>
                            <i className="fas fa-wallet mr-2"></i> 集金
                        </button>
                    </div>
                    <div className="flex-1 overflow-hidden relative">
                        {tab === 'dm' && <DmTab user={user} authUrl={authUrl} runGas={runGas} fetchAuthUrl={fetchAuthUrl} processInBatches={processInBatches} />}
                        {tab === 'channel' && <ChannelTab user={user} authUrl={authUrl} runGas={runGas} fetchAuthUrl={fetchAuthUrl} processInBatches={processInBatches} />}
                        {tab === 'mypage' && <MyPageTab runGas={runGas} />}
                        {tab === 'roster' && <RosterTab user={user} runGas={runGas} parseCsv={parseCsv} formatPreviewBirthday={formatPreviewBirthday} />}
                        {tab === 'survey' && <SurveyTab user={user} runGas={runGas} />}
                        {tab === 'collections' && <CollectTab user={user} runGas={runGas} parseCsv={parseCsv} formatPreviewBirthday={formatPreviewBirthday} />}
                    </div>
                </div>
            );
        }


export function LegacyAppPage() {
  return (
    <>
      <style jsx global>{`
        body { font-family: 'Helvetica Neue', Arial, 'Hiragino Kaku Gothic ProN', 'Hiragino Sans', Meiryo, sans-serif; background-color: #f3f4f6; color: #1d1c1d; }
        .modal-overlay { background-color: rgba(0, 0, 0, 0.5); backdrop-filter: blur(2px); }
        ::-webkit-scrollbar { width: 6px; }
        ::-webkit-scrollbar-track { background: #f1f1f1; }
        ::-webkit-scrollbar-thumb { background: #ccc; border-radius: 4px; }
        @keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
        .animate-fade-in { animation: fadeIn 0.3s ease-in-out; }
        button, input, select, textarea { touch-action: manipulation; }
      `}</style>
      <App />
    </>
  );
}
