// @ts-nocheck
'use client';
/* eslint-disable */

import { useEffect, useState } from 'react';
import SearchModal from '../shared/SearchModal';
import RecipientSelector from '../shared/RecipientSelector';
import ResultModal from '../shared/ResultModal';

        export default function ChannelTab({ user, authUrl, runGas, fetchAuthUrl, processInBatches, DialogModal }) {
            const [channels, setChannels] = useState([]);
            const [selectedChannel, setSelectedChannel] = useState('');
            const [recipients, setRecipients] = useState([]);
            const [loadingChannels, setLoadingChannels] = useState(true);
            const [needsAuth, setNeedsAuth] = useState(false);
            const [modalOpen, setModalOpen] = useState(false);
            const [sending, setSending] = useState(false);
            const [progress, setProgress] = useState(0);
            const [result, setResult] = useState(null);
            const [dialog, setDialog] = useState({ isOpen: false, type: 'alert', message: '', onOk: null, onCancel: null });
            const [loginUrl, setLoginUrl] = useState(authUrl || '');

            useEffect(() => {
                if (!loginUrl) fetchAuthUrl().then(setLoginUrl).catch(()=>{});
            }, []);

            const closeDialog = () => setDialog(prev => ({ ...prev, isOpen: false }));

            useEffect(() => {
                let mounted = true;
                setLoadingChannels(true);
                setNeedsAuth(false);

                // user が未セット (初期 null) の間は読み込みスピナーを継続する
                if (user == null) {
                    return () => { mounted = false; };
                }

                // Slack連携がない場合は読み込み表示ののち認証案内へ
                if (!user.hasToken) {
                    if (mounted) {
                        setNeedsAuth(true);
                        setChannels([]);
                        setLoadingChannels(false);
                    }
                    return () => { mounted = false; };
                }

                // 連携済: チャンネル取得
                runGas('getChannels', localStorage.getItem('slack_app_session'))
                    .then(list => { if (mounted) setChannels(list || []); })
                    .catch(e => {
                        if (mounted) {
                            // エラー詳細を表示
                            setDialog({
                                isOpen: true,
                                type: 'alert',
                                message: `チャンネル取得エラー: ${e.message}\n(デバッグ用: ${JSON.stringify(e, Object.getOwnPropertyNames(e))})`,
                                onOk: closeDialog
                            });
                        }
                    })
                    .finally(() => { if (mounted) setLoadingChannels(false); });

                return () => { mounted = false; };
            }, [user]);

            const handleInviteClick = () => {
                if(!selectedChannel) { setDialog({ isOpen: true, type: 'alert', message: "チャンネルを選択してください", onOk: closeDialog }); return; }
                if(!recipients.length) { setDialog({ isOpen: true, type: 'alert', message: "ユーザーを選択してください", onOk: closeDialog }); return; }
                setDialog({ isOpen: true, type: 'confirm', message: "実行しますか？", onOk: executeInvite, onCancel: closeDialog });
            };

            const executeInvite = async () => {
                closeDialog();
                setSending(true); setProgress(0); setResult(null);
                const token = localStorage.getItem('slack_app_session');
                try {
                    const res = await processInBatches(recipients, 5, (batch) => runGas('inviteToChannel', token, selectedChannel, batch), (pct) => setProgress(pct));
                    const failedEmails = new Set(res.failed.map(f => f.email));
                    const failedDetails = res.failed.map(f => {
                        const original = recipients.find(r => r.email === f.email);
                        return { email: f.email, error: f.error, name: original ? original.name : '不明' };
                    });
                    setRecipients(recipients.filter(r => failedEmails.has(r.email)));
                    setResult({ success: res.success, failed: failedDetails });
                } catch(e) { setDialog({ isOpen: true, type: 'alert', message: "エラー: " + e.message, onOk: closeDialog }); }
                finally { setSending(false); }
            };

            // 初期は読み込み中を表示
            if (loadingChannels) {
                return <div className="flex justify-center items-center h-full text-gray-500"><i className="fas fa-spinner fa-spin mr-2"></i>読み込み中...</div>;
            }

            // 認証が必要な場合は案内を表示
            if (needsAuth) {
                return (
                    <div className="flex flex-col items-center justify-center h-full p-6 text-center">
                        <div className="max-w-md w-full bg-white rounded-lg shadow p-6">
                            <h3 className="text-lg font-bold mb-2">Slackアカウントが連携されていません</h3>
                            <p className="text-sm text-gray-600 mb-4">Slackアカウントが連携されていないため、この機能は使用できません。</p>
                            <button onClick={async () => {
                                try {
                                    const url = loginUrl || await fetchAuthUrl();
                                    setLoginUrl(url);
                                    window.location.href = url;
                                } catch (e) {
                                    setDialog({ isOpen: true, type: 'alert', message: e.message || String(e), onOk: closeDialog });
                                }
                            }} className="w-full bg-[#4A154B] text-white font-bold py-3 rounded-lg shadow hover:bg-[#381039] transition flex items-center justify-center text-sm">
                                <i className="fab fa-slack text-lg mr-2"></i> Slackアカウントでログイン
                            </button>
                        </div>
                    </div>
                );
            }

            return (
                <div className="flex flex-col h-full p-3 md:p-6 space-y-3 md:space-y-4 overflow-y-auto">
                    <div className="space-y-2 shrink-0">
                        <label className="font-bold text-gray-700 block text-sm md:text-base">追加先のチャンネル</label>
                        <select value={selectedChannel} onChange={e=>setSelectedChannel(e.target.value)} disabled={loadingChannels} className="w-full border border-gray-300 p-2 md:p-3 rounded-lg bg-white text-sm">
                            <option value="">選択してください...</option>
                            {channels.map(c => <option key={c.id} value={c.id}>{c.is_private ? '🔒 ' : '# '}{c.name}</option>)}
                        </select>
                        <p className="text-xs text-gray-500">※ あなたが参加しているチャンネルのみ表示されます。</p>

                    </div>
                    <div className="flex-1 min-h-0">
                        <RecipientSelector labelText="招待するユーザー" recipients={recipients} setRecipients={setRecipients} onAddClick={()=>setModalOpen(true)} />
                    </div>
                    <div className="pt-2 md:pt-4 shrink-0 pb-2">
                        <button onClick={handleInviteClick} disabled={sending || !recipients.length} className={`w-full py-3 md:py-4 rounded-lg font-bold text-base md:text-lg shadow-md transition active:scale-95 ${sending?'bg-gray-400':'bg-green-600 hover:bg-green-700 text-white'}`}>
                            {sending ? `処理中 (${progress}%)` : <span><i className="fas fa-user-plus mr-2"></i>チャンネルに追加</span>}
                        </button>
                    </div>
                    {modalOpen && <SearchModal runGas={runGas} currentUserEmail={user.email} onClose={()=>setModalOpen(false)} onAdd={(ls)=>{
                        const ids = new Set(recipients.map(r=>r.email));
                        setRecipients([...recipients, ...ls.filter(r=>!ids.has(r.email))]);
                        setModalOpen(false);
                    }} />}
                    {result && <ResultModal result={result} onClose={()=>setResult(null)} />}
                    <DialogModal isOpen={dialog.isOpen} type={dialog.type} message={dialog.message} onOk={dialog.onOk} onCancel={dialog.onCancel} />
                </div>
            );
        }
