// @ts-nocheck
'use client';
/* eslint-disable */

import { useEffect, useRef, useState } from 'react';
import SearchModal from '../shared/SearchModal';
import RecipientSelector from '../shared/RecipientSelector';
import ResultModal from '../shared/ResultModal';

        export default function DmTab({ user, authUrl, runGas, fetchAuthUrl, processInBatches, DialogModal }) {
            const [message, setMessage] = useState('');
            const [recipients, setRecipients] = useState([]);
            const [modalOpen, setModalOpen] = useState(false);
            const [sending, setSending] = useState(false);
            const [progress, setProgress] = useState(0);
            const [result, setResult] = useState(null);
            const [dialog, setDialog] = useState({ isOpen: false, type: 'alert', message: '', onOk: null, onCancel: null });
            const txtRef = useRef(null);
            const [loginUrl, setLoginUrl] = useState(authUrl || '');

            useEffect(() => {
                if (!loginUrl) fetchAuthUrl().then(setLoginUrl).catch(()=>{});
            }, []);

            if (user && !user.hasToken) {
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


            const insertMention = () => {
                const t = txtRef.current;
                if(t){
                    const v = t.value;
                    const p = t.selectionStart;
                    setMessage(v.slice(0,p)+"{mention}"+v.slice(t.selectionEnd));
                    setTimeout(() => t.focus(), 0);
                }
            };

            const closeDialog = () => setDialog(prev => ({ ...prev, isOpen: false }));

            const handleSendClick = () => {
                if(!recipients.length) {
                    setDialog({ isOpen: true, type: 'alert', message: "送信先を選択してください", onOk: closeDialog });
                    return;
                }
                if(!message.trim()) {
                    setDialog({ isOpen: true, type: 'alert', message: "メッセージを入力してください", onOk: closeDialog });
                    return;
                }
                setDialog({
                    isOpen: true, type: 'confirm', message: "【確認】本当に送信しますか？",
                    onOk: executeSend, onCancel: closeDialog
                });
            };

            const executeSend = async () => {
                closeDialog();
                setSending(true); setProgress(0); setResult(null);
                const token = localStorage.getItem('slack_app_session');
                try {
                    const res = await processInBatches(recipients, 10, (batch) => runGas('sendDMs', token, message, batch), (pct) => setProgress(pct));
                    const failedEmails = new Set(res.failed.map(f => f.email));
                    const failedDetails = res.failed.map(f => {
                        const original = recipients.find(r => r.email === f.email);
                        return { email: f.email, error: f.error, name: original ? original.name : '不明' };
                    });
                    setRecipients(recipients.filter(r => failedEmails.has(r.email)));
                    setResult({ success: res.success, failed: failedDetails });
                } catch(e) {
                    setDialog({ isOpen: true, type: 'alert', message: "エラー: " + e.message, onOk: closeDialog });
                } finally { setSending(false); }
            };

            return (
                <div className="flex flex-col h-full p-3 md:p-6 space-y-3 md:space-y-4 overflow-y-auto">
                    <div className="space-y-2 shrink-0">
                        <div className="flex justify-between items-end">
                            <label className="font-bold text-gray-700 block text-sm md:text-base">メッセージ本文</label>
                            <button onClick={insertMention} className="text-xs bg-blue-50 text-blue-600 border border-blue-200 px-3 py-1 rounded active:bg-blue-100">@メンション</button>
                        </div>
                        <textarea ref={txtRef} className="w-full border border-gray-300 p-2 md:p-3 h-24 md:h-32 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none resize-none shadow-sm text-sm" value={message} onChange={e=>setMessage(e.target.value)}></textarea>
                    </div>
                    <div className="flex-1 min-h-0">
                        <RecipientSelector recipients={recipients} setRecipients={setRecipients} onAddClick={()=>setModalOpen(true)} />
                    </div>
                    <div className="pt-2 md:pt-4 shrink-0 pb-2">
                        <button onClick={handleSendClick} disabled={sending || !recipients.length} className={`w-full py-3 md:py-4 rounded-lg font-bold text-base md:text-lg shadow-md transition active:scale-95 ${sending?'bg-gray-400':'bg-blue-600 hover:bg-blue-700 text-white'}`}>
                            {sending ? `送信中 (${progress}%)` : <span><i className="far fa-paper-plane mr-2"></i>DMを一斉送信</span>}
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

