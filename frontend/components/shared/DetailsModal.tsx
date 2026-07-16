// @ts-nocheck
'use client';
/* eslint-disable */

        export default function DetailsModal({ isOpen, loading, details, onClose, survey }) {
            if (!isOpen) return null;
            return (
                <div className="fixed inset-0 flex items-center justify-center modal-overlay p-4 z-[210]">
                    <div className="bg-white rounded-xl shadow-2xl w-full max-w-3xl p-4 md:p-6 overflow-auto max-h-[80vh] animate-fade-in">
                        <div className="flex items-center justify-between mb-4">
                            <h3 className="text-lg font-bold">{details && details.viewType === 'collection' ? '集金詳細' : 'アンケート詳細'}</h3>
                            <div className="flex items-center gap-2">
                                {details && details.viewType !== 'collection' && (() => {
                                    const url = survey && String(survey.formUrl || '').trim();
                                    const collecting = (() => {
                                        if (!survey) return false;
                                        if (survey.collecting === true) return true;
                                        const sval = String(survey.collecting || '').trim().toLowerCase();
                                        if (sval === 'true' || sval === '1') return true;
                                        if (typeof survey.collecting === 'number' && Number(survey.collecting) === 1) return true;
                                        return false;
                                    })();
                                    if (url && collecting) {
                                        return (<button onClick={() => { try { window.open(url, '_blank', 'noopener'); } catch (e) { alert('リンクを開けませんでした'); } }} className="text-sm bg-blue-600 text-white px-3 py-1 rounded hover:bg-blue-700">フォームを開く</button>);
                                    }
                                    return (<button disabled className="text-sm bg-gray-100 text-gray-400 px-3 py-1 rounded cursor-not-allowed">フォームを開く</button>);
                                })()}
                                <button onClick={onClose} className="text-sm text-gray-500 hover:text-gray-700">閉じる</button>
                            </div>
                        </div>

                        {/* Details only; dropdown removed to avoid referencing outer scope variables */}
                        {loading ? (
                            <div className="flex flex-col items-center justify-center py-8"><div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mb-4"></div><div className="text-sm text-gray-600">読み込み中...</div></div>
                        ) : (
                            <>
                            <div className="space-y-4 text-sm text-gray-800">
                                {details && details.error && (<div className="text-red-600">{details.error}</div>)}
                                {details && details.response && (
                                    <div>
                                        <div className="font-medium mb-2">あなたの回答</div>
                                        <div className="text-xs text-gray-500 mb-2">
                                            回答日時: {details.response.timestamp ? new Date(details.response.timestamp).toLocaleString('ja-JP') : '-'} / {details.scoreName ? details.scoreName : 'スコア'}: {details.response.scoreFormatted ? details.response.scoreFormatted : (details.response.score !== null && typeof details.response.score !== 'undefined' ? String(details.response.score) + (details.scoreUnit ? (' ' + details.scoreUnit) : '') : '-')}
                                        </div>
                                        <div className="overflow-auto max-h-64 border rounded">
                                            <table className="min-w-full text-sm">
                                                <thead className="bg-gray-100">
                                                    <tr>
                                                        <th className="px-2 py-1 text-left">項目</th>
                                                        <th className="px-2 py-1 text-left">回答</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    {(details.headers && details.headers.length ? details.headers : Object.keys(details.response.answers || {})).map((h, idx) => {
                                                            const key = h || `col${idx+1}`;
                                                            const val = details.response.answers ? details.response.answers[key] : '';
                                                            const scoreKeys = ['スコア','Score','合計','点数'];
                                                            const isScore = scoreKeys.includes((h || '').toString());
                                                            const displayKey = (isScore && details.scoreName) ? details.scoreName : key;
                                                            let displayVal = '';
                                                            if (isScore) {
                                                                if (details.response && (details.response.scoreFormatted || (details.response.score !== null && typeof details.response.score !== 'undefined'))) {
                                                                    displayVal = details.response.scoreFormatted ? details.response.scoreFormatted : String(details.response.score) + (details.scoreUnit ? (' ' + details.scoreUnit) : '');
                                                                } else {
                                                                    displayVal = val === null || typeof val === 'undefined' ? '' : String(val);
                                                                }
                                                            } else {
                                                                displayVal = val === null || typeof val === 'undefined' ? '' : String(val);
                                                            }
                                                            return (
                                                            <tr key={idx} className={`${idx%2===0?'bg-white':'bg-gray-50'}`}>
                                                                <td className="px-2 py-1 border-t">{displayKey}</td>
                                                                <td className="px-2 py-1 border-t">{displayVal}</td>
                                                            </tr>
                                                            );
                                                        })}
                                                </tbody>
                                            </table>
                                        </div>
                                    </div>
                                )}

                                {details && details.viewType === 'collection' && details.historyEntries && details.historyEntries.length > 0 && (
                                    <div>
                                        <div className="font-medium mb-2">履歴</div>
                                        <div className="overflow-auto max-h-48 border rounded">
                                            <table className="min-w-full text-sm">
                                                <thead className="bg-gray-100">
                                                    <tr>
                                                        <th className="px-2 py-1 text-left">日時</th>
                                                        <th className="px-2 py-1 text-left">担当者</th>
                                                        <th className="px-2 py-1 text-left">種別</th>
                                                        <th className="px-2 py-1 text-right">金額</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    {details.historyEntries.map((h, i) => {
                                                        const v = Number(h.amount || 0);
                                                        const abs = Math.abs(v).toLocaleString();
                                                        const amt = v < 0 ? '△' + abs : abs;
                                                        const handlerLabel = h.handlerName || h.handler || '-';
                                                        return (
                                                            <tr key={i} className={`${i%2===0?'bg-white':'bg-gray-50'}`}>
                                                                <td className="px-2 py-1 border-t">{h.timestamp ? new Date(h.timestamp).toLocaleString('ja-JP') : '-'}</td>
                                                                <td className="px-2 py-1 border-t">{handlerLabel}</td>
                                                                <td className="px-2 py-1 border-t">{h.type || '-'}</td>
                                                                <td className="px-2 py-1 border-t text-right">{amt}円</td>
                                                            </tr>
                                                        );
                                                    })}
                                                </tbody>
                                            </table>
                                        </div>
                                    </div>
                                )}

                                {details && details.latestPerEmail && (
                                    <div>
                                        <div className="font-medium mb-2">最新の回答（メール毎）</div>
                                        <div className="overflow-auto max-h-48 border rounded">
                                            <table className="min-w-full text-sm"><thead className="bg-gray-100"><tr><th className="px-2 py-1 text-left">メール</th><th className="px-2 py-1 text-left">回答日時</th><th className="px-2 py-1 text-left">{details && details.scoreName ? details.scoreName : 'スコア'}</th></tr></thead>
                                            <tbody>{details.latestPerEmail.map((r,ri)=>(<tr key={ri} className={`${ri%2===0?'bg-white':'bg-gray-50'}`}><td className="px-2 py-1 border-t">{r.email||'-'}</td><td className="px-2 py-1 border-t">{r.timestamp? new Date(r.timestamp).toLocaleString('ja-JP') : '-'}</td><td className="px-2 py-1 border-t">{r.scoreFormatted ? r.scoreFormatted : (r.score !== null && typeof r.score !== 'undefined' ? String(r.score) + (details && details.scoreUnit ? (' ' + details.scoreUnit) : '') : '-')}</td></tr>))}</tbody>
                                            </table>
                                        </div>
                                    </div>
                                )}

                                {details && details.scoreStats && (
                                    <div>
                                        <div className="font-medium mb-2">スコア配分</div>
                                        <div className="text-sm text-gray-600">件数: {details.scoreStats.count} / 平均: {details.scoreStats.avgFormatted || (Math.round((details.scoreStats.avg||0)*100)/100 + (details.scoreUnit ? (' ' + details.scoreUnit) : ''))} / 最小: {details.scoreStats.minFormatted || details.scoreStats.min} / 最大: {details.scoreStats.maxFormatted || details.scoreStats.max}</div>
                                        <div className="mt-2 text-sm">
                                            {Object.keys(details.scoreStats.distributionFormatted || {}).sort((a,b)=>Number(a.replace(/[^0-9]/g,'')) - Number(b.replace(/[^0-9]/g,''))).reverse().map(k=>(<div key={k} className="flex justify-between text-xs py-1 border-b"><div>{k}</div><div>{details.scoreStats.distributionFormatted[k]} 件</div></div>))}
                                        </div>
                                    </div>
                                )}
                                {details && !details.error && !details.response && !details.latestPerEmail && !details.scoreStats && (
                                    <div className="text-gray-500">表示できる詳細がありません。</div>
                                )}
                            </div>
                            </>
                        )}
                    </div>
                </div>
            );
        }
