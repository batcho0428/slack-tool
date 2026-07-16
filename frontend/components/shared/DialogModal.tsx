// @ts-nocheck
'use client';
/* eslint-disable */

        export default function DialogModal({ isOpen, type, message, onOk, onCancel }) {
            if (!isOpen) return null;
            return (
                <div className="fixed inset-0 flex items-center justify-center modal-overlay p-4 z-[200]">
                    <div className="bg-white rounded-xl shadow-2xl w-full max-w-sm p-6 flex flex-col animate-fade-in">
                        {type === 'loading' ? (
                            <>
                                <div className="flex flex-col items-center justify-center py-4">
                                    <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mb-4"></div>
                                    <p className="text-gray-600 text-sm">{message}</p>
                                </div>
                            </>
                        ) : (
                            <>
                                <h3 className="text-lg font-bold mb-4 text-gray-800">{type === 'confirm' ? '確認' : '通知'}</h3>
                                <p className="text-gray-600 mb-6 whitespace-pre-wrap text-sm">{message}</p>
                                <div className="flex space-x-3 justify-end">
                                    {type === 'confirm' && (
                                        <button onClick={onCancel} className="px-4 py-2 rounded-lg text-gray-600 bg-gray-100 hover:bg-gray-200 font-medium text-sm transition-colors">キャンセル</button>
                                    )}
                                    <button onClick={onOk} className="px-4 py-2 rounded-lg text-white bg-blue-600 hover:bg-blue-700 font-bold text-sm shadow-md transition-transform active:scale-95">OK</button>
                                </div>
                            </>
                        )}
                    </div>
                </div>
            );
        }
