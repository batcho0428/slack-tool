// @ts-nocheck
'use client';
/* eslint-disable */

        export default function ResultModal({ result, onClose }) {
            return (
                <div className="fixed inset-0 flex items-center justify-center modal-overlay p-4 z-50">
                    <div className="bg-white rounded-xl shadow-2xl w-full max-w-md p-6 flex flex-col max-h-[80vh]">
                        <div className="text-center mb-6">
                            <h3 className="text-xl font-bold">完了</h3>
                            <div className="flex justify-center space-x-4 mt-2 text-sm">
                                <span className="text-green-600 font-bold">成功: {result.success}件</span>
                                {result.failed.length > 0 && <span className="text-red-600 font-bold">失敗: {result.failed.length}件</span>}
                            </div>
                        </div>
                        {result.failed.length > 0 && (
                            <div className="flex-1 overflow-y-auto bg-red-50 p-4 rounded-lg mb-4 text-sm border border-red-100">
                                <div className="mt-2 space-y-2">
                                    {result.failed.map((f, idx) => (
                                        <div key={idx} className="p-2 border rounded">
                                            <div className="font-medium">{f.email || f.name || ('行 ' + (f.row || idx+1))}</div>
                                            <div className="text-xs text-red-600 mt-1">{f.error || (f.message || JSON.stringify(f))}</div>
                                        </div>
                                    ))}
                                </div>
                            </div>
                        )}
                        <div className="flex justify-end mt-3">
                            <button onClick={onClose} className="px-4 py-2 bg-blue-600 text-white rounded">OK</button>
                        </div>
                    </div>
                </div>
            );
        }
