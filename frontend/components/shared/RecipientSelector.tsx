// @ts-nocheck
'use client';
/* eslint-disable */

import { useState } from 'react';

        export default function RecipientSelector({ recipients, setRecipients, onAddClick, labelText }) {
            return (
                <div className="space-y-2 h-full flex flex-col">
                    <div className="flex justify-between items-end">
                        <label className="font-bold text-gray-700 block text-sm md:text-base">
                            <i className="fas fa-users mr-1"></i> {labelText || "対象ユーザー"}
                            <span className="ml-2 bg-gray-200 text-gray-700 px-2 py-0.5 rounded-full text-xs">{recipients.length}名</span>
                        </label>
                        <button onClick={onAddClick} className="text-sm bg-gray-700 text-white px-3 py-1.5 md:px-4 md:py-1.5 rounded hover:bg-gray-900 shadow-sm active:scale-95 transition-transform">
                            <i className="fas fa-plus mr-1"></i> 追加
                        </button>
                    </div>
                    <div className="border border-gray-300 rounded-lg bg-gray-50 flex-1 overflow-y-auto p-2 min-h-[150px]">
                        {recipients.length === 0 ? (
                            <div className="h-full flex items-center justify-center text-gray-400 text-sm">ユーザーが選択されていません</div>
                        ) : (
                            <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
                                {recipients.map(r => (
                                    <div key={r.email} className="flex justify-between items-center bg-white p-2 px-3 rounded border border-gray-200 shadow-sm animate-fade-in">
                                        <div className="min-w-0 overflow-hidden">
                                            <div className="font-bold text-sm text-gray-800 truncate">{r.name}</div>
                                            <div className="text-xs text-gray-500 truncate mb-1">{r.email}</div>
                                            {(r.grade || r.field) && (
                                                <div className="text-xs text-blue-600 truncate flex items-center">
                                                    <i className="fas fa-school mr-1"></i>
                                                    {r.grade} {r.field}
                                                </div>
                                            )}
                                            <div className="text-xs text-gray-500 mt-1 whitespace-normal break-words" title={r.department}>{r.department}</div>
                                        </div>
                                        <button onClick={()=>setRecipients(recipients.filter(x=>x.email!==r.email))} className="text-gray-400 hover:text-red-500 ml-2 p-2">
                                            <i className="fas fa-times text-lg"></i>
                                        </button>
                                    </div>
                                ))}
                            </div>
                        )}
                    </div>
                </div>
            );
        }
