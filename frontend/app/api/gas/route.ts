import { NextRequest, NextResponse } from 'next/server';

export async function POST(req: NextRequest) {
  try {
    const gasUrl = process.env.GAS_WEB_APP_URL;
    const apiSecret = process.env.FRONTEND_API_SHARED_SECRET;
    if (!gasUrl) {
      return NextResponse.json({ ok: false, error: 'GAS_WEB_APP_URL is not set' }, { status: 500 });
    }
    if (!apiSecret) {
      return NextResponse.json(
        { ok: false, error: 'FRONTEND_API_SHARED_SECRET is not set' },
        { status: 500 }
      );
    }

    const body = await req.json();
    const requestBody = {
      ...(typeof body === 'object' && body !== null ? body : {}),
      authToken: apiSecret
    };

    const res = await fetch(gasUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(requestBody),
      cache: 'no-store'
    });

    const text = await res.text();
    let json: unknown;
    try {
      json = JSON.parse(text);
    } catch (e) {
      const snippet = String(text || '').slice(0, 200);
      const looksLikeHtml = /^\s*</.test(String(text || ''));
      const hint = looksLikeHtml
        ? 'GAS Web App がHTMLを返しています。デプロイのアクセス権を「全員」にし、GAS_WEB_APP_URL が最新デプロイURLか確認してください。'
        : 'Invalid response from GAS';
      return NextResponse.json({ ok: false, error: hint, raw: snippet }, { status: 502 });
    }

    return NextResponse.json(json);
  } catch (e) {
    return NextResponse.json(
      { ok: false, error: e instanceof Error ? e.message : String(e) },
      { status: 500 }
    );
  }
}
