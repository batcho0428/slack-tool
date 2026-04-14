import { NextRequest, NextResponse } from 'next/server';

const MAX_RETRIES = 3;

function looksLikeHtml(text: string): boolean {
  return /^\s*</.test(text);
}

function parseJsonSafe(text: string): { ok: true; value: unknown } | { ok: false } {
  try {
    return { ok: true, value: JSON.parse(text) };
  } catch {
    return { ok: false };
  }
}

async function delay(ms: number): Promise<void> {
  await new Promise((resolve) => setTimeout(resolve, ms));
}

async function callGasWithRetry(gasUrl: string, requestBody: Record<string, unknown>) {
  let lastText = '';
  let lastStatus = 0;

  for (let i = 0; i < MAX_RETRIES; i++) {
    const res = await fetch(gasUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Accept: 'application/json'
      },
      body: JSON.stringify(requestBody),
      cache: 'no-store',
      redirect: 'follow'
    });

    const text = await res.text();
    lastText = text;
    lastStatus = res.status;

    const parsed = parseJsonSafe(text);
    if (parsed.ok) {
      return { ok: true as const, value: parsed.value };
    }

    // 一時的なHTML応答は短時間リトライで回復することがある
    if (looksLikeHtml(text) && i < MAX_RETRIES - 1) {
      await delay(250 * (i + 1));
      continue;
    }

    break;
  }

  return { ok: false as const, text: lastText, status: lastStatus };
}

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

    const result = await callGasWithRetry(gasUrl, requestBody);
    if (!result.ok) {
      const snippet = String(result.text || '').slice(0, 200);
      const hint = looksLikeHtml(String(result.text || ''))
        ? 'GAS Web App がHTMLを返しています。アクセス権設定または一時的な応答異常の可能性があります。リトライして改善しない場合は、GASのデプロイ公開設定とGAS_WEB_APP_URLを確認してください。'
        : 'Invalid response from GAS';
      return NextResponse.json(
        { ok: false, error: hint, status: result.status, raw: snippet },
        { status: 502 }
      );
    }

    return NextResponse.json(result.value);
  } catch (e) {
    return NextResponse.json(
      { ok: false, error: e instanceof Error ? e.message : String(e) },
      { status: 500 }
    );
  }
}
