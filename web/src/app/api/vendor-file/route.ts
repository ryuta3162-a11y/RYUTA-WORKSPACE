import { NextResponse } from 'next/server';
import { gasGet } from '@/lib/gas';

const cors = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export const maxDuration = 30;

export async function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: cors });
}

export async function GET(req: Request) {
  const params = new URL(req.url).searchParams;
  const threadId = params.get('threadId') || '';
  const messageId = params.get('messageId') || '';
  const index = params.get('index') || '0';
  if (!threadId && !messageId) {
    return NextResponse.json({ ok: false, message: 'messageId がありません' }, { status: 400, headers: cors });
  }
  try {
    const data = await gasGet<{ name?: string; type?: string; data?: string }>(
      'vendorFile',
      { threadId, messageId, index }
    );
    const status = data.ok ? 200 : 502;
    return NextResponse.json(data, { status, headers: cors });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ ok: false, message }, { status: 502, headers: cors });
  }
}
