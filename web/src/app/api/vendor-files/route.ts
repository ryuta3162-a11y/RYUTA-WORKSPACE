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
  const threadId = new URL(req.url).searchParams.get('threadId') || '';
  if (!threadId) {
    return NextResponse.json({ ok: false, message: 'threadId がありません' }, { status: 400, headers: cors });
  }
  try {
    const data = await gasGet<{ files?: unknown[] }>('vendorFiles', { threadId });
    const status = data.ok ? 200 : 502;
    return NextResponse.json(data, { status, headers: cors });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ ok: false, message }, { status: 502, headers: cors });
  }
}
