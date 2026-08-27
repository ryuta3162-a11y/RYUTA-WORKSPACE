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
  const company = params.get('company') || '';
  const email = params.get('email') || '';
  const discover = params.get('discover') === '1';
  if (!company) {
    return NextResponse.json({ ok: true, cases: [] }, { headers: cors });
  }
  try {
    const extra: Record<string, string> = { company };
    if (email) extra.email = email;
    const since = params.get('since') || '';
    if (since) extra.since = since;
    const data = await gasGet(discover ? 'vendorDiscover' : 'vendorMail', extra);
    const status = data.ok ? 200 : 502;
    return NextResponse.json(data, { status, headers: cors });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ ok: false, message }, { status: 502, headers: cors });
  }
}
