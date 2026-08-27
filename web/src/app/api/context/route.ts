import { NextResponse } from 'next/server';
import { gasGet, type DayContext } from '@/lib/gas';

export async function GET(req: Request) {
  const { searchParams } = new URL(req.url);
  const date = String(searchParams.get('date') || '').trim();
  const data = await gasGet<DayContext>('dayContext', date ? { date } : undefined);
  if (!data.ok) {
    return NextResponse.json(data, { status: 502 });
  }
  return NextResponse.json(data);
}
