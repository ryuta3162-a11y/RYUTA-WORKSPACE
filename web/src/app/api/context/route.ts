import { NextResponse } from 'next/server';
import { gasGet, type DayContext } from '@/lib/gas';

export async function GET() {
  const data = await gasGet<DayContext>('dayContext');
  if (!data.ok) {
    return NextResponse.json(data, { status: 502 });
  }
  return NextResponse.json(data);
}
