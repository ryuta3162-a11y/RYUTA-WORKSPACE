import { NextResponse } from 'next/server';
import { gasGet } from '@/lib/gas';

export async function GET(req: Request) {
  const { searchParams } = new URL(req.url);
  const data = await gasGet('partnerMails', {
    emails: searchParams.get('emails') || '',
    label: searchParams.get('label') || '',
  });
  if (!data.ok) {
    return NextResponse.json(data, { status: 502 });
  }
  return NextResponse.json(data);
}
