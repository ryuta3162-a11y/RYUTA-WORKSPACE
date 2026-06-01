import { NextResponse } from 'next/server';
import { gasGet, type DayContext } from '@/lib/gas';
import { generateReportWithGemini } from '@/lib/gemini';

export async function POST(req: Request) {
  try {
    const body = (await req.json().catch(() => ({}))) as { notes?: string };
    const ctx = await gasGet<DayContext>('dayContext');
    if (!ctx.ok) {
      return NextResponse.json(ctx, { status: 502 });
    }
    const generated = await generateReportWithGemini(ctx, body.notes);
    return NextResponse.json({ ok: true, generated, contextDate: ctx.date });
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e);
    return NextResponse.json({ ok: false, message }, { status: 500 });
  }
}
