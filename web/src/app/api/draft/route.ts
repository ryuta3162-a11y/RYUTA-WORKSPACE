import { NextResponse } from 'next/server';
import { gasPost } from '@/lib/gas';

type DraftBody = {
  gyomu?: string[];
  kansou?: string;
};

export async function POST(req: Request) {
  try {
    const body = (await req.json()) as DraftBody;
    const gyomu = Array.isArray(body.gyomu) ? body.gyomu : [];
    const kansou = String(body.kansou ?? '');

    const result = await gasPost<{ subject?: string }>({
      api: 'createDailyDraft',
      active: [],
      done: gyomu,
      kansou,
    });

    if (!result.ok) {
      return NextResponse.json(result, { status: 502 });
    }
    return NextResponse.json(result);
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e);
    return NextResponse.json({ ok: false, message }, { status: 500 });
  }
}
