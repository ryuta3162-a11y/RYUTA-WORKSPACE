import { NextResponse } from 'next/server';
import { gasPost } from '@/lib/gas';

type SaveBody = {
  active?: string[];
  done?: string[];
  kansou?: string;
};

export async function POST(req: Request) {
  try {
    const body = (await req.json()) as SaveBody;
    const result = await gasPost({
      api: 'saveTasks',
      active: Array.isArray(body.active) ? body.active : [],
      done: Array.isArray(body.done) ? body.done : [],
      kansou: String(body.kansou ?? ''),
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
