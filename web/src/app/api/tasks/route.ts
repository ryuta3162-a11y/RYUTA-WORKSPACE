import { NextResponse } from 'next/server';
import { gasGet, gasPost } from '@/lib/gas';

export type PersonalTask = {
  id: string;
  title: string;
  bucket: 'today' | 'week' | 'month' | 'waiting' | 'followup';
  due_date?: string;
  priority?: 'high' | 'mid' | 'low';
  url?: string;
  status: 'open' | 'done';
  done_at?: string;
  created_at?: string;
  updated_at?: string;
  note?: string;
  period_key?: string;
};

type TasksBody = {
  action?: 'list' | 'upsert' | 'delete';
  task?: Partial<PersonalTask>;
  id?: string;
};

export async function GET() {
  const data = await gasGet<{ tasks: PersonalTask[]; carried?: number; date?: string }>(
    'personalTasks'
  );
  if (!data.ok) {
    return NextResponse.json(data, { status: 502 });
  }
  return NextResponse.json(data);
}

export async function POST(req: Request) {
  try {
    const body = (await req.json()) as TasksBody;
    const action = body.action || 'upsert';
    const result = await gasPost({
      api: 'personalTasks',
      action,
      task: body.task || {},
      id: body.id || '',
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
