const GAS_URL = process.env.GAS_WEB_APP_URL ?? '';
const GAS_TOKEN = process.env.GAS_API_TOKEN ?? '';

export type GasJson<T> = T & { ok: boolean; message?: string };

export type DayContext = {
  ok: true;
  date: string;
  timezone: string;
  calendarEvents: Array<{
    title: string;
    start: string;
    end: string;
    isAllDay: boolean;
  }>;
  workspace: {
    ok: boolean;
    active: string[];
    done: string[];
    kansou: string;
    hasTodayRow?: boolean;
  };
};

export async function gasGet<T>(api: string): Promise<GasJson<T>> {
  if (!GAS_URL) {
    return { ok: false, message: 'GAS_WEB_APP_URL が未設定です' } as GasJson<T>;
  }
  const url = new URL(GAS_URL);
  url.searchParams.set('api', api);
  if (GAS_TOKEN) url.searchParams.set('token', GAS_TOKEN);

  const res = await fetch(url.toString(), { cache: 'no-store' });
  return (await res.json()) as GasJson<T>;
}

export async function gasPost<T>(body: Record<string, unknown>): Promise<GasJson<T>> {
  if (!GAS_URL) {
    return { ok: false, message: 'GAS_WEB_APP_URL が未設定です' } as GasJson<T>;
  }
  const payload = GAS_TOKEN ? { ...body, token: GAS_TOKEN } : body;
  const res = await fetch(GAS_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
    cache: 'no-store',
  });
  return (await res.json()) as GasJson<T>;
}
