'use client';

import { useCallback, useEffect, useState, type CSSProperties } from 'react';

type DayContext = {
  date: string;
  calendarEvents: Array<{ title: string; start: string; end: string }>;
  workspace: { done: string[]; kansou: string };
};

export default function NippoStudioPage() {
  const [ctx, setCtx] = useState<DayContext | null>(null);
  const [notes, setNotes] = useState('');
  const [gyomuText, setGyomuText] = useState('');
  const [kansou, setKansou] = useState('');
  const [status, setStatus] = useState('');
  const [loading, setLoading] = useState(false);

  const loadContext = useCallback(async () => {
    setStatus('データ取得中…');
    const res = await fetch('/api/context');
    const data = await res.json();
    if (!data.ok) {
      setStatus(data.message ?? 'コンテキスト取得に失敗');
      return;
    }
    setCtx({
      date: data.date,
      calendarEvents: data.calendarEvents ?? [],
      workspace: data.workspace ?? { done: [], kansou: '' },
    });
    const done = (data.workspace?.done as string[]) ?? [];
    if (done.length) setGyomuText(done.map((t) => t.replace(/^[・\-]\s*/, '')).join('\n'));
    if (data.workspace?.kansou) setKansou(String(data.workspace.kansou));
    setStatus('');
  }, []);

  useEffect(() => {
    loadContext();
  }, [loadContext]);

  async function handleGenerate() {
    setLoading(true);
    setStatus('AI が業務・所感を生成中…');
    try {
      const res = await fetch('/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ notes }),
      });
      const data = await res.json();
      if (!data.ok) throw new Error(data.message ?? '生成に失敗');
      setGyomuText((data.generated.gyomu as string[]).join('\n'));
      setKansou(data.generated.kansou as string);
      setStatus('生成しました。内容を確認してから下書き作成へ。');
    } catch (e) {
      setStatus(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  }

  async function handleDraft() {
    const gyomu = gyomuText
      .split('\n')
      .map((l) => l.replace(/^[・\-•\s]+/, '').trim())
      .filter(Boolean);
    if (!gyomu.length) {
      setStatus('業務内容を 1 行以上入力してください');
      return;
    }
    if (!window.confirm('Gmail に日報下書きを作成します。よろしいですか？')) return;

    setLoading(true);
    setStatus('Gmail 下書き作成中…');
    try {
      const res = await fetch('/api/draft', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ gyomu, kansou }),
      });
      const data = await res.json();
      if (!data.ok) throw new Error(data.message ?? '下書き作成に失敗');
      setStatus(`下書きを作成しました: ${data.subject ?? ''}`);
    } catch (e) {
      setStatus(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  }

  return (
    <main style={{ maxWidth: 720, margin: '0 auto', padding: '2rem 1.25rem' }}>
      <p style={{ fontSize: 12, letterSpacing: '0.12em', color: '#71717a', margin: 0 }}>
        GAS × VERCEL × GITHUB
      </p>
      <h1 style={{ fontSize: 28, fontWeight: 700, margin: '0.25rem 0 0.5rem' }}>日報スタジオ</h1>
      <p style={{ color: '#a1a1aa', fontSize: 14, lineHeight: 1.6 }}>
        カレンダーと Workspace を GAS が集約 → Vercel で Gemini が業務・所感を生成 → GAS が経堂数値付き Gmail
        下書きを作成します。
      </p>

      {ctx && (
        <section
          style={{
            marginTop: 20,
            padding: 14,
            borderRadius: 12,
            border: '1px solid rgba(255,255,255,0.08)',
            background: 'rgba(255,255,255,0.03)',
            fontSize: 13,
          }}
        >
          <div style={{ fontWeight: 600, marginBottom: 8 }}>今日 ({ctx.date})</div>
          <div style={{ color: '#a1a1aa', marginBottom: 6 }}>カレンダー</div>
          {ctx.calendarEvents.length === 0 ? (
            <div>予定なし</div>
          ) : (
            <ul style={{ margin: 0, paddingLeft: 18 }}>
              {ctx.calendarEvents.map((ev, i) => (
                <li key={i}>
                  {ev.start}–{ev.end} {ev.title}
                </li>
              ))}
            </ul>
          )}
        </section>
      )}

      <label style={{ display: 'block', marginTop: 20, fontSize: 12, color: '#a1a1aa' }}>
        追記メモ（AI 用・任意）
        <textarea
          value={notes}
          onChange={(e) => setNotes(e.target.value)}
          rows={2}
          style={fieldStyle}
          placeholder="例: 本日は棚卸し対応が中心"
        />
      </label>

      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginTop: 12 }}>
        <button type="button" onClick={handleGenerate} disabled={loading} style={btnPrimary}>
          AI で業務・所感を生成
        </button>
        <button type="button" onClick={loadContext} disabled={loading} style={btnGhost}>
          データ再取得
        </button>
      </div>

      <label style={{ display: 'block', marginTop: 20, fontSize: 12, color: '#a1a1aa' }}>
        業務内容（1行1項目）
        <textarea value={gyomuText} onChange={(e) => setGyomuText(e.target.value)} rows={8} style={fieldStyle} />
      </label>

      <label style={{ display: 'block', marginTop: 16, fontSize: 12, color: '#a1a1aa' }}>
        所感
        <textarea value={kansou} onChange={(e) => setKansou(e.target.value)} rows={5} style={fieldStyle} />
      </label>

      <button
        type="button"
        onClick={handleDraft}
        disabled={loading}
        style={{ ...btnPrimary, marginTop: 16, width: '100%' }}
      >
        Gmail に下書き作成（経堂数値は GAS）
      </button>

      {status && (
        <p style={{ marginTop: 16, fontSize: 13, color: '#d4d4d8', whiteSpace: 'pre-wrap' }}>{status}</p>
      )}
    </main>
  );
}

const fieldStyle: CSSProperties = {
  display: 'block',
  width: '100%',
  marginTop: 6,
  padding: '10px 12px',
  borderRadius: 10,
  border: '1px solid rgba(255,255,255,0.12)',
  background: 'rgba(0,0,0,0.35)',
  color: '#f4f4f5',
  resize: 'vertical',
};

const btnPrimary: CSSProperties = {
  padding: '10px 16px',
  borderRadius: 10,
  border: '1px solid rgba(255,255,255,0.2)',
  background: '#f4f4f5',
  color: '#111',
  fontWeight: 600,
  fontSize: 13,
};

const btnGhost: CSSProperties = {
  padding: '10px 16px',
  borderRadius: 10,
  border: '1px solid rgba(255,255,255,0.15)',
  background: 'transparent',
  color: '#e4e4e7',
  fontSize: 13,
};
