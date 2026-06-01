import type { DayContext } from './gas';

export type GeneratedReport = {
  gyomu: string[];
  kansou: string;
};

function buildPrompt(ctx: DayContext, notes?: string): string {
  const events =
    ctx.calendarEvents.length > 0
      ? ctx.calendarEvents
          .map((e) => `- ${e.start}-${e.end} ${e.title}${e.isAllDay ? ' (終日)' : ''}`)
          .join('\n')
      : '（予定なし）';

  const done =
    ctx.workspace.done?.length > 0
      ? ctx.workspace.done.map((t) => `- ${t}`).join('\n')
      : '（なし）';

  const kansouDraft = ctx.workspace.kansou?.trim() || '（なし）';
  const extra = notes?.trim() ? `\n【ユーザーからの追記】\n${notes.trim()}` : '';

  return `あなたはフィットネスクラブ（JOYFIT24経堂）スタッフの業務日報アシスタントです。
以下の情報だけを根拠に、ビジネス敬語の日報用テキストを作成してください。
推測で事実を足さないでください。不明な点は書かないでください。

【今日の日付】${ctx.date}

【Googleカレンダー】
${events}

【Workspace 完了タスク】
${done}

【Workspace 所感メモ】
${kansouDraft}
${extra}

出力は次の JSON のみ（説明・マークダウン不要）:
{
  "gyomu": ["箇条書き1", "箇条書き2", ...],
  "kansou": "所感本文（200〜400文字、敬語）"
}

gyomu は3〜8項目、各項目は短い体言止めまたは動作の短文。重複は統合してください。`;
}

export async function generateReportWithGemini(
  ctx: DayContext,
  notes?: string
): Promise<GeneratedReport> {
  const key = process.env.GEMINI_API_KEY;
  if (!key) throw new Error('GEMINI_API_KEY が未設定です');

  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' +
    encodeURIComponent(key);

  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      contents: [{ parts: [{ text: buildPrompt(ctx, notes) }] }],
      generationConfig: { temperature: 0.35, maxOutputTokens: 1024 },
    }),
  });

  const json = (await res.json()) as {
    error?: { message?: string };
    candidates?: Array<{
      content?: { parts?: Array<{ text?: string }> };
    }>;
  };

  if (!res.ok) {
    throw new Error(json.error?.message ?? `Gemini API error (${res.status})`);
  }

  const raw =
    json.candidates?.[0]?.content?.parts?.[0]?.text?.trim() ?? '';
  const jsonMatch = raw.match(/\{[\s\S]*\}/);
  if (!jsonMatch) throw new Error('AI の応答を JSON として読み取れませんでした');

  const parsed = JSON.parse(jsonMatch[0]) as GeneratedReport;
  if (!Array.isArray(parsed.gyomu)) parsed.gyomu = [];
  parsed.gyomu = parsed.gyomu.map((s) => String(s).trim()).filter(Boolean);
  parsed.kansou = String(parsed.kansou ?? '').trim();
  return parsed;
}
