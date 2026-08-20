/** Gemini generateContent 用の共通ヘルパー */

export const GEMINI_MODEL_CANDIDATES = [
  'gemini-2.5-flash-lite',
  'gemini-3.1-flash-lite',
  'gemini-2.5-flash',
  'gemini-3.5-flash',
] as const;

export type GeminiGenerateResult = {
  ok: true;
  text: string;
  model: string;
} | {
  ok: false;
  message: string;
  status?: number;
};

function extractTextFromGeminiJson(json: unknown): string {
  const root = json as {
    candidates?: Array<{
      content?: { parts?: Array<{ text?: string; thought?: boolean }> };
      finishReason?: string;
    }>;
    error?: { message?: string };
  };
  if (root?.error?.message) return '';
  const parts = root?.candidates?.[0]?.content?.parts || [];
  const texts: string[] = [];
  for (let i = 0; i < parts.length; i++) {
    const p = parts[i];
    if (!p || p.thought) continue;
    const t = String(p.text || '').trim();
    if (t) texts.push(t);
  }
  return texts.join('\n').trim();
}

export async function generateGeminiText(opts: {
  apiKey: string;
  prompt: string;
  models?: readonly string[];
}): Promise<GeminiGenerateResult> {
  const models = opts.models || GEMINI_MODEL_CANDIDATES;
  let lastMessage = '添削に失敗しました。';
  let lastStatus = 502;

  for (let i = 0; i < models.length; i++) {
    const model = models[i];
    const url =
      'https://generativelanguage.googleapis.com/v1beta/models/' +
      encodeURIComponent(model) +
      ':generateContent?key=' +
      encodeURIComponent(opts.apiKey);
    let res: Response;
    try {
      res = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [{ text: opts.prompt }] }],
          generationConfig: {
            temperature: 0.4,
            maxOutputTokens: 1024,
          },
        }),
      });
    } catch (err) {
      lastMessage = err instanceof Error ? err.message : String(err);
      continue;
    }

    const raw = await res.text();
    let json: unknown = null;
    try {
      json = raw ? JSON.parse(raw) : null;
    } catch (e) {
      lastMessage = 'Gemini の応答が JSON ではありません（' + res.status + '）';
      lastStatus = res.status;
      continue;
    }

    if (!res.ok) {
      const errObj = json as { error?: { message?: string; status?: string } };
      const msg = errObj?.error?.message || ('HTTP ' + res.status);
      lastMessage = msg;
      lastStatus = res.status;
      // モデル未存在・廃止は次の候補へ
      if (
        res.status === 404 ||
        /not found|not supported|deprecated|INVALID_ARGUMENT/i.test(msg)
      ) {
        continue;
      }
      return { ok: false, message: msg, status: res.status };
    }

    const text = extractTextFromGeminiJson(json);
    if (text) return { ok: true, text, model };

    const finish = (json as { candidates?: Array<{ finishReason?: string }> })
      ?.candidates?.[0]?.finishReason;
    lastMessage =
      finish && finish !== 'STOP'
        ? '返答がブロックされました（' + finish + '）'
        : '返答を取得できませんでした。';
    lastStatus = 502;
  }

  return { ok: false, message: lastMessage, status: lastStatus };
}
