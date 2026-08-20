import { NextResponse } from 'next/server';
import { generateGeminiText } from '@/lib/gemini';

export const maxDuration = 30;

function tidyKansou(raw: string) {
  let s = String(raw || '').trim();
  s = s.replace(/^["「『]|["」』]$/g, '').trim();
  s = s.replace(/[（(]\s*\d+\s*(?:文字|字)\s*[）)]\s*$/g, '').trim();
  s = s.replace(/[（(]\s*\d+\s*(?:文字|字)\s*[）)]/g, '');
  s = s.replace(/```[\s\S]*?```/g, (block) =>
    block.replace(/^```(?:\w+)?\n?/, '').replace(/\n?```$/, '')
  );
  s = s.replace(/^(所感|本文)[：:\s]*/gm, '');
  s = s.replace(/。\s*/g, '。\n');
  s = s.replace(/\n{3,}/g, '\n\n').trim();
  return s;
}

export async function POST(req: Request) {
  try {
    let body: { gyomu?: string; kansou?: string } = {};
    try {
      body = (await req.json()) as { gyomu?: string; kansou?: string };
    } catch (e) {
      return NextResponse.json(
        { ok: false, message: 'リクエスト形式が不正です（JSON を送れませんでした）。' },
        { status: 400 }
      );
    }

    const gyomu = String(body.gyomu || '').trim();
    const kansou = String(body.kansou || '').trim();
    if (!kansou) {
      return NextResponse.json({ ok: false, message: '所感を入力してください。' }, { status: 400 });
    }

    const key = process.env.GEMINI_API_KEY || '';
    if (!key) {
      return NextResponse.json(
        {
          ok: false,
          message:
            'GEMINI_API_KEY が未設定です。Google AI Studio で無料キーを作り、Vercel の環境変数に入れてください。',
        },
        { status: 503 }
      );
    }

    const prompt =
      'あなたはJOYFIT24経堂の店長・日下竜太の日報校閲係です。' +
      '口語や音声入力のままの「所感」を、日報向けの正しい日本語に整えてください。' +
      '事実・固有名詞・数字は変えない。誇張しない。300字以内。' +
      '字数（〇〇字）は書かない。HTMLタグは使わない。各文の「。」の直後で改行する。' +
      '出力は所感の本文のみ（見出し・説明・引用符は不要）。\n\n' +
      '【今日の業務内容】\n' +
      (gyomu || '（なし）') +
      '\n\n【所感の下書き】\n' +
      kansou.slice(0, 4000);

    const result = await generateGeminiText({ apiKey: key, prompt });
    if (!result.ok) {
      return NextResponse.json(
        { ok: false, message: result.message },
        { status: result.status || 502 }
      );
    }

    return NextResponse.json({ ok: true, text: tidyKansou(result.text) });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ ok: false, message }, { status: 500 });
  }
}
