import { NextResponse } from 'next/server';
import { generateGeminiText } from '@/lib/gemini';

const cors = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export const maxDuration = 30;

export async function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: cors });
}

function templateDraft(company: string, subject: string, snippet: string) {
  const hint = snippet ? snippet.slice(0, 180).replace(/\s+/g, ' ') : '';
  return [
    'お世話になっております。',
    'JOYFIT24経堂の日下です。',
    '',
    `「${subject || 'ご連絡'}」について、確認いたしました。`,
    hint ? `（先方の内容：${hint}）` : '',
    '',
    'こちらの状況を確認のうえ、対応を進めます。',
    '追加で必要な資料やご都合があれば、ご連絡ください。',
    '',
    '何卒よろしくお願いいたします。',
    '日下 竜太',
    'JOYFIT24経堂',
    company ? `（案件：${company}）` : '',
  ]
    .filter((line, i, arr) => !(line === '' && arr[i - 1] === ''))
    .join('\n')
    .trim();
}

export async function POST(req: Request) {
  try {
    const body = (await req.json()) as {
      company?: string;
      subject?: string;
      snippet?: string;
    };
    const company = String(body.company || '');
    const subject = String(body.subject || '');
    const snippet = String(body.snippet || '');
    const fallback = templateDraft(company, subject, snippet);
    const key = process.env.GEMINI_API_KEY || '';
    if (!key) {
      return NextResponse.json({ ok: true, draft: fallback, source: 'template' }, { headers: cors });
    }
    const prompt =
      'あなたはJOYFIT24経堂の日下竜太の返信下書き係です。業者への短いビジネスメールを日本語で作ってください。' +
      '署名は「日下 竜太 / JOYFIT24経堂」。事実は変えない。出力は本文のみ。\n\n会社: ' +
      company +
      '\n件名: ' +
      subject +
      '\n相手の文面:\n' +
      snippet.slice(0, 1200);
    const result = await generateGeminiText({ apiKey: key, prompt });
    if (!result.ok) {
      return NextResponse.json({ ok: true, draft: fallback, source: 'template' }, { headers: cors });
    }
    return NextResponse.json(
      { ok: true, draft: result.text || fallback, source: result.text ? 'ai' : 'template' },
      { headers: cors }
    );
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ ok: false, message }, { status: 500, headers: cors });
  }
}
