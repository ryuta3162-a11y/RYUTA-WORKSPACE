import { redirect } from 'next/navigation';

const workspaceUrl =
  process.env.NEXT_PUBLIC_GAS_WEB_APP_URL || process.env.GAS_WEB_APP_URL || '';

/** Vercel のトップ = いつもの GAS Workspace（4分割） */
export default function HomePage() {
  if (!workspaceUrl) {
    return (
      <main
        style={{
          padding: '2rem',
          color: '#e4e4e7',
          fontFamily: 'Segoe UI, Meiryo, sans-serif',
          background: '#0a0a0f',
          minHeight: '100vh',
        }}
      >
        <h1 style={{ fontSize: 20 }}>Workspace URL が未設定です</h1>
        <p style={{ color: '#a1a1aa', lineHeight: 1.7, maxWidth: 520 }}>
          Vercel の環境変数に <strong>NEXT_PUBLIC_GAS_WEB_APP_URL</strong> を追加してください。
          値は GAS Web アプリの <code>/exec</code> URL（<code>GAS_WEB_APP_URL</code> と同じ）です。
        </p>
        <p>
          <a href="/nippo" style={{ color: '#93c5fd' }}>
            日報スタジオ（/nippo）だけ開く
          </a>
        </p>
      </main>
    );
  }

  return (
    <div style={{ position: 'fixed', inset: 0, margin: 0, background: '#030308', overflow: 'hidden' }}>
      <iframe
        title="RYUTA Workspace"
        src={workspaceUrl}
        style={{ width: '100%', height: '100%', border: 0, display: 'block' }}
      />
    </div>
  );
}
