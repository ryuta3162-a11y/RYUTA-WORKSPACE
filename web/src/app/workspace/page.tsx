const workspaceUrl = process.env.NEXT_PUBLIC_GAS_WEB_APP_URL ?? '';

export default function WorkspacePage() {
  if (!workspaceUrl) {
    return (
      <main style={{ padding: '2rem', color: '#e4e4e7', fontFamily: 'sans-serif' }}>
        <h1 style={{ fontSize: 20 }}>Workspace を表示できません</h1>
        <p style={{ color: '#a1a1aa', lineHeight: 1.6 }}>
          Vercel の環境変数に <code>NEXT_PUBLIC_GAS_WEB_APP_URL</code> を、GAS の Web アプリ{' '}
          <code>/exec</code> URL と同じ値で追加してから再デプロイしてください。
        </p>
        <p>
          <a href="/" style={{ color: '#93c5fd' }}>
            ← 日報スタジオへ
          </a>
        </p>
      </main>
    );
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', background: '#030308' }}>
      <header
        style={{
          flex: 'none',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          gap: 12,
          padding: '8px 12px',
          borderBottom: '1px solid rgba(255,255,255,0.08)',
          fontSize: 12,
          color: '#a1a1aa',
        }}
      >
        <span>RYUTA Workspace（GAS 本体）</span>
        <a href="/" style={{ color: '#e4e4e7', textDecoration: 'none' }}>
          日報スタジオ →
        </a>
      </header>
      <iframe
        title="RYUTA Workspace"
        src={workspaceUrl}
        style={{ flex: 1, width: '100%', border: 0 }}
      />
    </div>
  );
}
