const workspaceUrl =
  process.env.GAS_WEB_APP_URL?.trim() ||
  process.env.NEXT_PUBLIC_GAS_WEB_APP_URL?.trim() ||
  '';

/**
 * GAS_WEB_APP_URL 設定時は next.config.mjs の redirects で GAS へ転送。
 * 未設定時のみこのページが表示される。
 */
export default function HomePage() {
  if (workspaceUrl) {
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
        <p>
          <a href={workspaceUrl} style={{ color: '#93c5fd', fontSize: 16 }}>
            RYUTA Workspace（GAS）を開く
          </a>
        </p>
      </main>
    );
  }

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
      <p style={{ color: '#a1a1aa', lineHeight: 1.7, maxWidth: 560 }}>
        Vercel → Settings → Environment Variables に{' '}
        <strong>GAS_WEB_APP_URL</strong>（GAS の <code>/exec</code> URL）を入れてから Redeploy
        してください。
      </p>
      <p style={{ marginTop: '1rem' }}>
        <a href="/nippo" style={{ color: '#93c5fd' }}>
          日報スタジオ（/nippo）
        </a>
      </p>
    </main>
  );
}
