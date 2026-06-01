import { redirect } from 'next/navigation';

const workspaceUrl =
  process.env.GAS_WEB_APP_URL ||
  process.env.NEXT_PUBLIC_GAS_WEB_APP_URL ||
  '';

/**
 * Vercel トップ → GAS Workspace へリダイレクト
 * （GAS は iframe 埋め込みを拒否するため、全画面表示はリダイレクトが確実）
 */
export default function HomePage() {
  if (workspaceUrl) {
    redirect(workspaceUrl);
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
        Vercel の Environment Variables に次のどちらか（同じ <code>/exec</code> URL）を入れて再デプロイしてください。
      </p>
      <ul style={{ color: '#d4d4d8', lineHeight: 1.8 }}>
        <li>
          <strong>GAS_WEB_APP_URL</strong> … 推奨（サーバー用・リダイレクトに使用）
        </li>
        <li>
          <strong>NEXT_PUBLIC_GAS_WEB_APP_URL</strong> … 同上（ブラウザ用・どちらか一方で可）
        </li>
        <li>
          <strong>GAS_API_TOKEN</strong> … <code>/nippo</code> 用（WS_API_TOKEN と同じ）
        </li>
      </ul>
      <p style={{ marginTop: 1.5rem }}>
        <a href="/nippo" style={{ color: '#93c5fd' }}>
          日報スタジオ（/nippo）だけ開く
        </a>
      </p>
    </main>
  );
}
