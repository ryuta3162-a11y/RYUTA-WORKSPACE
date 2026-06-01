import Link from 'next/link';

/** 旧パス → トップ（GAS へリダイレクト） */
export default function WorkspaceLegacyPage() {
  return (
    <main style={{ padding: '2rem', color: '#e4e4e7', background: '#0a0a0f', minHeight: '100vh' }}>
      <p>
        <Link href="/" style={{ color: '#93c5fd' }}>
          トップへ（Workspace）
        </Link>
      </p>
    </main>
  );
}
