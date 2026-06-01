import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'RYUTA 日報スタジオ',
  description: 'GAS × Vercel — 業務日報の自動生成と Gmail 下書き',
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="ja">
      <body>{children}</body>
    </html>
  );
}
