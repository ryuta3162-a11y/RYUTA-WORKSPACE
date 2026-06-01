import type { ReactNode } from 'react';

export default function NippoLayout({ children }: { children: ReactNode }) {
  return <div style={{ minHeight: '100vh', overflow: 'auto', background: '#0a0a0f' }}>{children}</div>;
}
