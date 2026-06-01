export default function NippoLayout({ children }: { children: React.ReactNode }) {
  return <div style={{ minHeight: '100vh', overflow: 'auto', background: '#0a0a0f' }}>{children}</div>;
}
