import type { Metadata, Viewport } from 'next';
import type { ReactNode } from 'react';
import PwaRegister from '@/components/PwaRegister';
import './globals.css';

export const metadata: Metadata = {
  title: 'やるしかない',
  description: 'やるしかない — TODO・日報・カレンダー',
  applicationName: 'やるしかない',
  manifest: '/manifest.webmanifest',
  appleWebApp: {
    capable: true,
    statusBarStyle: 'black-translucent',
    title: 'やるしかない',
  },
  icons: {
    icon: '/favicon.png',
    apple: '/icons/apple-touch-icon.png',
  },
};

export const viewport: Viewport = {
  themeColor: '#000000',
  width: 'device-width',
  initialScale: 1,
};

export default function RootLayout({ children }: { children: ReactNode }) {
  return (
    <html lang="ja">
      <body>
        {children}
        <PwaRegister />
      </body>
    </html>
  );
}
