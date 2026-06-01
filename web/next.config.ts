import type { NextConfig } from 'next';

const gasUrl =
  process.env.GAS_WEB_APP_URL?.trim() ||
  process.env.NEXT_PUBLIC_GAS_WEB_APP_URL?.trim() ||
  '';

const nextConfig: NextConfig = {
  eslint: { ignoreDuringBuilds: true },
  typescript: { ignoreBuildErrors: false },
  async redirects() {
    if (!gasUrl) return [];
    return [
      {
        source: '/',
        destination: gasUrl,
        permanent: false,
      },
    ];
  },
};

export default nextConfig;
