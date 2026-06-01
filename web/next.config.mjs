/** @type {import('next').NextConfig} */
const gasUrl = (
  process.env.GAS_WEB_APP_URL ||
  process.env.NEXT_PUBLIC_GAS_WEB_APP_URL ||
  ''
).trim();

const nextConfig = {
  eslint: { ignoreDuringBuilds: true },
  typescript: { ignoreDuringBuilds: true },
  async redirects() {
    if (!gasUrl) return [];
    return [{ source: '/', destination: gasUrl, permanent: false }];
  },
};

export default nextConfig;
