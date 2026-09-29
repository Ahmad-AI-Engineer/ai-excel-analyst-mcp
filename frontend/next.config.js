/** @type {import('next').NextConfig} */
const nextConfig = {
  output: 'standalone',
  transpilePackages: ['three', '@react-three/fiber'],
  // Keep the internal backend connection alive across chat requests. This avoids
  // intermittent socket resets while Next.js proxies /api requests to FastAPI.
  httpAgentOptions: {
    keepAlive: true,
  },
  experimental: {
    // Next.js's built-in rewrite proxy (httpxy) kills the upstream connection
    // after 30s by default, regardless of client fetch/AbortController settings.
    // The agent's /chat requests can legitimately run for several minutes
    // (multiple LLM + MCP tool calls, plus provider-overload retries), so the
    // proxy must be allowed to wait at least that long before giving up.
    proxyTimeout: 600_000, // 10 minutes
  },
  async rewrites() {
    const backend = (process.env.BACKEND_INTERNAL_URL || 'http://localhost:8003').replace(/\/$/, '')
    return [{ source: '/api/:path*', destination: `${backend}/:path*` }]
  },
}

module.exports = nextConfig
