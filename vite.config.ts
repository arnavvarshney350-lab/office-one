/// <reference types="vitest/config" />
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'
import { defineConfig, type Plugin } from 'vite'

/**
 * Production-only Content-Security-Policy. The dev server needs inline/eval scripts for hot reload, so the
 * policy is injected at build time only. Google Fonts is the sole third-party origin.
 */
const csp = (): Plugin => ({
  name: 'office-one-csp',
  apply: 'build',
  transformIndexHtml: () => [{
    tag: 'meta', injectTo: 'head-prepend',
    attrs: {
      'http-equiv': 'Content-Security-Policy',
      content: [
        "default-src 'self'", "script-src 'self'", "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com",
        "font-src 'self' https://fonts.gstatic.com data:", "img-src 'self' data: blob:", "connect-src 'self' https://fonts.googleapis.com https://fonts.gstatic.com",
        "worker-src 'self' blob:", "object-src 'none'", "base-uri 'self'", "form-action 'self'",
      ].join('; '),
    },
  }],
})

export default defineConfig({
  plugins: [react(), tailwindcss(), csp()],
  build: { target: 'es2022', sourcemap: false },
  server: { port: 5173, host: true },
  test: { environment: 'node', include: ['src/**/*.test.ts'] },
})
