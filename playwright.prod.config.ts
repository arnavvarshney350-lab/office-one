import { defineConfig } from '@playwright/test'

// Runs against the production build (with its Content-Security-Policy and service worker).
export default defineConfig({
  testDir: './e2e-prod',
  timeout: 90_000,
  workers: 1,
  reporter: [['list']],
  use: { baseURL: 'http://localhost:4199', channel: 'chrome', viewport: { width: 1360, height: 860 }, acceptDownloads: true },
  webServer: { command: 'npm run build && npx vite preview --port 4199 --strictPort', url: 'http://localhost:4199', reuseExistingServer: false, timeout: 180_000 },
})
