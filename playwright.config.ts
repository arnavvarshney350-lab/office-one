import { defineConfig } from '@playwright/test'

// Uses the locally installed Google Chrome (no browser download needed).
export default defineConfig({
  testDir: './e2e',
  timeout: 60_000,
  workers: 1,
  reporter: [['list']],
  use: { baseURL: 'http://localhost:5199', channel: 'chrome', viewport: { width: 1360, height: 860 }, acceptDownloads: true },
  webServer: { command: 'npx vite --port 5199 --strictPort', url: 'http://localhost:5199', reuseExistingServer: true, timeout: 60_000 },
})
