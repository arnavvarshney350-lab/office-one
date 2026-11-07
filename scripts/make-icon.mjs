// Builds desktop/icon.icns from public/icon.svg (needs macOS: sips + iconutil, and Google Chrome via Playwright).
import { chromium } from '@playwright/test'
import { execFileSync } from 'node:child_process'
import { mkdirSync, readFileSync, rmSync } from 'node:fs'
import { fileURLToPath } from 'node:url'

const svg = readFileSync(fileURLToPath(new URL('../public/icon.svg', import.meta.url)), 'utf8')
const out = fileURLToPath(new URL('../desktop/', import.meta.url)) // not .pathname: that percent-encodes spaces and apostrophes
const set = out + 'icon.iconset'
rmSync(set, { recursive: true, force: true }); mkdirSync(set, { recursive: true })

const browser = await chromium.launch({ channel: 'chrome' })
const page = await browser.newPage({ viewport: { width: 1024, height: 1024 } })
// macOS icons sit inside a transparent margin; draw the artwork at ~82% on a transparent canvas
await page.setContent(`<body style="margin:0;background:transparent"><div style="width:1024px;height:1024px;display:flex;align-items:center;justify-content:center"><div style="width:840px;height:840px;filter:drop-shadow(0 14px 22px rgba(20,24,36,.35))">${svg.replace('<svg ', '<svg width="840" height="840" ')}</div></div></body>`)
await page.screenshot({ path: out + 'icon-1024.png', omitBackground: true })
await browser.close()

for (const s of [16, 32, 64, 128, 256, 512, 1024]) {
  execFileSync('sips', ['-z', String(s), String(s), out + 'icon-1024.png', '--out', `${set}/tmp_${s}.png`], { stdio: 'ignore' })
}
const names = { 16: ['icon_16x16'], 32: ['icon_16x16@2x', 'icon_32x32'], 64: ['icon_32x32@2x'], 128: ['icon_128x128'], 256: ['icon_128x128@2x', 'icon_256x256'], 512: ['icon_256x256@2x', 'icon_512x512'], 1024: ['icon_512x512@2x'] }
for (const [s, list] of Object.entries(names)) for (const n of list) execFileSync('cp', [`${set}/tmp_${s}.png`, `${set}/${n}.png`])
for (const s of Object.keys(names)) rmSync(`${set}/tmp_${s}.png`)
execFileSync('iconutil', ['-c', 'icns', set, '-o', out + 'icon.icns'])
rmSync(set, { recursive: true, force: true }); rmSync(out + 'icon-1024.png')
console.log('Wrote desktop/icon.icns')
