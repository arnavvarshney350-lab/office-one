import { expect, test, type Page } from '@playwright/test'

/** Collects everything that would show up as a problem in a user's DevTools console. */
function watch(page: Page) {
  const problems: string[] = []
  page.on('console', m => { if (m.type() === 'error' || (m.type() === 'warning' && /Content Security Policy/i.test(m.text()))) problems.push(`${m.type()}: ${m.text()}`) })
  page.on('pageerror', e => problems.push(`pageerror: ${e.message}`))
  page.on('requestfailed', r => { if (!/fonts\.(googleapis|gstatic)\.com/.test(r.url())) problems.push(`requestfailed: ${r.url()} ${r.failure()?.errorText}`) })
  return problems
}

test('production bundle: CSP is enforced and every app works without console errors', async ({ page }, testInfo) => {
  const problems = watch(page)
  await page.goto('/')
  // the policy really is on the page
  await expect(page.locator('meta[http-equiv="Content-Security-Policy"]')).toHaveCount(1)
  await expect(page.getByRole('heading', { level: 1 })).toBeVisible()

  // Word: type, format, table, image, export both formats
  await page.getByRole('button', { name: 'New document' }).click()
  const body = page.getByLabel('Document body')
  await body.click(); await page.keyboard.type('Production check'); await page.keyboard.press('ControlOrMeta+a'); await page.keyboard.press('ControlOrMeta+b')
  await page.getByRole('tab', { name: 'Insert' }).click()
  const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNQTX79HwAElwJzVt2vfAAAAABJRU5ErkJggg==', 'base64')
  await page.getByTestId('word-image-input').setInputFiles({ name: 'dot.png', mimeType: 'image/png', buffer: png })
  await expect(body.locator('img')).toHaveCount(1)
  await page.getByRole('button', { name: 'File' }).click()
  const [docx] = await Promise.all([page.waitForEvent('download'), page.getByRole('menuitem', { name: 'Export as DOCX' }).click()])
  expect(docx.suggestedFilename()).toMatch(/\.docx$/)
  await page.getByRole('button', { name: 'File' }).click()
  const [pdf] = await Promise.all([page.waitForEvent('download'), page.getByRole('menuitem', { name: 'Export as PDF' }).click()])
  expect(pdf.suggestedFilename()).toMatch(/\.pdf$/)

  // Excel: formula + chart + xlsx
  await page.goto('/excel'); await page.getByRole('button', { name: 'New blank workbook' }).click()
  const cell = (r: number, c: number) => page.locator(`td[data-r="${r}"][data-c="${c}"]`)
  for (const [r, v] of [[0, '10'], [1, '20'], [2, '=SUM(A1:A2)']] as [number, string][]) { await cell(r, 0).click(); await page.keyboard.type(v); await page.keyboard.press('Enter') }
  await expect(cell(2, 0)).toHaveText('30')
  await cell(0, 2).click(); await page.keyboard.type('=TEXT(TODAY(),"yyyy")'); await page.keyboard.press('Enter')
  await expect(cell(0, 2)).toHaveText(String(new Date().getFullYear()))
  await page.getByRole('tab', { name: 'Insert' }).click(); await page.getByRole('button', { name: 'Chart' }).click()
  await page.getByLabel('Chart data range').fill('A1:A2'); await page.getByRole('dialog', { name: 'Insert chart' }).getByRole('button', { name: 'Insert' }).click()
  await expect(page.getByTestId('chart')).toHaveCount(1)
  await page.getByRole('button', { name: 'File' }).click()
  const [xlsx] = await Promise.all([page.waitForEvent('download'), page.getByRole('menuitem', { name: 'Export as XLSX' }).click()])
  expect(xlsx.suggestedFilename()).toMatch(/\.xlsx$/)

  // PowerPoint: Google font, transition, present, both exports
  await page.goto('/powerpoint'); await page.getByRole('button', { name: 'New blank presentation' }).click()
  await page.getByRole('button', { name: 'New slide' }).click(); await page.getByRole('menuitem', { name: 'Title and content' }).click()
  await page.getByTestId('stage').locator('[data-obj]').first().click({ position: { x: 4, y: 4 } })
  await page.getByRole('button', { name: 'Font family' }).first().click(); await page.getByLabel('Search fonts').fill('lobster')
  await page.getByRole('listbox', { name: 'Fonts' }).getByRole('option', { name: 'Lobster', exact: true }).click()
  await expect(page.getByTestId('stage').locator('.slide-text').first()).toHaveCSS('font-family', /Lobster/)
  await page.getByRole('tab', { name: 'Transitions' }).click(); await page.getByRole('radio', { name: 'Fade', exact: true }).click()
  await page.getByRole('button', { name: 'Present from current slide' }).click(); await page.keyboard.press('Home'); await page.keyboard.press('ArrowRight')
  await expect(page.getByTestId('transition-stage')).toHaveAttribute('data-transition', 'fade')
  await page.keyboard.press('Escape')
  await page.getByRole('button', { name: 'File' }).click()
  const [pptx] = await Promise.all([page.waitForEvent('download'), page.getByRole('menuitem', { name: 'Export as PPTX' }).click()])
  expect(pptx.suggestedFilename()).toMatch(/\.pptx$/)
  await page.getByRole('button', { name: 'File' }).click()
  const [spdf] = await Promise.all([page.waitForEvent('download'), page.getByRole('menuitem', { name: 'Export as PDF' }).click()])
  expect(spdf.suggestedFilename()).toMatch(/\.pdf$/)
  const pdfPath = testInfo.outputPath('slides.pdf'); await spdf.saveAs(pdfPath)
  expect((await import('node:fs')).statSync(pdfPath).size).toBeGreaterThan(2000)

  // Drive lists everything, theme persists across reload, shortcuts dialog opens
  await page.goto('/files')
  await expect(page.getByTestId('file-row')).toHaveCount(3)
  await page.getByRole('button', { name: /Switch to dark mode/ }).click(); await page.reload()
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark')
  await page.keyboard.press('?'); await expect(page.getByRole('dialog', { name: 'Keyboard shortcuts' })).toBeVisible()

  expect(problems, 'console/CSP problems').toEqual([])
})

test('service worker installs and the app shell opens offline', async ({ page, context }) => {
  await page.goto('/')
  await page.evaluate(() => navigator.serviceWorker.ready.then(() => true))
  await page.reload() // let the worker take control
  await page.waitForFunction(() => !!navigator.serviceWorker.controller)
  await context.setOffline(true)
  await page.reload()
  await expect(page.getByRole('heading', { level: 1 })).toBeVisible()
  await context.setOffline(false)
})
