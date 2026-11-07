import { expect, test, type Page } from '@playwright/test'

const MOD = process.platform === 'darwin' ? 'Meta' : 'Control'
/** Move the caret to the end of the line (Cmd+Right on macOS; the End key does not move the caret there). */
const LINE_END = process.platform === 'darwin' ? 'Meta+ArrowRight' : 'End'
test.beforeEach(async ({ page }) => { await page.goto('/'); await page.evaluate(() => localStorage.setItem('oo-theme', 'light')); await page.reload() })

const setStyle = async (page: Page, name: string) => { await page.getByRole('button', { name: 'Paragraph style' }).click(); await page.getByRole('menuitem', { name }).click() }
const fontOptions = (page: Page, name?: string) => page.getByRole('listbox', { name: 'Fonts' }).getByRole('option', name ? { name } : undefined)
const cell = (page: Page, r: number, c: number) => page.locator(`td[data-r="${r}"][data-c="${c}"]`)

test('navigation and theme switching', async ({ page }) => {
  await expect(page.getByRole('heading', { level: 1 })).toBeVisible()
  for (const [app, url] of [['Word', '/word'], ['Excel', '/excel'], ['PowerPoint', '/powerpoint'], ['Drive', '/files']] as const) {
    await page.goto('/'); await page.getByRole('navigation', { name: 'Applications' }).getByRole('link', { name: app }).click(); await expect(page).toHaveURL(new RegExp(url))
  }
  await page.goto('/')
  await page.getByRole('button', { name: /Switch to dark mode/ }).click()
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark')
  await page.reload(); await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark')
})

test('Word: type, format, table, persist, export', async ({ page }) => {
  await page.getByRole('button', { name: 'New document' }).click()
  await expect(page).toHaveURL(/\/word\//)
  const body = page.getByLabel('Document body')
  await body.click(); await page.keyboard.type('Hello Office One')
  await page.keyboard.press(`${MOD}+a`); await page.keyboard.press(`${MOD}+b`)
  await page.getByRole('button', { name: 'Italic (⌘I)' }).click()
  await page.keyboard.press('ArrowRight')
  await page.getByRole('tab', { name: 'Insert' }).click()
  await page.getByRole('button', { name: 'Table', exact: true }).click()
  await page.getByRole('button', { name: 'Insert', exact: true }).click()
  await expect(body.locator('table')).toHaveCount(1)
  await expect(page.getByText('Saved')).toBeVisible()
  await page.reload()
  await expect(page.getByLabel('Document body').locator('strong em, em strong').first()).toContainText('Hello Office One')
  await expect(page.getByLabel('Document body').locator('table')).toHaveCount(1)
  await expect(page.getByText(/\d+ words?/)).toBeVisible()
  await page.getByRole('button', { name: 'File' }).click()
  const [dl] = await Promise.all([page.waitForEvent('download'), page.getByRole('menuitem', { name: 'Export as DOCX' }).click()])
  expect(dl.suggestedFilename()).toMatch(/\.docx$/)
  await page.getByRole('button', { name: 'File' }).click()
  const [pdf] = await Promise.all([page.waitForEvent('download'), page.getByRole('menuitem', { name: 'Export as PDF' }).click()])
  expect(pdf.suggestedFilename()).toMatch(/\.pdf$/)
})

test('Excel: values, formulas, recalculation, sheets, persistence, export', async ({ page }) => {
  await page.getByRole('button', { name: 'New spreadsheet' }).click()
  await expect(page).toHaveURL(/\/excel\//)
  await cell(page, 0, 0).click()
  await page.keyboard.type('10'); await page.keyboard.press('Enter')
  await page.keyboard.type('20'); await page.keyboard.press('Enter')
  await page.keyboard.type('=SUM(A1:A2)'); await page.keyboard.press('Enter')
  await expect(cell(page, 2, 0)).toHaveText('30')
  await cell(page, 0, 0).click(); await page.keyboard.type('15'); await page.keyboard.press('Enter')
  await expect(cell(page, 2, 0)).toHaveText('35')
  await cell(page, 0, 1).click(); await page.keyboard.type('=IF(A1>10,"big","small")'); await page.keyboard.press('Enter')
  await expect(cell(page, 0, 1)).toHaveText('big')
  await cell(page, 1, 1).click(); await page.keyboard.type('=1/0'); await page.keyboard.press('Enter')
  await expect(cell(page, 1, 1)).toHaveText('#DIV/0!')
  await page.getByRole('button', { name: 'Add worksheet' }).click()
  await expect(page.getByRole('button', { name: 'Sheet Sheet2' })).toBeVisible()
  await expect(page.getByText('Saved')).toBeVisible()
  await page.reload()
  await expect(page.getByRole('button', { name: 'Sheet Sheet2' })).toBeVisible()
  await page.getByRole('button', { name: 'Sheet Sheet1' }).click()
  await expect(cell(page, 2, 0)).toHaveText('35')
  await page.getByRole('button', { name: 'File' }).click()
  const [dl] = await Promise.all([page.waitForEvent('download'), page.getByRole('menuitem', { name: 'Export as XLSX' }).click()])
  expect(dl.suggestedFilename()).toMatch(/\.xlsx$/)
})

test('PowerPoint: slides, shapes, present, export', async ({ page }) => {
  await page.goto('/powerpoint'); await page.getByRole('button', { name: 'Template Midnight' }).click()
  await expect(page).toHaveURL(/\/powerpoint\//)
  await page.getByRole('button', { name: 'New slide' }).click(); await page.getByRole('menuitem', { name: 'Title and content' }).click()
  await expect(page.getByTestId('thumb')).toHaveCount(2)
  await page.getByRole('tab', { name: 'Insert' }).click(); await page.getByRole('button', { name: 'Ellipse' }).click()
  await expect(page.getByLabel('Resize se')).toBeVisible()
  const box = (await page.locator('[data-testid=stage] [data-obj]:has(ellipse)').boundingBox())!
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2); await page.mouse.down(); await page.mouse.move(box.x + box.width / 2 + 80, box.y + box.height / 2 + 40, { steps: 6 }); await page.mouse.up()
  await page.getByRole('tab', { name: 'Home' }).click()
  await page.getByRole('button', { name: 'Move slide up' }).click()
  await expect(page.getByText('Slide 1 of 2')).toBeVisible()
  await page.getByRole('button', { name: 'Present from current slide' }).click()
  await expect(page.getByTestId('presenter')).toBeVisible()
  await page.keyboard.press('ArrowRight'); await expect(page.getByTestId('present-counter')).toHaveText('2 / 2')
  await page.keyboard.press('Escape'); await expect(page.getByTestId('presenter')).toHaveCount(0)
  await page.getByRole('button', { name: 'File' }).click()
  const [dl] = await Promise.all([page.waitForEvent('download'), page.getByRole('menuitem', { name: 'Export as PPTX' }).click()])
  expect(dl.suggestedFilename()).toMatch(/\.pptx$/)
})

test('Files: shared file system, favourites, rename, trash', async ({ page }) => {
  await page.getByRole('button', { name: 'New document' }).click(); await page.getByLabel('Document body').click(); await page.keyboard.type('x')
  await page.getByLabel('Document title').fill('Shared doc'); await page.getByLabel('Document title').blur()
  await page.getByRole('link', { name: 'Back to home' }).click()
  await expect(page.getByRole('button', { name: /Shared doc\.docx/ })).toBeVisible() // recent updates
  await page.goto('/files')
  const row = page.getByTestId('file-row').filter({ hasText: 'Shared doc.docx' })
  await expect(row).toBeVisible()
  await row.getByRole('button', { name: 'More actions' }).click(); await page.getByRole('menuitem', { name: 'Add to favourites' }).click()
  await page.locator('aside').getByRole('button', { name: 'Favourites' }).click(); await expect(page.getByTestId('file-row').filter({ hasText: 'Shared doc' })).toHaveCount(1)
  await page.locator('aside').getByRole('button', { name: 'My files' }).click()
  await page.getByTestId('file-row').filter({ hasText: 'Shared doc' }).getByRole('button', { name: 'More actions' }).click(); await page.getByRole('menuitem', { name: 'Move to trash' }).click()
  await page.locator('aside').getByRole('button', { name: 'Trash' }).click(); await expect(page.getByTestId('file-row').filter({ hasText: 'Shared doc' })).toBeVisible()
})

test('Import: CSV, TXT and DOCX open in the right editor', async ({ page }) => {
  await page.getByTestId('home-upload').setInputFiles({ name: 'sales.csv', mimeType: 'text/csv', buffer: Buffer.from('a,b\n1,2\n=A2+B2,x') })
  await expect(page).toHaveURL(/\/excel\//)
  await expect(cell(page, 0, 0)).toHaveText('a'); await expect(cell(page, 2, 0)).toHaveText('3')
  await page.goto('/')
  await page.getByTestId('home-upload').setInputFiles({ name: 'notes.txt', mimeType: 'text/plain', buffer: Buffer.from('first line\nsecond line') })
  await expect(page).toHaveURL(/\/word\//); await expect(page.getByLabel('Document body')).toContainText('second line')
  const { Document, Packer, Paragraph, TextRun } = await import('docx')
  const buf = await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph({ children: [new TextRun({ text: 'From a real docx', bold: true })] })] }] }))
  await page.goto('/')
  await page.getByTestId('home-upload').setInputFiles({ name: 'real.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: buf })
  await expect(page.getByLabel('Document body').locator('strong')).toHaveText('From a real docx')
})

async function enter(page: Page, r: number, c: number, text: string) {
  await cell(page, r, c).click(); await page.keyboard.type(text); await page.keyboard.press('Enter')
}

test('Excel formulas: the six required examples, formula bar, recalculation, reload', async ({ page }) => {
  await page.getByRole('button', { name: 'New spreadsheet' }).click()
  await enter(page, 0, 0, '10'); await enter(page, 1, 0, '20'); await enter(page, 2, 0, '=SUM(A1:A2)')
  await enter(page, 0, 1, '50'); await enter(page, 1, 1, '25'); await enter(page, 2, 1, '=AVERAGE(B1:B2)')
  await enter(page, 0, 2, '100'); await enter(page, 1, 2, '=IF(C1>50,"PASS","FAIL")')
  await enter(page, 0, 3, '10'); await enter(page, 1, 3, '20'); await enter(page, 2, 3, '=D1+D2')
  await enter(page, 0, 4, '5'); await enter(page, 1, 4, '5'); await enter(page, 2, 4, '=PRODUCT(E1:E2)')
  await enter(page, 0, 5, '15'); await enter(page, 1, 5, '20'); await enter(page, 2, 5, '=MAX(F1:F2)')
  const expected: [number, number, string][] = [[2, 0, '30'], [2, 1, '37.5'], [1, 2, 'PASS'], [2, 3, '30'], [2, 4, '25'], [2, 5, '20']]
  for (const [r, c, v] of expected) await expect(cell(page, r, c)).toHaveText(v)
  // the formula bar shows the real formula, not the result
  await cell(page, 2, 1).click(); await expect(page.getByLabel('Formula bar')).toHaveValue('=AVERAGE(B1:B2)')
  // edit through the formula bar
  await cell(page, 2, 0).click(); await page.getByLabel('Formula bar').fill('=SUM(A1:A2)*2'); await page.keyboard.press('Enter')
  await expect(cell(page, 2, 0)).toHaveText('60')
  // recalculation: change a precedent, dependants follow (chain A1 -> A3)
  await enter(page, 0, 0, '15'); await expect(cell(page, 2, 0)).toHaveText('70')
  await enter(page, 0, 6, '=A3+1'); await expect(cell(page, 0, 6)).toHaveText('71')
  await enter(page, 0, 0, '0'); await expect(cell(page, 0, 6)).toHaveText('41')
  // text, lookup and error handling
  await enter(page, 4, 0, '=UPPER(LEFT("office",3))&LEN("ab")'); await expect(cell(page, 4, 0)).toHaveText('OFF2')
  await enter(page, 5, 0, '=1/0'); await expect(cell(page, 5, 0)).toHaveText('#DIV/0!')
  await enter(page, 6, 0, '=FOO(1)'); await expect(cell(page, 6, 0)).toHaveText('#NAME?')
  await expect(page.getByText('Saved')).toBeVisible()
  await page.reload()
  for (const [r, c, v] of [[2, 1, '37.5'], [1, 2, 'PASS'], [2, 3, '30'], [2, 4, '25'], [0, 6, '41']] as [number, number, string][]) await expect(cell(page, r, c)).toHaveText(v)
  await cell(page, 2, 4).click(); await expect(page.getByLabel('Formula bar')).toHaveValue('=PRODUCT(E1:E2)')
})

test('Excel: context menu insert row rewrites references; large sheets stay responsive', async ({ page }) => {
  await page.getByRole('button', { name: 'New spreadsheet' }).click()
  await enter(page, 0, 0, '1'); await enter(page, 1, 0, '2'); await enter(page, 2, 0, '=SUM(A1:A2)')
  await cell(page, 1, 0).click({ button: 'right' })
  await page.getByRole('menuitem', { name: 'Insert row above' }).click()
  await expect(cell(page, 3, 0)).toHaveText('3'); await cell(page, 3, 0).click(); await expect(page.getByLabel('Formula bar')).toHaveValue('=SUM(A1:A3)')
  // 20k-row CSV import
  const csv = Array.from({ length: 20000 }, (_, i) => `${i + 1},${(i + 1) * 2},=A${i + 1}+B${i + 1}`).join('\n')
  await page.goto('/')
  await page.getByTestId('home-upload').setInputFiles({ name: 'big.csv', mimeType: 'text/csv', buffer: Buffer.from(csv) })
  await expect(page).toHaveURL(/\/excel\//)
  await expect(cell(page, 0, 2)).toHaveText('3')
  expect(await page.locator('td[data-r]').count()).toBeLessThan(2000) // rows are windowed, not all 20k rendered
  const t0 = Date.now(); await cell(page, 0, 0).click(); await page.keyboard.type('100'); await page.keyboard.press('Enter')
  await expect(cell(page, 0, 2)).toHaveText('102'); expect(Date.now() - t0).toBeLessThan(2500)
  await page.getByTestId('grid').evaluate(el => { el.scrollTop = 19990 * 24 })
  await expect(page.locator('td[data-r="19999"][data-c="2"]')).toHaveText('60000')
})

test('PowerPoint fonts: searchable list shows only fonts that render; applied to selection and whole box', async ({ page }) => {
  await page.goto('/powerpoint'); await page.getByRole('button', { name: 'New blank presentation' }).click()
  const stage = page.getByTestId('stage')
  const title = stage.locator('[data-obj]').first()
  await title.dblclick()
  await page.keyboard.press('ControlOrMeta+a'); await page.keyboard.type('Hello fonts world')
  // select just "fonts" and style it
  for (let i = 0; i < 6; i++) await page.keyboard.press('ArrowLeft')
  for (let i = 0; i < 5; i++) await page.keyboard.press('Shift+ArrowLeft')
  await page.getByRole('button', { name: 'Font family' }).first().click()
  await page.getByLabel('Search fonts').fill('georg')
  await expect(page.getByRole('listbox', { name: 'Fonts' }).getByRole('option', { name: 'Georgia', exact: true })).toHaveCount(1) // other families may also contain "georg"
  await page.getByRole('listbox', { name: 'Fonts' }).getByRole('option', { name: 'Georgia', exact: true }).click()
  await page.getByRole('button', { name: 'Bold (⌘B)' }).first().click()
  const run = stage.locator('.slide-editor span[style*="font-family"]')
  await expect(run).toHaveText('fonts'); await expect(run).toHaveCSS('font-family', /Georgia/)
  // fonts offered must genuinely load and render when requested (metrics differ from the generic fallback)
  await page.getByRole('button', { name: 'Font family' }).first().click()
  const offered = await fontOptions(page).allInnerTexts()
  expect(offered.length).toBeGreaterThan(5)
  const bad = await page.evaluate(async names => { const m = await import('/src/services/fonts.ts'); const out: string[] = []; for (const n of names.filter(x => x !== 'Use the theme default').slice(0, 10)) { if (!(await m.loadFont(n.trim())) || !m.isRenderable(n.trim())) out.push(n) } return out }, offered)
  expect(bad).toEqual([])
  await page.keyboard.press('Escape'); await stage.click({ position: { x: 4, y: 4 } })
  // whole box: select the box (not editing) and choose a font -> every run changes
  await title.click({ position: { x: 3, y: 3 } })
  await page.getByRole('button', { name: 'Font family' }).first().click(); await page.getByLabel('Search fonts').fill('courier'); await fontOptions(page, 'Courier New').click()
  await expect(stage.locator('.slide-text').first()).toHaveCSS('font-family', /Courier New/)
  await expect(stage.locator('.slide-text span').first()).toHaveCSS('font-family', /Courier New/)
})

test('PowerPoint transitions: set, persist, apply to all, and actually animate in the slideshow', async ({ page }) => {
  await page.goto('/powerpoint'); await page.getByRole('button', { name: 'New blank presentation' }).click()
  for (let i = 0; i < 2; i++) { await page.getByRole('button', { name: 'New slide' }).click(); await page.getByRole('menuitem', { name: 'Title and content' }).click() }
  await expect(page.getByTestId('thumb')).toHaveCount(3)
  await page.getByTestId('thumb').nth(1).click()
  await page.getByRole('tab', { name: 'Transitions' }).click()
  await page.getByRole('radio', { name: 'Push', exact: true }).click()
  await page.getByLabel('Transition direction').selectOption('left')
  await page.getByLabel('Transition duration (seconds)').fill('1.5')
  await expect(page.getByTestId('thumb-transition')).toHaveCount(1)
  await page.getByRole('button', { name: 'Apply to all' }).click()
  await expect(page.getByTestId('thumb-transition')).toHaveCount(3)
  await page.getByRole('radio', { name: 'Zoom', exact: true }).click()
  await expect(page.getByTestId('thumb-transition')).toHaveCount(3)
  // preview plays inside the editor
  await page.getByRole('button', { name: 'Preview' }).click(); await expect(page.getByTestId('transition-preview')).toBeVisible()
  await expect(page.getByTestId('transition-preview')).toHaveCount(0, { timeout: 5000 })
  await page.getByRole('radio', { name: 'Push', exact: true }).click(); await page.getByLabel('Transition direction').selectOption('left'); await page.getByLabel('Transition duration (seconds)').fill('1.5')
  await expect(page.getByLabel('Transition duration (seconds)')).toHaveValue('1.5')
  // persists
  await expect(page.getByText('Saved')).toBeVisible(); await page.reload()
  await page.getByTestId('thumb').nth(1).click(); await page.getByRole('tab', { name: 'Transitions' }).click()
  await expect(page.getByRole('radio', { name: 'Push', exact: true })).toHaveAttribute('aria-checked', 'true'); await expect(page.getByLabel('Transition duration (seconds)')).toHaveValue('1.5'); await expect(page.getByLabel('Transition direction')).toHaveValue('left')
  // slideshow: entering slide 2 plays a push of 1500ms
  await page.getByTestId('thumb').nth(0).click()
  await page.getByRole('button', { name: 'Present from current slide' }).click()
  await expect(page.getByTestId('presenter')).toBeVisible()
  await page.keyboard.press('ArrowRight')
  await expect(page.getByTestId('transition-stage')).toHaveAttribute('data-transition', 'push')
  const timing = await page.evaluate(() => document.getAnimations().filter(a => (a.effect as KeyframeEffect).target?.closest('[data-testid=transition-stage]')).map(a => ({ d: (a.effect as KeyframeEffect).getTiming().duration, p: (a.effect as KeyframeEffect).getKeyframes().some(k => 'transform' in k) })))
  expect(timing.length).toBe(2); expect(timing.every(t => t.d === 1500 && t.p)).toBe(true)
  await expect(page.getByTestId('present-counter')).toHaveText('2 / 3')
  // pressing next during the animation finishes it and moves on without glitching
  await page.keyboard.press('ArrowRight'); await expect(page.getByTestId('present-counter')).toHaveText('3 / 3')
  await expect(page.getByTestId('transition-stage')).toHaveCount(0, { timeout: 4000 })
  await page.keyboard.press('ArrowLeft'); await expect(page.getByTestId('transition-stage')).toHaveCount(1); await expect(page.getByTestId('present-counter')).toHaveText('2 / 3')
  await expect(page.getByTestId('transition-stage')).toHaveCount(0, { timeout: 4000 })
  await page.keyboard.press('Escape'); await expect(page.getByTestId('presenter')).toHaveCount(0)
})

test('PowerPoint: every transition type animates; none is a placeholder', async ({ page }) => {
  await page.goto('/powerpoint'); await page.getByRole('button', { name: 'New blank presentation' }).click()
  await page.getByRole('button', { name: 'New slide' }).click(); await page.getByRole('menuitem', { name: 'Title and content' }).click()
  await page.getByRole('tab', { name: 'Transitions' }).click()
  const types = ['Fade', 'Push', 'Wipe', 'Split', 'Reveal', 'Cover', 'Uncover', 'Zoom', 'Dissolve']
  for (const name of types) {
    await page.getByRole('radio', { name, exact: true }).click(); await page.getByLabel('Transition duration (seconds)').fill('0.6')
    await page.getByRole('button', { name: 'Present from current slide' }).click()
    await page.keyboard.press('Home'); await page.keyboard.press('ArrowRight')
    await expect(page.getByTestId('transition-stage')).toHaveAttribute('data-transition', name.toLowerCase())
    const n = await page.evaluate(() => document.getAnimations().filter(a => (a.effect as KeyframeEffect).target?.closest('[data-testid=transition-stage]')).length)
    expect(n, name).toBeGreaterThan(0)
    await expect(page.getByTestId('transition-stage')).toHaveCount(0, { timeout: 3000 })
    await page.keyboard.press('Escape')
  }
  await page.getByRole('radio', { name: 'None', exact: true }).click(); await expect(page.getByTestId('thumb-transition')).toHaveCount(0)
  await page.getByRole('button', { name: 'Present from current slide' }).click(); await page.keyboard.press('Home'); await page.keyboard.press('ArrowRight')
  await expect(page.getByTestId('transition-stage')).toHaveCount(0); await expect(page.getByTestId('present-counter')).toHaveText('2 / 2')
})

test('Word: font picker, size, line spacing, lists, zoom, page indicator, image, DOCX round trip', async ({ page }, testInfo) => {
  await page.getByRole('button', { name: 'New document' }).click()
  const body = page.getByLabel('Document body')
  await body.click(); await page.keyboard.type('Styled heading text'); await page.keyboard.press('ControlOrMeta+a')
  await page.getByRole('button', { name: 'Font family' }).click(); await page.getByLabel('Search fonts').fill('verd'); await fontOptions(page, 'Verdana').click()
  await page.getByLabel('Font size').fill('28'); await page.getByLabel('Font size').press('Enter')
  await page.getByRole('button', { name: 'Line spacing' }).click(); await page.getByRole('menuitem', { name: '2', exact: true }).click()
  const span = body.locator('span[style*="font-family"]').first()
  await expect(span).toHaveCSS('font-family', /Verdana/); await expect(span).toHaveCSS('font-size', '28px')
  await expect(body.locator('span[style*="line-height"]').first()).toBeVisible()
  // lists + alignment + colour
  await page.getByRole('button', { name: 'Bullets' }).click(); await expect(body.locator('ul')).toHaveCount(1)
  await page.getByRole('button', { name: 'Centre' }).click(); await expect(body.locator('[style*="text-align: center"]').first()).toBeVisible()
  // status bar + zoom
  await expect(page.getByTestId('word-count')).toHaveText('3 words'); await expect(page.getByTestId('page-indicator')).toHaveText('Page 1 of 1')
  await page.getByRole('button', { name: 'Zoom in' }).click(); await expect(page.getByTitle('Reset zoom to 100%')).toHaveText('110%')
  // many paragraphs -> more pages
  await body.press('ArrowRight') // focuses the editor, then collapses the select-all to its end
  await body.press(LINE_END)
  await expect.poll(() => page.evaluate(() => (document.querySelector('.ProseMirror') as unknown as { editor: { state: { selection: { empty: boolean } } } }).editor.state.selection.empty)).toBe(true)
  await expect(body).toContainText('Styled heading text')
  for (let i = 0; i < 70; i++) { await page.keyboard.press('Enter'); await page.keyboard.type('line ' + i) }
  await expect(page.getByTestId('page-indicator')).toHaveText(/Page \d+ of [2-9]/)
  // image
  await page.getByRole('tab', { name: 'Insert' }).click()
  const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNQTX79HwAElwJzVt2vfAAAAABJRU5ErkJggg==', 'base64')
  await page.getByTestId('word-image-input').setInputFiles({ name: 'dot.png', mimeType: 'image/png', buffer: png })
  await expect(body.locator('img')).toHaveCount(1)
  await expect(body).toContainText('Styled heading text') // inserting a picture must not replace the selection
  await expect(page.getByText('Saved')).toBeVisible(); await page.reload()
  await expect(page.getByLabel('Document body').locator('img')).toHaveCount(1)
  await expect(page.getByLabel('Document body').locator('span[style*="Verdana"]').first()).toBeVisible()
  // DOCX export then re-import keeps content and formatting
  await page.getByRole('button', { name: 'File' }).click()
  const [dl] = await Promise.all([page.waitForEvent('download'), page.getByRole('menuitem', { name: 'Export as DOCX' }).click()])
  const path = testInfo.outputPath('roundtrip.docx')
  await dl.saveAs(path) // keep the real extension so the importer recognises it
  await page.goto('/'); await page.getByTestId('home-upload').setInputFiles(path)
  await expect(page).toHaveURL(/\/word\//)
  const re = page.getByLabel('Document body')
  await expect(re).toContainText('Styled heading text'); await expect(re).toContainText('line 69'); await expect(re.locator('li').first()).toBeVisible(); await expect(re.locator('img')).toHaveCount(1)
})

test('Word: find and replace, page setup, styles, navigation pane, word count', async ({ page }) => {
  await page.getByRole('button', { name: 'New document' }).click()
  const body = page.getByLabel('Document body')
  await body.click()
  await page.keyboard.type('Alpha chapter')
  await page.keyboard.press('Enter'); await page.keyboard.type('the quick brown fox jumps over the lazy dog. The End.')
  await page.keyboard.press('Enter'); await page.keyboard.type('Beta chapter')
  // styles + navigation pane outline
  await page.keyboard.press('Home'); await setStyle(page, 'Heading 1')
  await body.getByText('Alpha chapter').click(); await setStyle(page, 'Heading 1')
  await expect(body.locator('h1')).toHaveCount(2)
  await page.getByRole('button', { name: 'Navigation pane' }).first().click()
  const nav = page.getByTestId('nav-pane')
  await expect(nav.getByRole('button', { name: 'Alpha chapter' })).toBeVisible()
  await expect(nav.getByRole('button', { name: 'Beta chapter' })).toBeVisible()
  // find
  await page.keyboard.press('ControlOrMeta+f')
  await page.getByLabel('Find text').fill('the')
  await expect(page.getByTestId('find-count')).toHaveText('1 of 3')
  await page.getByRole('button', { name: 'Match case' }).click()
  await expect(page.getByTestId('find-count')).toHaveText('1 of 2') // "The End." no longer matches
  await page.getByRole('button', { name: 'Match case' }).click()
  await page.getByLabel('Find text').fill('he')
  await expect(page.getByTestId('find-count')).toHaveText('1 of 3') // substring of each "the"
  await page.getByRole('button', { name: 'Whole word' }).click()
  await expect(page.getByTestId('find-count')).toHaveText('No results')
  await page.getByRole('button', { name: 'Whole word' }).click()
  await page.getByLabel('Find text').fill('the')
  // replace all
  await page.getByRole('button', { name: 'Show replace' }).click()
  await page.getByLabel('Replace with').fill('THE')
  await page.getByRole('button', { name: 'Replace all' }).click()
  await expect(body).toContainText('THE quick brown fox jumps over THE lazy dog. THE End.')
  await page.getByRole('button', { name: 'Close find' }).click()
  // page setup changes the canvas and persists
  await page.getByRole('tab', { name: 'Layout' }).click()
  const docPage = page.locator('.doc-page')
  const before = (await docPage.boundingBox())!.width
  await page.getByRole('button', { name: 'Orientation' }).click(); await page.getByRole('menuitem', { name: /Landscape/ }).click()
  await expect.poll(async () => (await docPage.boundingBox())!.width).toBeGreaterThan(before)
  await page.getByRole('button', { name: 'Paper size' }).click(); await page.getByRole('menuitem', { name: /A4/ }).click()
  await page.getByRole('button', { name: 'Margins' }).click(); await page.getByRole('menuitem', { name: /Narrow/ }).click()
  await page.getByRole('button', { name: 'Columns' }).click(); await page.getByRole('menuitem', { name: /Two/ }).click()
  await expect(docPage).toHaveClass(/doc-columns-2/)
  await expect(page.getByText('Saved')).toBeVisible()
  await page.reload()
  await expect(page.locator('.doc-page')).toHaveClass(/doc-columns-2/)
  await expect(page.getByLabel('Document body')).toContainText('THE quick brown fox')
  // word count dialog
  await page.getByRole('tab', { name: 'Review' }).click()
  await page.getByRole('button', { name: 'Word count' }).click()
  await expect(page.getByRole('dialog', { name: 'Word count' })).toContainText('Paragraphs')
  await page.getByRole('button', { name: 'Close' }).click()
})

test('Word: case change, superscript, checklist, symbols and format painter', async ({ page }) => {
  await page.getByRole('button', { name: 'New document' }).click()
  const body = page.getByLabel('Document body')
  await body.click(); await page.keyboard.type('shout this')
  await page.keyboard.press('ControlOrMeta+a')
  await page.getByRole('button', { name: 'Change case' }).click()
  await page.getByRole('menuitem', { name: 'UPPERCASE' }).click()
  await expect(body).toContainText('SHOUT THIS')
  // superscript
  await page.keyboard.press('ControlOrMeta+End'); await page.getByRole('button', { name: 'Superscript' }).click()
  await page.keyboard.type('2')
  await expect(body.locator('sup')).toHaveText('2')
  // checklist
  await page.keyboard.press('Enter'); await page.getByRole('button', { name: 'Checklist' }).click(); await page.keyboard.type('task one')
  await expect(body.locator('ul[data-type="taskList"] input[type=checkbox]')).toHaveCount(1)
  // symbol + emoji
  await page.getByRole('tab', { name: 'Insert' }).click()
  await page.getByRole('button', { name: 'Insert symbol' }).click(); await page.getByRole('button', { name: '±', exact: true }).click()
  await expect(body).toContainText('±')
  await page.getByRole('button', { name: 'Insert emoji' }).click(); await page.getByRole('button', { name: '🚀', exact: true }).click()
  await expect(body).toContainText('🚀')
})

test('Excel: sort, conditional formatting, chart, freeze panes, find & replace, cell fonts', async ({ page }) => {
  await page.getByRole('button', { name: 'New spreadsheet' }).click()
  const rows: [string, string][] = [['Name', 'Score'], ['Zoe', '30'], ['Adam', '95'], ['Mia', '60']]
  for (let r = 0; r < rows.length; r++) { await enter(page, r, 0, rows[r][0]); await enter(page, r, 1, rows[r][1]) }
  // sort descending by Score, header stays on top
  await cell(page, 1, 1).click()
  await page.getByRole('tab', { name: 'Data' }).click()
  await page.getByRole('button', { name: 'Z → A' }).click()
  await expect(cell(page, 0, 0)).toHaveText('Name')
  await expect(cell(page, 1, 0)).toHaveText('Adam'); await expect(cell(page, 2, 0)).toHaveText('Mia'); await expect(cell(page, 3, 0)).toHaveText('Zoe')
  // conditional formatting: highlight scores above 50
  await cell(page, 1, 1).click({ modifiers: [] }); await page.getByLabel('Name box').fill('B2:B4'); await page.keyboard.press('Enter')
  await page.getByRole('tab', { name: 'Home' }).click()
  await page.getByRole('button', { name: 'Conditional formatting' }).click(); await page.getByRole('menuitem', { name: 'Greater than' }).click()
  await page.getByLabel('Value').fill('50'); await page.getByRole('button', { name: 'Apply' }).click()
  await expect(cell(page, 1, 1)).toHaveCSS('background-color', 'rgb(254, 202, 202)')
  await expect(cell(page, 3, 1)).not.toHaveCSS('background-color', 'rgb(254, 202, 202)')
  // chart from the data range
  await page.getByRole('tab', { name: 'Insert' }).click(); await page.getByRole('button', { name: 'Chart' }).click()
  await page.getByLabel('Chart data range').fill('A1:B4'); await page.getByLabel('Chart title').fill('Scores')
  await page.getByRole('dialog', { name: 'Insert chart' }).getByRole('button', { name: 'Insert' }).click()
  await expect(page.getByTestId('chart')).toHaveCount(1)
  await expect(page.getByTestId('chart').locator('rect')).toHaveCount(3) // one bar per row
  // freeze the top row and keep it visible while scrolled
  await page.getByRole('tab', { name: 'Data' }).click(); await page.getByRole('button', { name: 'Freeze panes' }).click(); await page.getByRole('menuitem', { name: 'Freeze top row' }).click()
  await page.getByTestId('grid').evaluate(el => { el.scrollTop = 600 })
  await expect.poll(async () => (await cell(page, 0, 0).boundingBox())?.y ?? -1).toBeGreaterThan(100)
  // find & replace
  await page.keyboard.press('Control+h')
  await page.getByLabel('Find in sheet').fill('mia'); await expect(page.getByTestId('xl-find-count')).toHaveText('1 of 1')
  await page.getByLabel('Replace with').fill('Maya'); await page.getByRole('button', { name: 'Replace all' }).click()
  await page.getByTestId('grid').evaluate(el => { el.scrollTop = 0 })
  await expect(cell(page, 2, 0)).toHaveText('Maya')
  // per-cell font from the full font library, plus underline and decimals
  await page.getByRole('button', { name: 'Close find' }).click()
  await page.getByRole('tab', { name: 'Home' }).click()
  await cell(page, 2, 0).click()
  await page.getByRole('button', { name: 'Font family' }).click(); await page.getByLabel('Search fonts').fill('georgia')
  await page.getByRole('listbox', { name: 'Fonts' }).getByRole('option', { name: 'Georgia', exact: true }).click()
  await page.getByRole('button', { name: 'Underline (⌘U)' }).click()
  await expect(cell(page, 2, 0)).toHaveCSS('font-family', /Georgia/); await expect(cell(page, 2, 0)).toHaveCSS('text-decoration-line', 'underline')
  await cell(page, 2, 1).click(); await page.getByRole('button', { name: 'Increase decimal places' }).click()
  await expect(cell(page, 2, 1)).toHaveText('60.0')
  // everything persists
  await expect(page.getByText('Saved')).toBeVisible(); await page.reload()
  await expect(page.getByTestId('chart')).toHaveCount(1)
  await expect(cell(page, 2, 0)).toHaveText('Maya'); await expect(cell(page, 2, 1)).toHaveText('60.0')
  await expect(cell(page, 1, 1)).toHaveCSS('background-color', 'rgb(254, 202, 202)')
})

test('PowerPoint: multi-select, group, align, distribute, notes and presenter view', async ({ page }, testInfo) => {
  await page.goto('/powerpoint'); await page.getByRole('button', { name: 'New blank presentation' }).click()
  await page.getByRole('button', { name: 'New slide' }).click(); await page.getByRole('menuitem', { name: 'Blank' }).click()
  await page.getByRole('tab', { name: 'Insert' }).click()
  const objs = page.getByTestId('stage').locator('[data-obj]')
  const rectOf = async (i: number) => (await objs.nth(i).boundingBox())!
  // a new shape is selected as soon as it exists, so position each one straight away through the properties panel
  const spots: [string, number, number][] = [['Rectangle', 100, 100], ['Ellipse', 400, 250], ['Triangle', 700, 150]]
  for (const [shape, x, y] of spots) {
    await page.getByRole('button', { name: shape, exact: true }).click()
    await page.getByLabel('X', { exact: true }).fill(String(x)); await page.getByLabel('Y', { exact: true }).fill(String(y))
  }
  await expect(objs).toHaveCount(3)
  await page.getByTestId('stage').click({ position: { x: 3, y: 3 } })
  await page.keyboard.press('ControlOrMeta+a')
  await expect(page.getByTestId('multi-title')).toHaveText('3 objects selected')
  await expect(page.getByTestId('selection-box')).toBeVisible()
  // align tops, then distribute horizontally
  await page.getByRole('button', { name: 'Align top edge' }).click()
  const ys = [await rectOf(0), await rectOf(1), await rectOf(2)].map(b => Math.round(b.y))
  expect(new Set(ys).size).toBe(1)
  await page.getByRole('button', { name: 'Distribute H' }).click()
  const [b0, b1, b2] = [await rectOf(0), await rectOf(1), await rectOf(2)]
  expect(Math.abs((b1.x - (b0.x + b0.width)) - (b2.x - (b1.x + b1.width)))).toBeLessThan(3)
  // group, then dragging one member moves them all
  await page.getByRole('button', { name: 'Group' }).first().click()
  await page.getByTestId('stage').click({ position: { x: 3, y: 3 } })
  await expect(page.getByTestId('multi-title')).toHaveCount(0)
  await objs.nth(1).click({ position: { x: 5, y: 5 } })
  await expect(page.getByTestId('multi-title')).toContainText('(grouped)') // clicking one selects the whole group
  const before = await rectOf(0)
  const mid = await rectOf(1)
  await page.mouse.move(mid.x + mid.width / 2, mid.y + mid.height / 2); await page.mouse.down()
  await page.mouse.move(mid.x + mid.width / 2 + 40, mid.y + mid.height / 2 + 30, { steps: 6 }); await page.mouse.up()
  const after = await rectOf(0)
  expect(Math.round(after.x - before.x)).toBeGreaterThan(20); expect(Math.round(after.y - before.y)).toBeGreaterThan(15)
  // undo restores the position in one step
  await page.keyboard.press('ControlOrMeta+z')
  expect(Math.abs((await rectOf(0)).x - before.x)).toBeLessThan(2)
  // marquee selection picks up everything it touches
  await page.getByTestId('stage').click({ position: { x: 3, y: 3 } })
  const stage = (await page.getByTestId('stage').boundingBox())!
  await page.mouse.move(stage.x + 4, stage.y + 4); await page.mouse.down(); await page.mouse.move(stage.x + stage.width - 4, stage.y + stage.height - 4, { steps: 5 })
  await expect(page.getByTestId('marquee')).toBeVisible(); await page.mouse.up()
  await expect(page.getByTestId('multi-title')).toContainText('3 objects selected')
  // speaker notes persist, appear in presenter view, and are exported
  await page.getByRole('button', { name: 'Show notes pane' }).click()
  await page.getByLabel('Speaker notes').fill('Remember to mention the roadmap.')
  await page.getByTestId('thumb').nth(0).click(); await page.getByTestId('thumb').nth(1).click()
  await expect(page.getByLabel('Speaker notes')).toHaveValue('Remember to mention the roadmap.')
  await expect(page.getByText('Saved')).toBeVisible(); await page.reload()
  await page.getByTestId('thumb').nth(1).click(); await page.getByRole('button', { name: 'Show notes pane' }).click()
  await expect(page.getByLabel('Speaker notes')).toHaveValue('Remember to mention the roadmap.')
  await page.getByRole('button', { name: 'Present from current slide' }).click()
  await expect(page.getByTestId('presenter')).toHaveAttribute('data-view', 'audience')
  await page.keyboard.press('v')
  await expect(page.getByTestId('presenter')).toHaveAttribute('data-view', 'presenter')
  await expect(page.getByTestId('present-notes')).toHaveText('Remember to mention the roadmap.')
  await expect(page.getByTestId('present-timer')).toHaveText(/^\d\d:\d\d$/)
  await page.keyboard.press('Home'); await expect(page.getByTestId('present-counter')).toHaveText('1 / 2')
  await expect(page.getByTestId('present-notes')).toContainText('No notes')
  await page.keyboard.press('Escape')
  await page.getByRole('button', { name: 'File' }).click()
  const [dl] = await Promise.all([page.waitForEvent('download'), page.getByRole('menuitem', { name: 'Export as PPTX' }).click()])
  const file = testInfo.outputPath('notes.pptx'); await dl.saveAs(file)
  const zip = (await import('node:fs')).readFileSync(file)
  const { default: JSZip } = await import('jszip')
  const z = await JSZip.loadAsync(zip)
  const notes = await Promise.all(Object.keys(z.files).filter(n => /notesSlides\/notesSlide\d+\.xml$/.test(n)).map(n => z.files[n].async('string')))
  expect(notes.join(' ')).toContain('Remember to mention the roadmap.')
})

test('Home: templates create real documents, live search finds files, greeting is shown', async ({ page }) => {
  await expect(page.getByRole('heading', { level: 1 })).toHaveText(/Good (morning|afternoon|evening)|Working late/)
  // Excel template: formulas compute
  await page.getByRole('button', { name: 'Template Monthly budget' }).click()
  await expect(page).toHaveURL(/\/excel\//)
  await expect(cell(page, 4, 7)).toHaveText(/^\$?[\d,]+\.\d\d$/) // total income formula evaluated, shown as currency
  await cell(page, 4, 7).click(); await expect(page.getByLabel('Formula bar')).toHaveValue('=SUM(B5:G5)')
  await expect(cell(page, 13, 1)).toContainText('%')
  // Word template
  await page.goto('/'); await page.getByRole('button', { name: 'Template Meeting notes' }).click()
  await expect(page).toHaveURL(/\/word\//)
  await expect(page.getByLabel('Document body').locator('h1')).toHaveText('Meeting notes')
  await expect(page.getByLabel('Document body').locator('input[type=checkbox]')).toHaveCount(1)
  // live search from the header, in the files we just made
  await page.goto('/')
  await page.getByRole('button', { name: 'Search files and commands' }).click()
  await page.getByLabel('Search commands and files').fill('budget')
  await expect(page.getByRole('listbox', { name: 'Results' }).getByRole('option', { name: /Monthly budget/ })).toBeVisible()
  await page.keyboard.press('Enter')
  await expect(page).toHaveURL(/\/excel\//)
})

test('UX: command palette, drop-to-import, trash with undo, grid and list layouts', async ({ page }) => {
  // palette: actions by typing, files by name, and it works from inside an editor
  await expect(page.getByRole('heading', { level: 1 })).toBeVisible() // wait for the app to be interactive
  await page.keyboard.press('Control+k')
  await page.getByLabel('Search commands and files').fill('new spread'); await page.keyboard.press('Enter')
  await expect(page).toHaveURL(/\/excel\//)
  await page.keyboard.press('Control+k')
  await page.getByLabel('Search commands and files').fill('dark'); await expect(page.getByRole('option', { name: /dark theme/ })).toBeVisible(); await page.keyboard.press('Enter')
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark')
  await page.keyboard.press('Control+k'); await page.getByLabel('Search commands and files').fill('zzzz')
  await expect(page.getByText(/Nothing matches/)).toBeVisible(); await page.keyboard.press('Escape')

  // dropping a file anywhere imports it and opens it
  await page.goto('/')
  await page.evaluate(() => {
    const dt = new DataTransfer(); dt.items.add(new File(['x,y\n1,2'], 'dropped.csv', { type: 'text/csv' }))
    for (const type of ['dragenter', 'dragover']) window.dispatchEvent(new DragEvent(type, { dataTransfer: dt, bubbles: true, cancelable: true }))
  })
  await expect(page.getByTestId('drop-overlay')).toBeVisible()
  await page.evaluate(() => {
    const dt = new DataTransfer(); dt.items.add(new File(['x,y\n1,2'], 'dropped.csv', { type: 'text/csv' }))
    window.dispatchEvent(new DragEvent('drop', { dataTransfer: dt, bubbles: true, cancelable: true }))
  })
  await expect(page).toHaveURL(/\/excel\//)
  await expect(cell(page, 0, 0)).toHaveText('x')

  // Drive: grid by default with real previews, switch to list, trash then undo
  await page.goto('/files')
  await expect(page.getByTestId('file-row').first()).toBeVisible()
  await expect(page.getByRole('group', { name: 'Layout' }).getByRole('button', { name: 'Grid view' })).toHaveAttribute('aria-pressed', 'true')
  await page.getByRole('button', { name: 'List view' }).click(); await expect(page.getByRole('table', { name: 'Files' })).toBeVisible()
  await page.reload(); await expect(page.getByRole('table', { name: 'Files' })).toBeVisible() // the choice is remembered
  const before = await page.getByTestId('file-row').count()
  await page.getByTestId('file-row').filter({ hasText: 'dropped.xlsx' }).getByRole('button', { name: 'More actions' }).click()
  await page.getByRole('menuitem', { name: 'Move to trash' }).click()
  await expect(page.getByTestId('file-row')).toHaveCount(before - 1)
  await page.getByRole('button', { name: 'Undo' }).click()
  await expect(page.getByTestId('file-row')).toHaveCount(before)
  // permanent deletion asks first
  await page.getByTestId('file-row').filter({ hasText: 'dropped.xlsx' }).getByRole('button', { name: 'More actions' }).click(); await page.getByRole('menuitem', { name: 'Move to trash' }).click()
  await page.locator('aside').getByRole('button', { name: /Trash/ }).click()
  await page.getByTestId('file-row').first().getByRole('button', { name: 'More actions' }).click(); await page.getByRole('menuitem', { name: 'Delete forever…' }).click()
  await expect(page.getByRole('dialog', { name: 'Delete forever?' })).toBeVisible()
  await page.getByRole('button', { name: 'Cancel' }).click(); await expect(page.getByTestId('file-row')).toHaveCount(1)
})

/* ------------------------------------------------------------------ PowerPoint theme library */
const openDesign = async (page: Page, template = 'Template Midnight') => {
  await page.goto('/powerpoint'); await page.getByRole('button', { name: template }).click()
  await expect(page).toHaveURL(/\/powerpoint\//); await page.getByRole('tab', { name: 'Design' }).click()
}
const stageBg = (page: Page) => page.getByTestId('stage').locator('xpath=./div').evaluate(el => getComputedStyle(el).backgroundColor)

test('PowerPoint themes: landing page shows every theme as a live preview and can filter', async ({ page }) => {
  await page.goto('/powerpoint')
  const tiles = page.getByRole('button', { name: /^Template / })
  await expect(page.getByRole('button', { name: 'Template Midnight' })).toBeVisible()
  await expect.poll(() => tiles.count()).toBeGreaterThanOrEqual(21)
  await expect(page.getByRole('button', { name: 'Template Midnight' })).toBeVisible(); await expect(page.getByRole('button', { name: 'Template Sunrise' })).toBeVisible()
  await expect(page.getByRole('button', { name: 'Template Midnight' }).locator('.slide-text').first()).toHaveText('Presentation title') // drawn by SlideView
  await page.getByRole('group', { name: 'Filter themes' }).getByRole('button', { name: 'Dark' }).click()
  await expect(page.getByRole('button', { name: 'Template Midnight' })).toBeVisible(); await expect(page.getByRole('button', { name: 'Template Sunrise' })).toHaveCount(0)
  await page.getByRole('group', { name: 'Filter themes' }).getByRole('button', { name: 'Colourful' }).click()
  await expect(page.getByRole('button', { name: 'Template Kiln' })).toBeVisible(); await expect(page.getByRole('button', { name: 'Template Midnight' })).toHaveCount(0)
})

test('PowerPoint themes: gallery lists all themes, filters, marks the current one and switches without losing content', async ({ page }) => {
  await openDesign(page)
  const gallery = page.getByTestId('theme-gallery')
  await expect(gallery.getByRole('button', { name: 'Theme Midnight' })).toBeVisible()
  await expect.poll(() => gallery.getByRole('button', { name: /^Theme / }).count()).toBeGreaterThanOrEqual(21)
  await expect(gallery.getByRole('button', { name: 'Theme Midnight' })).toHaveAttribute('aria-pressed', 'true')
  await expect(gallery.getByRole('button', { name: 'Theme Classic Blue' })).toHaveAttribute('aria-pressed', 'false')
  await expect(gallery.getByRole('button', { name: 'Theme Midnight' }).locator('.slide-text').first()).toHaveText('Presentation title')
  await page.getByRole('group', { name: 'Filter themes' }).getByRole('button', { name: 'Light' }).click()
  await expect(gallery.getByRole('button', { name: 'Theme Midnight' })).toHaveCount(0); await expect(gallery.getByRole('button', { name: 'Theme Harbour' })).toBeVisible()
  await page.getByRole('group', { name: 'Filter themes' }).getByRole('button', { name: 'All' }).click()

  // content survives a theme change; the slide takes the theme's colours
  await page.getByRole('tab', { name: 'Home' }).click()
  await page.getByRole('button', { name: 'New slide' }).click(); await page.getByRole('menuitem', { name: 'Title and content' }).click()
  await page.getByRole('tab', { name: 'Design' }).click()
  await page.getByRole('button', { name: 'Theme Sorbet' }).click()
  await expect(page.getByTestId('thumb')).toHaveCount(2)
  await expect(page.getByTestId('stage').getByText('Slide title')).toBeVisible()
  expect(await stageBg(page)).toBe('rgb(255, 227, 236)')
  await expect(page.getByRole('button', { name: 'Theme Sorbet' })).toHaveAttribute('aria-pressed', 'true')
  await expect(page.getByRole('button', { name: 'Theme Midnight' })).toHaveAttribute('aria-pressed', 'false')
  // heading and body faces differ and are applied
  const faces = await page.getByTestId('stage').locator('.slide-text').evaluateAll(els => els.map(e => getComputedStyle(e).fontFamily))
  expect(faces[0]).toContain('Young Serif'); expect(faces[1]).toContain('Figtree')
})

test('PowerPoint themes: gradient and pattern backgrounds render, and the chosen theme survives a reload', async ({ page }) => {
  await openDesign(page)
  await page.getByRole('button', { name: 'Theme Blueprint' }).click()
  await expect(page.getByTestId('stage').locator('xpath=./div')).toHaveCSS('background-image', /linear-gradient/)
  await page.getByRole('button', { name: 'Theme Lumen' }).click()
  await expect(page.getByTestId('stage').locator('xpath=./div')).toHaveCSS('background-image', /linear-gradient\(135deg/)
  await expect(page.getByText('Saved', { exact: false }).first()).toBeVisible({ timeout: 10_000 })
  await page.waitForTimeout(1200)
  await page.reload(); await page.getByRole('tab', { name: 'Design' }).click()
  await expect(page.getByRole('button', { name: 'Theme Lumen' })).toHaveAttribute('aria-pressed', 'true')
  await expect(page.getByTestId('stage').locator('xpath=./div')).toHaveCSS('background-image', /linear-gradient\(135deg/)
  await expect(page.getByText('Lumen').last()).toBeVisible()
})

test('PowerPoint themes: theme fonts are fetched and render when the deck opens', async ({ page }) => {
  await openDesign(page)
  await page.getByRole('button', { name: 'Theme Plum Noir' }).click()
  await expect.poll(() => page.evaluate(() => document.fonts.check('16px "Gloock"') && document.fonts.check('16px "Figtree"')), { timeout: 20_000 }).toBe(true)
  await page.waitForTimeout(1200); await page.reload()
  await expect.poll(() => page.evaluate(() => document.fonts.check('16px "Gloock"') && document.fonts.check('16px "Figtree"')), { timeout: 20_000 }).toBe(true)
})

test('PowerPoint themes: PPTX carries theme fonts, colours and background; PDF exports', async ({ page }, testInfo) => {
  await openDesign(page)
  await page.getByRole('button', { name: 'Theme Lumen' }).click()
  await page.getByRole('tab', { name: 'Home' }).click(); await page.getByRole('button', { name: 'New slide' }).click(); await page.getByRole('menuitem', { name: 'Title and content' }).click()
  await page.getByRole('button', { name: 'File' }).click()
  const [dl] = await Promise.all([page.waitForEvent('download'), page.getByRole('menuitem', { name: 'Export as PPTX' }).click()])
  const file = testInfo.outputPath('lumen.pptx'); await dl.saveAs(file)
  const { default: JSZip } = await import('jszip')
  const zip = await JSZip.loadAsync(await (await import('node:fs/promises')).readFile(file))
  const slide1 = await zip.file('ppt/slides/slide1.xml')!.async('string'), theme = await zip.file('ppt/theme/theme1.xml')!.async('string')
  expect(theme).toContain('typeface="Bricolage Grotesque"'); expect(theme).toContain('typeface="Hanken Grotesk"')
  expect(slide1).toContain('typeface="Bricolage Grotesque"') // title run
  expect(slide1).toContain('FFFFFF') // title colour
  expect(slide1).toContain('FF6FB5'); expect(slide1).toContain('8B6CFF') // decoration colours
  expect(slide1).toMatch(/<a:alpha val="\d+"/) // translucent shapes keep transparency
  expect(slide1).toContain('<p:bg>'); expect(slide1).toContain('r:embed') // gradient background exported as picture
  expect(Object.keys(zip.files).some(f => f.startsWith('ppt/media/'))).toBe(true)

  await page.getByRole('button', { name: 'File' }).click()
  const [pdf] = await Promise.all([page.waitForEvent('download', { timeout: 60_000 }), page.getByRole('menuitem', { name: 'Export as PDF' }).click()])
  expect(pdf.suggestedFilename()).toMatch(/\.pdf$/)
  const pdfFile = testInfo.outputPath('lumen.pdf'); await pdf.saveAs(pdfFile)
  const head = (await (await import('node:fs/promises')).readFile(pdfFile)).subarray(0, 5).toString()
  expect(head).toBe('%PDF-')
})
