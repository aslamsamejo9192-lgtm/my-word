import { chromium } from 'playwright'
import fs from 'fs'

const BASE = 'http://127.0.0.1:4173'
const results = []
const check = (name, ok, detail = '') => {
  results.push({ name, ok, detail })
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? ' — ' + detail : ''}`)
}

const browser = await chromium.launch({ args: ['--no-sandbox'] })
const ctx = await browser.newContext({
  viewport: { width: 1280, height: 860 },
  acceptDownloads: true,
})
const page = await ctx.newPage()
const consoleErrors = []
const pageErrors = []
page.on('console', (m) => {
  if (m.type() === 'error') consoleErrors.push(m.text())
})
page.on('pageerror', (e) => pageErrors.push(String(e)))

try {
  // 1. Dashboard loads
  await page.goto(BASE, { waitUntil: 'networkidle' })
  await page.waitForTimeout(1500)
  const dashText = await page.textContent('body')
  check('dashboard loads', /My Word/.test(dashText || '') && /Blank Document/.test(dashText || ''))

  // 2. Create blank document -> editor loads
  await page.click('button:has-text("New Document")')
  await page.waitForSelector('.ProseMirror', { timeout: 10000 })
  await page.waitForTimeout(1000)
  const ribbon = await page.textContent('body')
  check('editor loads with ribbon', /HOME/.test(ribbon || '') && /INSERT/.test(ribbon || '') && /VIEW/.test(ribbon || ''))

  // 3. Type text, select all, bold via ribbon button
  await page.click('.ProseMirror')
  await page.keyboard.type('Smoke test heading')
  await page.keyboard.press('Enter')
  await page.keyboard.type('Body text for export validation.')
  await page.keyboard.press('ControlOrMeta+a')
  await page.click('button[title="Bold (Ctrl+B)"]')
  await page.waitForTimeout(400)
  const boldOn = await page.evaluate(() => {
    const b = document.querySelector('button[title="Bold (Ctrl+B)"]')
    return b ? b.className.includes('bg-blue-200') : false
  })
  check('bold applies to selection', boldOn)
  await page.keyboard.press('ArrowRight')
  await page.keyboard.press('ControlOrMeta+s')
  await page.waitForTimeout(2500) // autosave/debounce + save
  const status = await page.textContent('body')
  check('save status shows Saved', /Saved/.test(status || ''))

  // 4. Reload -> document persists from IndexedDB (recovery)
  await page.reload({ waitUntil: 'networkidle' })
  await page.waitForSelector('.ProseMirror', { timeout: 10000 })
  await page.waitForTimeout(1500)
  const afterReload = await page.evaluate(() => document.querySelector('.ProseMirror')?.textContent || '')
  check(
    'document persists after reload',
    afterReload.includes('Smoke test heading') && afterReload.includes('Body text for export validation.'),
    JSON.stringify(afterReload.slice(0, 80)),
  )

  // 5. DOCX export -> valid file (download)
  await page.click('button:has-text("FILE")')
  await page.waitForTimeout(500)
  const downloadPromise = page.waitForEvent('download', { timeout: 15000 })
  await page.click('text=Export as DOCX')
  const download = await downloadPromise
  const dlPath = '/home/hatch/workspace/my-word/e2e/exported.docx'
  await download.saveAs(dlPath)
  const size = fs.statSync(dlPath).size
  check('DOCX export downloads', size > 1000, `${size} bytes`)

  // 6. Back to dashboard, import the sample DOCX
  await page.click('button[title="Close document"]')
  await page.waitForTimeout(1200)
  const dash2 = await page.textContent('body')
  check('back to dashboard', /My Documents/.test(dash2 || ''))
  const fileInput = page.locator('input[type="file"][accept*=".docx"]')
  await fileInput.setInputFiles('/home/hatch/workspace/my-word/e2e/sample-import.docx')
  await page.waitForSelector('.ProseMirror', { timeout: 15000 })
  await page.waitForTimeout(2000)
  const imported = await page.evaluate(() => document.querySelector('.ProseMirror')?.textContent || '')
  check(
    'DOCX import shows content',
    imported.includes('My Word Import Test') && imported.includes('bold text') && imported.includes('Cell A1'),
    JSON.stringify(imported.slice(0, 100)),
  )

  // 7. PDF workspace: open sample PDF
  await page.click('button[title="Close document"]')
  await page.waitForTimeout(1200)
  await page.click('button:has-text("Open PDF")')
  await page.waitForTimeout(800)
  const pdfInput = page.locator('input[type="file"][accept*="pdf"]')
  await pdfInput.setInputFiles('/home/hatch/workspace/my-word/e2e/sample.pdf')
  await page.waitForTimeout(4000)
  const pdfText = await page.textContent('body')
  const canvases = await page.locator('canvas').count()
  check(
    'PDF workspace renders pages',
    /Page 1 of 3/.test(pdfText || '') && canvases >= 1,
    `${canvases} canvas(es)`,
  )

  // 8. PDF search finds text
  const searchInput = page.locator('input[placeholder*="Search" i], input[type="search"]').first()
  if (await searchInput.count()) {
    await searchInput.fill('sample text')
    await page.keyboard.press('Enter')
    await page.waitForTimeout(2500)
    const afterSearch = await page.textContent('body')
    check('PDF search finds matches', /match/i.test(afterSearch || ''))
  } else {
    check('PDF search finds matches', false, 'no search input found')
  }
} catch (e) {
  check('smoke run completed without exception', false, String(e).slice(0, 300))
}

check('no page errors', pageErrors.length === 0, pageErrors.slice(0, 3).join(' | '))
check(
  'no console errors',
  consoleErrors.length === 0,
  consoleErrors.slice(0, 3).join(' | '),
)

await browser.close()
const failed = results.filter((r) => !r.ok)
console.log(`\n${results.length - failed.length}/${results.length} checks passed`)
process.exit(failed.length ? 1 : 0)
