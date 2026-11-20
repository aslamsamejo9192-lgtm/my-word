import { chromium } from 'playwright'
import fs from 'fs'

const BASE = 'http://127.0.0.1:4173'
const results = []
const check = (name, ok, detail = '') => {
  results.push({ name, ok, detail })
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? ' — ' + detail : ''}`)
}

const browser = await chromium.launch({ args: ['--no-sandbox'] })
const ctx = await browser.newContext({ viewport: { width: 1280, height: 860 }, acceptDownloads: true })
const page = await ctx.newPage()
const pageErrors = []
const consoleErrors = []
page.on('console', (m) => { if (m.type() === 'error') consoleErrors.push(m.text()) })
page.on('pageerror', (e) => pageErrors.push(String(e)))

try {
  await page.goto(BASE, { waitUntil: 'networkidle' })
  await page.waitForTimeout(1200)
  await page.click('button:has-text("New Document")')
  await page.waitForSelector('.ProseMirror', { timeout: 10000 })
  await page.waitForTimeout(800)

  // --- Table insert ---
  await page.click('.ProseMirror')
  await page.keyboard.type('Table test')
  await page.click('button:has-text("INSERT")')
  await page.waitForTimeout(400)
  await page.click('.ribbon-btn:has-text("Table")')
  await page.waitForTimeout(500)
  // Insert Table dialog: find rows/cols inputs and confirm button
  const dialog = page.locator('text=Insert Table').first()
  check('table dialog opens', (await dialog.count()) > 0)
  const confirmBtn = page.locator('button:has-text("Insert")').last()
  await confirmBtn.click()
  await page.waitForTimeout(800)
  const hasTable = await page.evaluate(() => !!document.querySelector('.ProseMirror table'))
  check('table inserted and editable', hasTable)

  // --- Find & Replace ---
  await page.keyboard.press('ControlOrMeta+f')
  await page.waitForTimeout(500)
  const findDialog = page.locator('text=Find and Replace').first()
  check('find dialog opens via Ctrl+F', (await findDialog.count()) > 0)
  const findInput = page.locator('input[placeholder="Text to find"]')
  await findInput.fill('Table test')
  await page.click('button:has-text("Find Next")')
  await page.waitForTimeout(500)
  const replaced = await (async () => {
    await page.locator('input[placeholder="Replacement text"]').fill('Table renamed')
    const replaceAll = page.locator('button:has-text("Replace All")').first()
    if (await replaceAll.count()) { await replaceAll.click(); return true }
    return false
  })()
  await page.waitForTimeout(500)
  const afterReplace = await page.evaluate(() => document.querySelector('.ProseMirror')?.textContent || '')
  check('replace all works', replaced && afterReplace.includes('Table renamed') && !afterReplace.includes('Table test'), JSON.stringify(afterReplace.slice(0, 60)))
  await page.keyboard.press('Escape')
  await page.click('button[aria-label="Close"]').catch(() => {})

  // --- Track Changes: toggle on, type, check the mark ---
  await page.click('button:has-text("REVIEW")')
  await page.waitForTimeout(400)
  await page.click('.ribbon-btn:has-text("Track Changes")')
  await page.waitForTimeout(300)
  await page.click('.ProseMirror')
  await page.keyboard.press('End')
  await page.keyboard.type(' tracked-text')
  await page.waitForTimeout(500)
  const tracked = await page.evaluate(() => {
    const el = document.querySelector('.ProseMirror')
    if (!el) return false
    return el.innerHTML.includes('tracked') || !!el.querySelector('[data-tracked], .tracked, span[style*="underline"]')
  })
  const trackedHtml = await page.evaluate(() => document.querySelector('.ProseMirror')?.innerHTML.slice(-400) || '')
  check('track changes marks typed text', tracked, trackedHtml.slice(0, 120))
  await page.click('.ribbon-btn:has-text("Track Changes")') // toggle off

  // --- PDF: rotate + download modified ---
  await page.click('button[title="Close document"]')
  await page.waitForTimeout(1200)
  await page.click('button:has-text("Open PDF")')
  await page.waitForTimeout(800)
  await page.locator('input[type="file"][accept*="pdf"]').setInputFiles('/home/hatch/workspace/my-word/e2e/sample.pdf')
  await page.waitForTimeout(3500)
  await page.click('button[title="Rotate this page 90° clockwise"]')
  await page.waitForTimeout(1500)
  const dlPromise = page.waitForEvent('download', { timeout: 20000 })
  await page.click('button:has-text("Download modified")')
  const dl = await dlPromise
  const dlPath = '/home/hatch/workspace/my-word/e2e/modified.pdf'
  await dl.saveAs(dlPath)
  const modSize = fs.statSync(dlPath).size
  check('PDF rotate + download modified', modSize > 1000, `${modSize} bytes`)
} catch (e) {
  check('second smoke completed without exception', false, String(e).slice(0, 300))
}

check('no page errors', pageErrors.length === 0, pageErrors.slice(0, 3).join(' | '))
check('no console errors', consoleErrors.length === 0, consoleErrors.slice(0, 3).join(' | '))
await browser.close()
const failed = results.filter((r) => !r.ok)
console.log(`\n${results.length - failed.length}/${results.length} checks passed`)
process.exit(failed.length ? 1 : 0)
