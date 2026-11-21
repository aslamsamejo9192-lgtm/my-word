import { PDFDocument, rgb, type PDFPage, type PDFFont, type RGB } from 'pdf-lib'
import fontkit from '@pdf-lib/fontkit'
import type { MyDocument } from './types'
import { pageDims } from './types'
import sansUrl from '../assets/fonts/DejaVuSans.ttf?url'
import sansBoldUrl from '../assets/fonts/DejaVuSans-Bold.ttf?url'
import sansItalicUrl from '../assets/fonts/LiberationSans-Italic.ttf?url'
import sansBoldItalicUrl from '../assets/fonts/LiberationSans-BoldItalic.ttf?url'
import serifUrl from '../assets/fonts/DejaVuSerif.ttf?url'
import serifBoldUrl from '../assets/fonts/DejaVuSerif-Bold.ttf?url'
import serifItalicUrl from '../assets/fonts/LiberationSerif-Italic.ttf?url'
import serifBoldItalicUrl from '../assets/fonts/LiberationSerif-BoldItalic.ttf?url'
import monoUrl from '../assets/fonts/DejaVuSansMono.ttf?url'
import monoBoldUrl from '../assets/fonts/DejaVuSansMono-Bold.ttf?url'
import {
  sansItalicRanges,
  sansBoldItalicRanges,
  serifItalicRanges,
  serifBoldItalicRanges,
} from './fontCoverage'

/* ------------------------------------------------------------------ */
/* Real PDF export (Tiptap JSON -> A4 PDF via pdf-lib)                 */
/*                                                                     */
/* The standard Helvetica fonts cannot render the sub/superscripts,  */
/* arrows and Greek letters in these documents, so real TTFs are     */
/* embedded: DejaVu (Sans / Serif / Mono, regular + bold) for full    */
/* glyph coverage, Liberation italics for slanted text. An italic     */
/* run whose text the italic font cannot render falls back to the     */
/* upright font, so the PDF never shows missing-glyph boxes. Fonts    */
/* are fetched as separate assets only when the user exports.         */
/* ------------------------------------------------------------------ */

type AnyNode = Record<string, unknown>

type FontFamily = 'sans' | 'serif' | 'mono'

interface LoadedFonts {
  sans: PDFFont
  sansBold: PDFFont
  sansItalic: PDFFont
  sansBoldItalic: PDFFont
  serif: PDFFont
  serifBold: PDFFont
  serifItalic: PDFFont
  serifBoldItalic: PDFFont
  mono: PDFFont
  monoBold: PDFFont
}

async function loadFonts(pdfDoc: PDFDocument): Promise<LoadedFonts> {
  const load = async (url: string): Promise<PDFFont> => {
    const res = await fetch(url)
    if (!res.ok) throw new Error(`Could not load font ${url}`)
    return pdfDoc.embedFont(await res.arrayBuffer(), { subset: true })
  }
  const [sans, sansBold, sansItalic, sansBoldItalic, serif, serifBold, serifItalic, serifBoldItalic, mono, monoBold] =
    await Promise.all([
      load(sansUrl),
      load(sansBoldUrl),
      load(sansItalicUrl),
      load(sansBoldItalicUrl),
      load(serifUrl),
      load(serifBoldUrl),
      load(serifItalicUrl),
      load(serifBoldItalicUrl),
      load(monoUrl),
      load(monoBoldUrl),
    ])
  return { sans, sansBold, sansItalic, sansBoldItalic, serif, serifBold, serifItalic, serifBoldItalic, mono, monoBold }
}

function inRanges(ranges: ReadonlyArray<readonly [number, number]>, cp: number): boolean {
  let lo = 0
  let hi = ranges.length - 1
  while (lo <= hi) {
    const mid = (lo + hi) >> 1
    const [a, b] = ranges[mid]
    if (cp < a) hi = mid - 1
    else if (cp > b) lo = mid + 1
    else return true
  }
  return false
}

function italicCoverage(family: FontFamily, bold: boolean): ReadonlyArray<readonly [number, number]> | null {
  if (family === 'sans') return bold ? sansBoldItalicRanges : sansItalicRanges
  if (family === 'serif') return bold ? serifBoldItalicRanges : serifItalicRanges
  return null // mono has no italic face: always fall back to upright
}

function parseHexColor(hex: string | undefined, fallback: RGB): RGB {
  if (!hex) return fallback
  let h = hex.trim().replace(/^#/, '')
  if (/^[0-9a-fA-F]{3}$/.test(h)) h = h.split('').map((c) => c + c).join('')
  if (!/^[0-9a-fA-F]{6}$/.test(h)) return fallback
  const n = parseInt(h, 16)
  return rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255)
}

const BLACK = rgb(0, 0, 0)
const LINK_BLUE = rgb(5 / 255, 99 / 255, 193 / 255)

function marksOf(node: AnyNode): AnyNode[] {
  return Array.isArray(node.marks) ? (node.marks as AnyNode[]) : []
}

function hasMark(node: AnyNode, type: string): AnyNode | undefined {
  return marksOf(node).find((m) => m.type === type)
}

function familyOf(node: AnyNode): FontFamily {
  const ts = hasMark(node, 'textStyle')
  const fam =
    ts?.attrs && typeof (ts.attrs as AnyNode).fontFamily === 'string'
      ? (((ts.attrs as AnyNode).fontFamily as string) || '').toLowerCase()
      : ''
  if (/serif|georgia|times|garamond|palatino|book/.test(fam)) return 'serif'
  if (/mono|courier|consolas|monaco/.test(fam)) return 'mono'
  return 'sans'
}

/** px (as stored by the editor, e.g. "22px") -> pt. */
function fontSizePt(node: AnyNode): number | undefined {
  const ts = hasMark(node, 'textStyle')
  const size =
    ts?.attrs && typeof (ts.attrs as AnyNode).fontSize === 'string'
      ? ((ts.attrs as AnyNode).fontSize as string)
      : undefined
  if (!size) return undefined
  const px = parseFloat(size)
  if (Number.isNaN(px) || px <= 0) return undefined
  return px * 0.75
}

/* ------------------------------------------------------------------ */
/* Atoms: the smallest drawable pieces of a paragraph                 */
/* ------------------------------------------------------------------ */

interface TextAtom {
  kind: 'word' | 'space' | 'break'
  text: string
  font: PDFFont
  size: number
  rise: number
  color: RGB
  underline: boolean
  strike: boolean
  highlight: RGB | null
  width: number
}

interface ImageAtom {
  kind: 'image'
  bytes: Uint8Array
  mime: 'png' | 'jpg'
  natW: number
  natH: number
}

type Atom = TextAtom | ImageAtom

interface Line {
  atoms: Atom[]
  width: number
  maxSize: number
}

function dataUrlToBytes(dataUrl: string): { bytes: Uint8Array; mime: string } | null {
  const match = /^data:(image\/[a-zA-Z+]+);base64,(.+)$/.exec(dataUrl)
  if (!match) return null
  const binary = atob(match[2])
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
  return { bytes, mime: match[1] }
}

function imageDimensions(dataUrl: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    const img = new Image()
    img.onload = () => resolve({ width: img.naturalWidth || 600, height: img.naturalHeight || 400 })
    img.onerror = () => resolve({ width: 600, height: 400 })
    img.src = dataUrl
  })
}

/* ------------------------------------------------------------------ */
/* Writer: cursor, pages, line packing, drawing                       */
/* ------------------------------------------------------------------ */

const BODY_PT = 12
const LINE_FACTOR = 1.4

class Writer {
  pdfDoc: PDFDocument
  fonts: LoadedFonts
  pageW: number
  pageH: number
  mL: number
  mR: number
  mT: number
  mB: number
  contentW: number
  paraAfter: number
  page: PDFPage
  yTop: number
  pageNum = 0

  constructor(pdfDoc: PDFDocument, fonts: LoadedFonts, pageW: number, pageH: number, margins: number, paraAfter: number) {
    this.pdfDoc = pdfDoc
    this.fonts = fonts
    this.pageW = pageW
    this.pageH = pageH
    this.mL = margins
    this.mR = margins
    this.mT = margins
    this.mB = margins
    this.contentW = pageW - margins * 2
    this.paraAfter = paraAfter
    this.page = this.pdfDoc.addPage([pageW, pageH])
    this.pageNum = 1
    this.yTop = pageH - margins
  }

  get contentWidth(): number {
    return this.contentW
  }

  get left(): number {
    return this.mL
  }

  newPage(): void {
    this.page = this.pdfDoc.addPage([this.pageW, this.pageH])
    this.pageNum += 1
    this.yTop = this.pageH - this.mT
  }

  /** Move to a fresh page unless `height` fits in the space left. */
  ensureSpace(height: number): void {
    if (this.yTop - height < this.mB) this.newPage()
  }

  pickFont(family: FontFamily, bold: boolean, italic: boolean, text: string): PDFFont {
    const f = this.fonts
    if (family === 'serif') {
      if (italic) {
        const ranges = italicCoverage('serif', bold)
        if (ranges && [...text].every((ch) => inRanges(ranges, ch.codePointAt(0) as number))) {
          return bold ? f.serifBoldItalic : f.serifItalic
        }
      }
      return bold ? f.serifBold : f.serif
    }
    if (family === 'mono') return bold ? f.monoBold : f.mono
    if (italic) {
      const ranges = italicCoverage('sans', bold)
      if (ranges && [...text].every((ch) => inRanges(ranges, ch.codePointAt(0) as number))) {
        return bold ? f.sansBoldItalic : f.sansItalic
      }
    }
    return bold ? f.sansBold : f.sans
  }

  /** Turn one inline node into atoms (words / spaces / breaks / images). */
  async inlineAtoms(nodes: AnyNode[], baseSize: number, maxWidth = this.contentW): Promise<Atom[]> {
    const out: Atom[] = []
    for (const n of nodes) {
      if (n.type === 'text') {
        const text = typeof n.text === 'string' ? n.text : ''
        if (!text) continue
        const bold = Boolean(hasMark(n, 'bold'))
        const italic = Boolean(hasMark(n, 'italic'))
        const underline = Boolean(hasMark(n, 'underline'))
        const strike = Boolean(hasMark(n, 'strike'))
        const script = hasMark(n, 'subscript') ? 'sub' : hasMark(n, 'superscript') ? 'super' : 'normal'
        const link = hasMark(n, 'link')
        const ts = hasMark(n, 'textStyle')
        const colorHex =
          ts?.attrs && typeof (ts.attrs as AnyNode).color === 'string'
            ? ((ts.attrs as AnyNode).color as string)
            : undefined
        const hl = hasMark(n, 'highlight')
        const hlHex =
          hl?.attrs && typeof (hl.attrs as AnyNode).color === 'string'
            ? ((hl.attrs as AnyNode).color as string)
            : undefined
        const size = fontSizePt(n) ?? baseSize
        const drawSize = script === 'normal' ? size : size * 0.7
        const rise = script === 'super' ? size * 0.3 : script === 'sub' ? -size * 0.12 : 0
        const font = this.pickFont(familyOf(n), bold, italic, text)
        const color = link ? LINK_BLUE : parseHexColor(colorHex, BLACK)
        const highlight = hlHex ? parseHexColor(hlHex, rgb(1, 1, 0)) : null
        // Split on whitespace runs and hard newlines; keep the separators.
        const parts = text.split(/(\s+)/)
        for (const part of parts) {
          if (!part) continue
          if (part.includes('\n')) {
            const segs = part.split('\n')
            segs.forEach((s, i) => {
              if (s) this.pushWord(out, s, font, drawSize, rise, color, underline || Boolean(link), strike, highlight, maxWidth)
              if (i < segs.length - 1) out.push(this.breakAtom(font, drawSize, rise, color))
            })
            continue
          }
          if (/^\s+$/.test(part)) {
            out.push({
              kind: 'space',
              text: ' ',
              font,
              size: drawSize,
              rise,
              color,
              underline: false,
              strike: false,
              highlight: null,
              width: font.widthOfTextAtSize(' ', drawSize),
            })
            continue
          }
          this.pushWord(out, part, font, drawSize, rise, color, underline || Boolean(link), strike, highlight, maxWidth)
        }
      } else if (n.type === 'image') {
        const attrs = (n.attrs ?? {}) as AnyNode
        const src = typeof attrs.src === 'string' ? attrs.src : ''
        const parsed = src.startsWith('data:') ? dataUrlToBytes(src) : null
        if (parsed && (parsed.mime.includes('png') || parsed.mime.includes('jpeg') || parsed.mime.includes('jpg'))) {
          const dims = await imageDimensions(src)
          out.push({
            kind: 'image',
            bytes: parsed.bytes,
            mime: parsed.mime.includes('png') ? 'png' : 'jpg',
            natW: dims.width,
            natH: dims.height,
          })
        }
      } else if (n.type === 'hardBreak') {
        out.push(this.breakAtom(this.fonts.sans, baseSize, 0, BLACK))
      } else if (n.type === 'pageNumber') {
        const t = String(this.pageNum)
        const font = this.fonts.sans
        out.push({
          kind: 'word',
          text: t,
          font,
          size: baseSize,
          rise: 0,
          color: BLACK,
          underline: false,
          strike: false,
          highlight: null,
          width: font.widthOfTextAtSize(t, baseSize),
        })
      }
      // Unknown inline nodes are skipped.
    }
    return out
  }

  private breakAtom(font: PDFFont, size: number, rise: number, color: RGB): TextAtom {
    return { kind: 'break', text: '', font, size, rise, color, underline: false, strike: false, highlight: null, width: 0 }
  }

  /** Push a word, hard-splitting it if it is wider than one line. */
  private pushWord(
    out: Atom[],
    word: string,
    font: PDFFont,
    size: number,
    rise: number,
    color: RGB,
    underline: boolean,
    strike: boolean,
    highlight: RGB | null,
    maxWidth: number,
  ): void {
    const mk = (t: string): TextAtom => ({
      kind: 'word',
      text: t,
      font,
      size,
      rise,
      color,
      underline,
      strike,
      highlight,
      width: font.widthOfTextAtSize(t, size),
    })
    if (font.widthOfTextAtSize(word, size) <= maxWidth) {
      out.push(mk(word))
      return
    }
    let cur = ''
    for (const ch of word) {
      if (cur && font.widthOfTextAtSize(cur + ch, size) > maxWidth) {
        out.push(mk(cur))
        cur = ch
      } else {
        cur += ch
      }
    }
    if (cur) out.push(mk(cur))
  }

  /** Greedy word-wrap. Image atoms always get a line of their own. */
  layoutLines(atoms: Atom[], firstIndent: number, restIndent: number, maxWidth = this.contentW): Line[] {
    const lines: Line[] = []
    let cur: Atom[] = []
    let curW = firstIndent
    let maxSize = BODY_PT
    const flush = () => {
      // Drop a trailing space.
      while (cur.length > 0 && cur[cur.length - 1].kind === 'space') cur.pop()
      if (cur.length > 0) {
        let w = 0
        let ms = BODY_PT
        for (const a of cur) {
          if (a.kind === 'image') continue
          w += a.width
          if (a.size > ms) ms = a.size
        }
        lines.push({ atoms: cur, width: w, maxSize: ms })
      }
      cur = []
      curW = restIndent
      maxSize = BODY_PT
    }
    for (const a of atoms) {
      if (a.kind === 'break') {
        flush()
        continue
      }
      if (a.kind === 'image') {
        flush()
        lines.push({ atoms: [a], width: 0, maxSize: BODY_PT })
        continue
      }
      if (a.kind === 'space' && cur.length === 0) continue // no leading spaces
      if (curW + a.width > maxWidth && cur.length > 0) flush()
      if (a.kind === 'space' && cur.length === 0) continue
      cur.push(a)
      curW += a.width
      if (a.size > maxSize) maxSize = a.size
    }
    flush()
    return lines
  }

  drawTextAtom(a: TextAtom, x: number, baseline: number): void {
    if (a.highlight) {
      const h = a.size * 1.05
      this.page.drawRectangle({
        x,
        y: baseline - a.size * 0.22,
        width: a.width,
        height: h,
        color: a.highlight,
      })
    }
    this.page.drawText(a.text, { x, y: baseline + a.rise, size: a.size, font: a.font, color: a.color })
    if (a.underline) {
      this.page.drawLine({
        start: { x, y: baseline - 1.5 },
        end: { x: x + a.width, y: baseline - 1.5 },
        thickness: 0.7,
        color: a.color,
      })
    }
    if (a.strike) {
      const y = baseline + a.rise + a.size * 0.32
      this.page.drawLine({
        start: { x, y },
        end: { x: x + a.width, y },
        thickness: 0.7,
        color: a.color,
      })
    }
  }

  /** Draw one laid-out line. `lineH` is the line's full height. Image lines are
      drawn via drawImageLine, never through here. */
  drawLine(line: Line, align: 'left' | 'center' | 'right' | 'justify', indent: number, isLast: boolean, lineH: number): void {
    const baseline = this.yTop - lineH + line.maxSize * 0.35
    let x = this.left + indent
    const extra = this.contentW - indent - line.width
    if (align === 'center') x += extra / 2
    else if (align === 'right') x += extra
    if (align === 'justify' && !isLast) {
      const gaps = line.atoms.filter((a) => a.kind === 'space').length
      const perGap = gaps > 0 && extra > 0 ? extra / gaps : 0
      for (const a of line.atoms) {
        if (a.kind !== 'word' && a.kind !== 'space') continue
        if (a.kind === 'space') {
          x += a.width + perGap
          continue
        }
        this.drawTextAtom(a, x, baseline)
        x += a.width
      }
    } else {
      for (const a of line.atoms) {
        if (a.kind !== 'word' && a.kind !== 'space') continue
        if (a.kind === 'space') {
          x += a.width
          continue
        }
        this.drawTextAtom(a, x, baseline)
        x += a.width
      }
    }
    this.yTop -= lineH
  }

  async drawImageLine(img: ImageAtom): Promise<void> {
    const scale = Math.min(1, this.contentW / img.natW)
    const w = img.natW * scale
    const h = img.natH * scale
    this.ensureSpace(h + 4)
    const y = this.yTop - h
    const embedded = img.mime === 'png' ? await this.pdfDoc.embedPng(img.bytes) : await this.pdfDoc.embedJpg(img.bytes)
    this.page.drawImage(embedded, { x: this.left, y, width: w, height: h })
    this.yTop -= h + 4
  }

  /** Draw laid-out lines with per-line height, alignment and spacing. */
  async drawLines(
    lines: Line[],
    opts: {
      align: 'left' | 'center' | 'right' | 'justify'
      indent: number
      lineFactor: number
      spaceBefore?: number
      spaceAfter?: number
    },
  ): Promise<void> {
    if (opts.spaceBefore) {
      this.ensureSpace(opts.spaceBefore)
      this.yTop -= opts.spaceBefore
    }
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i]
      const img = line.atoms.find((a) => a.kind === 'image') as ImageAtom | undefined
      if (img) {
        await this.drawImageLine(img)
        continue
      }
      const lineH = Math.max(line.maxSize * opts.lineFactor, 10)
      this.ensureSpace(lineH)
      this.drawLine(line, opts.align, opts.indent, i === lines.length - 1, lineH)
    }
    this.yTop -= opts.spaceAfter ?? this.paraAfter
  }
}

/* ------------------------------------------------------------------ */
/* Block rendering                                                     */
/* ------------------------------------------------------------------ */

type Align = 'left' | 'center' | 'right' | 'justify'

function alignOf(attrs: AnyNode | undefined): Align {
  const a = attrs?.textAlign as string | undefined
  if (a === 'center') return 'center'
  if (a === 'right') return 'right'
  if (a === 'justify') return 'justify'
  return 'left'
}

function lineFactorOf(attrs: AnyNode | undefined): number {
  const lh = typeof attrs?.lineHeight === 'string' ? parseFloat(attrs.lineHeight as string) : NaN
  if (!Number.isNaN(lh) && lh >= 1 && lh <= 3) return lh
  return LINE_FACTOR
}

function indentPt(attrs: AnyNode | undefined): number {
  const ind = typeof attrs?.indent === 'number' ? (attrs.indent as number) : 0
  return ind > 0 ? ind * 0.9 : 0
}

function childNodes(node: AnyNode): AnyNode[] {
  return Array.isArray(node.content) ? (node.content as AnyNode[]) : []
}

async function renderParagraph(w: Writer, node: AnyNode, indent = 0, baseSize = BODY_PT): Promise<void> {
  const atoms = await w.inlineAtoms(childNodes(node), baseSize)
  const hasText = atoms.some((a) => a.kind === 'word' || a.kind === 'image')
  const attrs = (node.attrs ?? {}) as AnyNode
  const ind = indent + indentPt(attrs)
  if (!hasText) {
    // Blank paragraph: one empty line of space, but never turn the page for it.
    const lh = baseSize * LINE_FACTOR
    if (w.yTop - lh >= w.mB) {
      w.yTop -= lh
      w.yTop -= w.paraAfter
    }
    return
  }
  const lines = w.layoutLines(atoms, ind, ind)
  await w.drawLines(lines, {
    align: alignOf(attrs),
    indent: ind,
    lineFactor: lineFactorOf(attrs),
  })
}

async function renderHeading(w: Writer, node: AnyNode): Promise<void> {
  const attrs = (node.attrs ?? {}) as AnyNode
  const level = (attrs.level as number | undefined) ?? 2
  const size = level === 1 ? 22 : level === 2 ? 17 : 13.5
  const before = level === 1 ? 14 : level === 2 ? 12 : 10
  const after = level === 1 ? 8 : level === 2 ? 6 : 4
  const atoms = await w.inlineAtoms(childNodes(node), size)
  const lines = w.layoutLines(atoms, 0, 0)
  await w.drawLines(lines, {
    align: alignOf(attrs),
    indent: 0,
    lineFactor: 1.25,
    spaceBefore: before,
    spaceAfter: after,
  })
}

async function renderTitle(w: Writer, node: AnyNode): Promise<void> {
  const atoms = await w.inlineAtoms(childNodes(node), 26)
  const lines = w.layoutLines(atoms, 0, 0)
  await w.drawLines(lines, { align: 'center', indent: 0, lineFactor: 1.25, spaceAfter: 10 })
}

async function renderSubtitle(w: Writer, node: AnyNode): Promise<void> {
  const atoms = await w.inlineAtoms(childNodes(node), 15)
  const lines = w.layoutLines(atoms, 0, 0)
  await w.drawLines(lines, { align: 'center', indent: 0, lineFactor: 1.3, spaceAfter: 10 })
}

async function renderList(w: Writer, node: AnyNode, level: number, ordered: boolean): Promise<void> {
  let n = 0
  for (const item of childNodes(node)) {
    if (item.type !== 'listItem') continue
    n += 1
    const marker = ordered ? `${n}.` : '•'
    for (const sub of childNodes(item)) {
      if (sub.type === 'paragraph') {
        const atoms = await w.inlineAtoms(childNodes(sub), BODY_PT)
        const markerAtom: TextAtom = {
          kind: 'word',
          text: marker,
          font: w.fonts.sansBold,
          size: BODY_PT,
          rise: 0,
          color: BLACK,
          underline: false,
          strike: false,
          highlight: null,
          width: w.fonts.sansBold.widthOfTextAtSize(marker, BODY_PT),
        }
        const spaceAtom: TextAtom = {
          kind: 'space',
          text: ' ',
          font: w.fonts.sans,
          size: BODY_PT,
          rise: 0,
          color: BLACK,
          underline: false,
          strike: false,
          highlight: null,
          width: w.fonts.sans.widthOfTextAtSize(' ', BODY_PT),
        }
        const firstIndent = level * 18
        const restIndent = level * 18 + 14
        const lines = w.layoutLines([markerAtom, spaceAtom, ...atoms], firstIndent, restIndent)
        await w.drawLines(lines, {
          align: alignOf((sub.attrs ?? {}) as AnyNode),
          indent: 0,
          lineFactor: LINE_FACTOR,
        })
      } else if (sub.type === 'bulletList' || sub.type === 'orderedList') {
        await renderList(w, sub, level + 1, sub.type === 'orderedList')
      }
    }
  }
}

async function renderTable(w: Writer, node: AnyNode): Promise<void> {
  const rows = childNodes(node).filter((r) => r.type === 'tableRow')
  if (rows.length === 0) return
  const grid: AnyNode[][][] = rows.map((row) =>
    childNodes(row)
      .filter((c) => c.type === 'tableCell' || c.type === 'tableHeader')
      .map((cell) => childNodes(cell).filter((p) => p.type === 'paragraph')),
  )
  const colCount = Math.max(...grid.map((r) => r.length), 1)
  // Column widths: proportional to each column's longest word, min 36pt.
  const colW: number[] = new Array(colCount).fill(36)
  for (const row of grid) {
    row.forEach((cellParas, ci) => {
      for (const p of cellParas) {
        for (const t of childNodes(p)) {
          if (t.type !== 'text' || typeof t.text !== 'string') continue
          for (const word of (t.text as string).split(/\s+/)) {
            if (!word) continue
            const f = w.pickFont(familyOf(t), Boolean(hasMark(t, 'bold')), Boolean(hasMark(t, 'italic')), word)
            const size = fontSizePt(t) ?? 11
            colW[ci] = Math.max(colW[ci], f.widthOfTextAtSize(word, size) + 8)
          }
        }
      }
    })
  }
  const total = colW.reduce((a, b) => a + b, 0)
  const avail = w.contentWidth
  const scale = total > avail ? avail / total : 1
  for (let i = 0; i < colCount; i++) colW[i] *= scale
  // A single-column table has nothing to balance against: fill the content
  // width, matching the editor (table width: 100%), instead of shrinking to
  // the longest word (which made a 1x1 table a ~100pt strip).
  if (colCount === 1) colW[0] = avail

  const PAD = 4
  const BORDER = rgb(0.6, 0.6, 0.6)
  const HEADER_FILL = rgb(217 / 255, 226 / 255, 243 / 255)
  const usableH = w.pageH - w.mT - w.mB
  const lineH = (l: Line) => Math.max(l.maxSize * LINE_FACTOR, 10)
  for (let ri = 0; ri < grid.length; ri++) {
    const row = grid[ri]
    const isHeader = childNodes(rows[ri]).some((c) => c.type === 'tableHeader')
    // Lay out every cell against its real column width to get the row height.
    const cellLines: Line[][] = []
    let rowH = 0
    for (let ci = 0; ci < colCount; ci++) {
      const paras = row[ci] ?? []
      const lines: Line[] = []
      for (const p of paras) {
        const atoms = await w.inlineAtoms(childNodes(p), 11, colW[ci] - PAD * 2)
        lines.push(...w.layoutLines(atoms, 0, 0, colW[ci] - PAD * 2))
      }
      if (lines.length === 0) lines.push({ atoms: [], width: 0, maxSize: 11 })
      cellLines.push(lines)
      const h = lines.reduce((sum, l) => sum + lineH(l), 0) + PAD * 2
      rowH = Math.max(rowH, h)
    }
    // Draw one page-segment of this row: take[ci] lines of each cell starting
    // at from[ci], inside a box of height segH whose top is segTop. The top
    // padding applies only on the row's first segment; the bottom padding is
    // already included in segH when the tallest cell finishes in this segment.
    // Every segment gets its own full border, so each page shows a closed box.
    const drawRowSegment = (from: number[], take: number[], segTop: number, segH: number, first: boolean) => {
      let x = w.left
      for (let ci = 0; ci < colCount; ci++) {
        const cw = colW[ci]
        if (isHeader) {
          w.page.drawRectangle({ x, y: segTop - segH, width: cw, height: segH, color: HEADER_FILL })
        }
        let ly = segTop - (first ? PAD : 0)
        for (let k = 0; k < take[ci]; k++) {
          const line = cellLines[ci][from[ci] + k]
          const lh = lineH(line)
          const baseline = ly - lh + line.maxSize * 0.35
          let lx = x + PAD
          for (const a of line.atoms) {
            if (a.kind !== 'word' && a.kind !== 'space') continue
            if (a.kind === 'space') {
              lx += a.width
              continue
            }
            w.drawTextAtom(a, lx, baseline)
            lx += a.width
          }
          ly -= lh
        }
        w.page.drawLine({ start: { x, y: segTop }, end: { x: x + cw, y: segTop }, thickness: 0.5, color: BORDER })
        w.page.drawLine({ start: { x, y: segTop - segH }, end: { x: x + cw, y: segTop - segH }, thickness: 0.5, color: BORDER })
        w.page.drawLine({ start: { x, y: segTop }, end: { x, y: segTop - segH }, thickness: 0.5, color: BORDER })
        w.page.drawLine({ start: { x: x + cw, y: segTop }, end: { x: x + cw, y: segTop - segH }, thickness: 0.5, color: BORDER })
        x += cw
      }
    }
    if (rowH <= usableH) {
      // Normal row: fits on one page (move to a fresh page if needed).
      w.ensureSpace(rowH)
      drawRowSegment(
        cellLines.map(() => 0),
        cellLines.map((ls) => ls.length),
        w.yTop,
        rowH,
        true,
      )
      w.yTop -= rowH
      continue
    }
    // A row taller than a whole page: split it across pages, drawing the
    // cell border around every page segment so each page shows a closed box
    // and no content is lost (previously it was drawn as-is off the page).
    const from = cellLines.map(() => 0)
    let first = true
    for (;;) {
      // Barely any room left on this page: start the segment on a fresh page
      // instead of drawing a sliver box.
      if (w.yTop - w.mB < PAD * 2 + 16) w.newPage()
      const segTop = w.yTop
      const segAvail = segTop - w.mB
      const take: number[] = []
      let segH = 0
      for (let ci = 0; ci < colCount; ci++) {
        const lines = cellLines[ci]
        const remaining = lines.length - from[ci]
        const topPad = first ? PAD : 0
        const budget = segAvail - topPad
        const restH = lines.slice(from[ci]).reduce((s, l) => s + lineH(l), 0)
        let n: number
        let sum: number
        if (restH + PAD <= budget) {
          // This cell finishes in this segment: include the bottom padding.
          n = remaining
          sum = restH
          segH = Math.max(segH, topPad + sum + PAD)
        } else {
          n = 0
          sum = 0
          while (n < remaining && sum + lineH(lines[from[ci] + n]) <= budget) {
            sum += lineH(lines[from[ci] + n])
            n++
          }
          if (n === 0) {
            // A single line taller than the budget: force progress (a line can
            // never exceed a whole page, so this cannot loop forever).
            n = 1
            sum = lineH(lines[from[ci]])
          }
          segH = Math.max(segH, topPad + sum)
        }
        take.push(n)
      }
      drawRowSegment(from, take, segTop, segH, first)
      for (let ci = 0; ci < colCount; ci++) from[ci] += take[ci]
      w.yTop = segTop - segH
      if (from.every((f, ci) => f >= cellLines[ci].length)) break
      w.newPage()
      first = false
    }
  }
  w.yTop -= 6
}

async function renderBlock(w: Writer, node: AnyNode): Promise<void> {
  switch (node.type) {
    case 'paragraph':
      await renderParagraph(w, node)
      return
    case 'heading':
      await renderHeading(w, node)
      return
    case 'title':
      await renderTitle(w, node)
      return
    case 'subtitle':
      await renderSubtitle(w, node)
      return
    case 'bulletList':
      await renderList(w, node, 0, false)
      return
    case 'orderedList':
      await renderList(w, node, 0, true)
      return
    case 'table':
      await renderTable(w, node)
      return
    case 'horizontalRule': {
      w.ensureSpace(12)
      const y = w.yTop - 6
      w.page.drawLine({
        start: { x: w.left, y },
        end: { x: w.left + w.contentWidth, y },
        thickness: 0.75,
        color: rgb(0.5, 0.5, 0.5),
      })
      w.yTop -= 12
      return
    }
    case 'pageBreak':
      w.newPage()
      return
    case 'blockquote': {
      const before = w.yTop
      for (const sub of childNodes(node)) {
        await renderBlockIndented(w, sub)
      }
      const after = w.yTop
      w.page.drawLine({
        start: { x: w.left + 6, y: before },
        end: { x: w.left + 6, y: after },
        thickness: 1.5,
        color: rgb(0.6, 0.6, 0.6),
      })
      return
    }
    default:
      return
  }
}

async function renderBlockIndented(w: Writer, node: AnyNode): Promise<void> {
  if (node.type === 'paragraph') {
    await renderParagraph(w, node, 18)
    return
  }
  await renderBlock(w, node)
}

/* ------------------------------------------------------------------ */
/* Entry point                                                         */
/* ------------------------------------------------------------------ */

export async function exportPdf(doc: MyDocument): Promise<Blob> {
  const pdfDoc = await PDFDocument.create()
  pdfDoc.registerFontkit(fontkit)
  const fonts = await loadFonts(pdfDoc)

  const { size, orientation } = doc.pageSettings
  const dims = pageDims(size, orientation)
  const pageW = dims.width * 0.75
  const pageH = dims.height * 0.75
  const margins =
    doc.pageSettings.margins === 'narrow' ? 36 : doc.pageSettings.margins === 'wide' ? 144 : 72
  const paraAfter =
    doc.pageSettings.paragraphSpacing === 'compact' ? 4 : doc.pageSettings.paragraphSpacing === 'relaxed' ? 10 : 6

  const w = new Writer(pdfDoc, fonts, pageW, pageH, margins, paraAfter)

  const topNodes: AnyNode[] = Array.isArray((doc.content as AnyNode).content)
    ? ((doc.content as AnyNode).content as AnyNode[])
    : []
  for (const node of topNodes) {
    await renderBlock(w, node)
  }

  // Footnotes / endnotes / bibliography (same order as the DOCX export).
  if (doc.footnotes.length > 0) {
    await renderHeading(w, { type: 'heading', attrs: { level: 2 }, content: [{ type: 'text', text: 'Footnotes' }] })
    for (let i = 0; i < doc.footnotes.length; i++) {
      await renderParagraph(w, {
        type: 'paragraph',
        content: [{ type: 'text', text: `[${i + 1}] ${doc.footnotes[i].text}` }],
      })
    }
  }
  if (doc.endnotes.length > 0) {
    await renderHeading(w, { type: 'heading', attrs: { level: 2 }, content: [{ type: 'text', text: 'Endnotes' }] })
    for (let i = 0; i < doc.endnotes.length; i++) {
      await renderParagraph(w, {
        type: 'paragraph',
        content: [{ type: 'text', text: `[${i + 1}] ${doc.endnotes[i].text}` }],
      })
    }
  }
  if (doc.citations.length > 0) {
    await renderHeading(w, { type: 'heading', attrs: { level: 2 }, content: [{ type: 'text', text: 'Bibliography' }] })
    for (const c of doc.citations) {
      await renderParagraph(w, {
        type: 'paragraph',
        content: [{ type: 'text', text: `${c.author} (${c.year}). ${c.title}.` }],
      })
    }
  }

  const bytes = await pdfDoc.save()
  return new Blob([bytes as unknown as BlobPart], { type: 'application/pdf' })
}
