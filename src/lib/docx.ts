import * as mammoth from 'mammoth'
import DOMPurify from 'dompurify'
import {
  AlignmentType,
  BorderStyle,
  Document,
  ExternalHyperlink,
  HeadingLevel,
  ImageRun,
  LineRuleType,
  NumberFormat,
  Packer,
  PageBreak,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
  type IParagraphOptions,
  type ISectionOptions,
} from 'docx'
import type { MyDocument } from './types'
import { pageDims } from './types'

/* ------------------------------------------------------------------ */
/* DOCX import (mammoth -> sanitized HTML)                             */
/* ------------------------------------------------------------------ */

export async function importDocxFile(file: File): Promise<{ html: string; title: string }> {
  const arrayBuffer = await file.arrayBuffer()
  const result = await mammoth.convertToHtml(
    { arrayBuffer },
    {
      convertImage: mammoth.images.imgElement((image) =>
        image.read('base64').then((base64: string) => ({
          src: `data:${image.contentType};base64,${base64}`,
        })),
      ),
    },
  )
  const clean = DOMPurify.sanitize(result.value, {
    ADD_TAGS: ['img'],
    ADD_ATTR: ['src', 'alt', 'width', 'height'],
  })
  const title = file.name.replace(/\.docx$/i, '')
  return { html: clean, title }
}

/* ------------------------------------------------------------------ */
/* DOCX export (Tiptap JSON -> real .docx via the `docx` library)      */
/* ------------------------------------------------------------------ */

const PT_TO_TWIPS = 20

function twips(pt: number): number {
  return Math.round(pt * PT_TO_TWIPS)
}

function imageDimensions(dataUrl: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    const img = new Image()
    img.onload = () => resolve({ width: img.naturalWidth || 600, height: img.naturalHeight || 400 })
    img.onerror = () => resolve({ width: 600, height: 400 })
    img.src = dataUrl
  })
}

function dataUrlToBytes(dataUrl: string): { bytes: Uint8Array; mime: string } | null {
  const match = /^data:(image\/[a-zA-Z+]+);base64,(.+)$/.exec(dataUrl)
  if (!match) return null
  const binary = atob(match[2])
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
  return { bytes, mime: match[1] }
}

type AnyNode = Record<string, unknown>

function marksOf(node: AnyNode): AnyNode[] {
  return Array.isArray(node.marks) ? (node.marks as AnyNode[]) : []
}

function hasMark(node: AnyNode, type: string): AnyNode | undefined {
  return marksOf(node).find((m) => m.type === type)
}

/** docx v9 only accepts named highlight colours — map common hex values, drop the rest. */
type HighlightColor =
  | 'yellow'
  | 'green'
  | 'cyan'
  | 'magenta'
  | 'blue'
  | 'red'
  | 'darkBlue'
  | 'darkCyan'
  | 'darkGreen'
  | 'darkMagenta'
  | 'darkRed'
  | 'darkYellow'
  | 'lightGray'
  | 'darkGray'
  | 'black'
  | 'white'
  | 'none'

function toHighlightColor(hex: string | undefined): HighlightColor | undefined {
  if (!hex) return undefined
  const h = hex.toLowerCase().replace('#', '')
  if (['ffff00', 'fde047', 'fef08a', 'facc15', 'eab308'].includes(h)) return 'yellow'
  if (['00ff00', '22c55e', '4ade80', '86efac'].includes(h)) return 'green'
  if (['00ffff', '22d3ee', '67e8f9'].includes(h)) return 'cyan'
  if (['ff00ff', 'e879f9', 'd946ef'].includes(h)) return 'magenta'
  if (['0000ff', '3b82f6', '60a5fa'].includes(h)) return 'blue'
  if (['ff0000', 'ef4444', 'f87171'].includes(h)) return 'red'
  return undefined
}

function fontSizeHalfPoints(node: AnyNode): number | undefined {
  const ts = hasMark(node, 'textStyle')
  const size = ts?.attrs && typeof (ts.attrs as AnyNode).fontSize === 'string'
    ? ((ts.attrs as AnyNode).fontSize as string)
    : undefined
  if (!size) return undefined
  const px = parseFloat(size)
  if (Number.isNaN(px)) return undefined
  return Math.round(px * 1.5) // px -> pt (x0.75) -> half-points (x2)
}

function textRun(node: AnyNode): TextRun | ExternalHyperlink {
  const text = typeof node.text === 'string' ? node.text : ''
  const link = hasMark(node, 'link')
  const color = hasMark(node, 'textStyle')
  const colorHex =
    color?.attrs && typeof (color.attrs as AnyNode).color === 'string'
      ? ((color.attrs as AnyNode).color as string).replace('#', '')
      : undefined
  const fontFamily =
    color?.attrs && typeof (color.attrs as AnyNode).fontFamily === 'string'
      ? ((color.attrs as AnyNode).fontFamily as string)
      : undefined
  const highlight = hasMark(node, 'highlight')
  const highlightColor =
    highlight?.attrs && typeof (highlight.attrs as AnyNode).color === 'string'
      ? ((highlight.attrs as AnyNode).color as string)
      : undefined
  const run = new TextRun({
    text,
    bold: Boolean(hasMark(node, 'bold')),
    italics: Boolean(hasMark(node, 'italic')),
    underline: hasMark(node, 'underline') ? {} : undefined,
    strike: Boolean(hasMark(node, 'strike')),
    subScript: Boolean(hasMark(node, 'subscript')),
    superScript: Boolean(hasMark(node, 'superscript')),
    color: colorHex,
    font: fontFamily,
    size: fontSizeHalfPoints(node),
    highlight: toHighlightColor(highlightColor),
  })
  if (link?.attrs && typeof (link.attrs as AnyNode).href === 'string') {
    return new ExternalHyperlink({
      link: (link.attrs as AnyNode).href as string,
      children: [run],
    })
  }
  return run
}

function alignmentOf(attrs: AnyNode | undefined) {
  const a = attrs?.textAlign as string | undefined
  if (a === 'center') return AlignmentType.CENTER
  if (a === 'right') return AlignmentType.RIGHT
  if (a === 'justify') return AlignmentType.JUSTIFIED
  return undefined
}

function paragraphOptions(node: AnyNode): IParagraphOptions {
  const attrs = (node.attrs ?? {}) as AnyNode
  const lineHeight = typeof attrs.lineHeight === 'string' ? parseFloat(attrs.lineHeight) : undefined
  return {
    alignment: alignmentOf(attrs),
    spacing: lineHeight
      ? { line: Math.round(lineHeight * 240), lineRule: LineRuleType.AUTO }
      : undefined,
    indent:
      typeof attrs.indent === 'number' && attrs.indent > 0
        ? { left: twips(attrs.indent * 18) }
        : undefined,
  }
}

async function inlineChildren(
  nodes: AnyNode[],
): Promise<(TextRun | ImageRun | ExternalHyperlink)[]> {
  const out: (TextRun | ImageRun | ExternalHyperlink)[] = []
  for (const n of nodes) {
    if (n.type === 'text') out.push(textRun(n))
    else if (n.type === 'image') {
      const attrs = (n.attrs ?? {}) as AnyNode
      const src = typeof attrs.src === 'string' ? attrs.src : ''
      const parsed = src.startsWith('data:') ? dataUrlToBytes(src) : null
      if (parsed) {
        const dims = await imageDimensions(src)
        const scale = Math.min(1, 560 / dims.width)
        const type = parsed.mime.includes('png')
          ? 'png'
          : parsed.mime.includes('gif')
            ? 'gif'
            : 'jpg'
        out.push(
          new ImageRun({
            data: parsed.bytes,
            transformation: {
              width: Math.round(dims.width * scale),
              height: Math.round(dims.height * scale),
            },
            type,
          }),
        )
      }
    } else if (n.type === 'hardBreak') {
      out.push(new TextRun({ text: '', break: 1 }))
    }
  }
  return out
}

async function convertBlock(
  node: AnyNode,
  listLevel = 0,
): Promise<(Paragraph | Table)[]> {
  const children = Array.isArray(node.content) ? (node.content as AnyNode[]) : []
  switch (node.type) {
    case 'paragraph': {
      const runs = await inlineChildren(children)
      return [new Paragraph({ ...paragraphOptions(node), children: runs })]
    }
    case 'heading': {
      const attrs = (node.attrs ?? {}) as AnyNode
      const level = attrs.level as number | undefined
      const runs = await inlineChildren(children)
      if (level === 1) return [new Paragraph({ heading: HeadingLevel.HEADING_1, children: runs })]
      if (level === 2) return [new Paragraph({ heading: HeadingLevel.HEADING_2, children: runs })]
      if (level === 3) return [new Paragraph({ heading: HeadingLevel.HEADING_3, children: runs })]
      return [new Paragraph({ ...paragraphOptions(node), children: runs })]
    }
    case 'bulletList': {
      const out: Paragraph[] = []
      for (const item of children) {
        if (item.type !== 'listItem') continue
        const itemChildren = Array.isArray(item.content) ? (item.content as AnyNode[]) : []
        for (const sub of itemChildren) {
          if (sub.type === 'paragraph') {
            const runs = await inlineChildren(
              Array.isArray(sub.content) ? (sub.content as AnyNode[]) : [],
            )
            out.push(
              new Paragraph({
                ...paragraphOptions(sub),
                bullet: { level: listLevel },
                children: runs,
              }),
            )
          } else if (sub.type === 'bulletList' || sub.type === 'orderedList') {
            for (const b of await convertBlock(sub, listLevel + 1)) {
              if (b instanceof Paragraph) out.push(b)
            }
          }
        }
      }
      return out
    }
    case 'orderedList': {
      const out: Paragraph[] = []
      for (const item of children) {
        if (item.type !== 'listItem') continue
        const itemChildren = Array.isArray(item.content) ? (item.content as AnyNode[]) : []
        for (const sub of itemChildren) {
          if (sub.type === 'paragraph') {
            const runs = await inlineChildren(
              Array.isArray(sub.content) ? (sub.content as AnyNode[]) : [],
            )
            out.push(
              new Paragraph({
                ...paragraphOptions(sub),
                numbering: { reference: 'my-word-numbering', level: listLevel },
                children: runs,
              }),
            )
          } else if (sub.type === 'bulletList' || sub.type === 'orderedList') {
            for (const b of await convertBlock(sub, listLevel + 1)) {
              if (b instanceof Paragraph) out.push(b)
            }
          }
        }
      }
      return out
    }
    case 'table': {
      const rows: TableRow[] = []
      for (const row of children) {
        if (row.type !== 'tableRow') continue
        const cells: TableCell[] = []
        for (const cell of Array.isArray(row.content) ? (row.content as AnyNode[]) : []) {
          if (cell.type !== 'tableCell' && cell.type !== 'tableHeader') continue
          const cellParas: Paragraph[] = []
          for (const p of Array.isArray(cell.content) ? (cell.content as AnyNode[]) : []) {
            // Table cells hold paragraphs only; nested tables are dropped.
            for (const b of await convertBlock(p)) {
              if (b instanceof Paragraph) cellParas.push(b)
            }
          }
          const isHeader = cell.type === 'tableHeader'
          cells.push(
            new TableCell({
              children: cellParas.length ? cellParas : [new Paragraph({ children: [] })],
              shading: isHeader
                ? { type: ShadingType.CLEAR, color: 'auto', fill: 'D9E2F3' }
                : undefined,
            }),
          )
        }
        rows.push(new TableRow({ children: cells }))
      }
      return [
        new Table({
          rows,
          width: { size: 100, type: WidthType.PERCENTAGE },
        }),
      ]
    }
    case 'horizontalRule':
      return [
        new Paragraph({
          children: [],
          border: { bottom: { color: '808080', style: BorderStyle.SINGLE, size: 6, space: 1 } },
        }),
      ]
    case 'pageBreak':
      return [new Paragraph({ children: [new PageBreak()] })]
    case 'blockquote': {
      const out: Paragraph[] = []
      for (const sub of children) {
        // Blockquotes hold paragraphs only; nested tables are dropped.
        for (const b of await convertBlock(sub)) {
          if (b instanceof Paragraph) out.push(b)
        }
      }
      return out
    }
    default:
      return []
  }
}

export async function exportDocx(doc: MyDocument): Promise<Blob> {
  const { size, orientation } = doc.pageSettings
  const dims = pageDims(size, orientation)
  // px @96dpi -> twips (1px = 15 twips)
  const pageWidthTwips = dims.width * 15
  const pageHeightTwips = dims.height * 15
  const marginTwips =
    doc.pageSettings.margins === 'narrow' ? 720 : doc.pageSettings.margins === 'wide' ? 2880 : 1440

  const topNodes: AnyNode[] = Array.isArray((doc.content as AnyNode).content)
    ? ((doc.content as AnyNode).content as AnyNode[])
    : []

  const children: (Paragraph | Table)[] = []
  for (const node of topNodes) {
    // Title / Subtitle are custom node types from the editor's style dropdown
    if (node.type === 'title' || node.type === 'subtitle') {
      const runs = await inlineChildren(
        Array.isArray(node.content) ? (node.content as AnyNode[]) : [],
      )
      children.push(
        new Paragraph({
          heading: node.type === 'title' ? HeadingLevel.TITLE : undefined,
          children: runs,
          alignment: AlignmentType.CENTER,
        }),
      )
      continue
    }
    children.push(...(await convertBlock(node)))
  }

  if (doc.footnotes.length > 0) {
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('Footnotes')] }))
    doc.footnotes.forEach((fn, i) => {
      children.push(new Paragraph({ children: [new TextRun(`[${i + 1}] ${fn.text}`)] }))
    })
  }
  if (doc.endnotes.length > 0) {
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('Endnotes')] }))
    doc.endnotes.forEach((en, i) => {
      children.push(new Paragraph({ children: [new TextRun(`[${i + 1}] ${en.text}`)] }))
    })
  }
  if (doc.citations.length > 0) {
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('Bibliography')] }))
    doc.citations.forEach((c) => {
      children.push(
        new Paragraph({ children: [new TextRun(`${c.author} (${c.year}). ${c.title}.`)] }),
      )
    })
  }

  const section: ISectionOptions = {
    properties: {
      page: {
        size: { width: pageWidthTwips, height: pageHeightTwips },
        margin: {
          top: marginTwips,
          bottom: marginTwips,
          left: marginTwips,
          right: marginTwips,
        },
      },
    },
    children,
  }

  const docxDoc = new Document({
    numbering: {
      config: [
        {
          reference: 'my-word-numbering',
          levels: Array.from({ length: 9 }, (_, level) => ({
            level,
            format: NumberFormat.DECIMAL,
            text: '%1.',
            alignment: AlignmentType.LEFT,
            style: {
              paragraph: { indent: { left: twips(18 * (level + 1)), hanging: twips(18) } },
            },
          })),
        },
      ],
    },
    sections: [section],
  })

  return Packer.toBlob(docxDoc)
}
