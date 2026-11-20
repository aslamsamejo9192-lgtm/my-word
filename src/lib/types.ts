export type DocOrigin = 'blank' | 'docx'

export interface Footnote {
  id: string
  text: string
}

export interface Citation {
  id: string
  author: string
  year: string
  title: string
}

export interface DocComment {
  id: string
  text: string
  quote: string
  from: number
  to: number
}

export interface PageSettings {
  margins: 'narrow' | 'normal' | 'wide'
  orientation: 'portrait' | 'landscape'
  size: 'A4' | 'A3' | 'Letter' | 'Legal'
  columns: 1 | 2 | 3
  border: boolean
  background: string
  paragraphSpacing: 'compact' | 'normal' | 'relaxed'
  indent: number
  view: 'print' | 'web'
  showRuler: boolean
  zoom: number
  twoPage: boolean
}

export interface MyDocument {
  id: string
  title: string
  /** Tiptap JSON document */
  content: Record<string, unknown>
  html: string
  headerText: string
  footerText: string
  showHeaderFooter: boolean
  footnotes: Footnote[]
  endnotes: Footnote[]
  citations: Citation[]
  comments: DocComment[]
  pageSettings: PageSettings
  origin: DocOrigin
  originalFileName?: string
  originalFileBlob?: Blob
  createdAt: number
  updatedAt: number
}

export const DEFAULT_PAGE_SETTINGS: PageSettings = {
  margins: 'normal',
  orientation: 'portrait',
  size: 'A4',
  columns: 1,
  border: false,
  background: '#c8ccd0',
  paragraphSpacing: 'normal',
  indent: 0,
  view: 'print',
  showRuler: true,
  zoom: 100,
  twoPage: false,
}

/** Page size in CSS px at 96dpi, portrait. */
export function pageDims(
  size: PageSettings['size'],
  orientation: PageSettings['orientation'],
): { width: number; height: number } {
  const sizes: Record<PageSettings['size'], { w: number; h: number }> = {
    A4: { w: 794, h: 1123 },
    A3: { w: 1123, h: 1587 },
    Letter: { w: 816, h: 1056 },
    Legal: { w: 816, h: 1344 },
  }
  const s = sizes[size]
  return orientation === 'portrait' ? { width: s.w, height: s.h } : { width: s.h, height: s.w }
}

export function marginPx(margins: PageSettings['margins']): number {
  return margins === 'narrow' ? 36 : margins === 'wide' ? 96 : 72
}

export function uid(): string {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 9)}`
}

export function blankContent(): Record<string, unknown> {
  return {
    type: 'doc',
    content: [{ type: 'paragraph', content: [{ type: 'text', text: '' }] }],
  }
}

export function newDocument(title = 'Untitled Document'): MyDocument {
  const now = Date.now()
  return {
    id: uid(),
    title,
    content: blankContent(),
    html: '',
    headerText: '',
    footerText: '',
    showHeaderFooter: false,
    footnotes: [],
    endnotes: [],
    citations: [],
    comments: [],
    pageSettings: { ...DEFAULT_PAGE_SETTINGS },
    origin: 'blank',
    createdAt: now,
    updatedAt: now,
  }
}
