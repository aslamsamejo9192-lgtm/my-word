import { useCallback, useEffect, useRef, useState } from 'react'
import * as pdfjsLib from 'pdfjs-dist'
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url'
import {
  PDFDocument,
  StandardFonts,
  degrees,
  rgb,
  type PDFPage,
  type RGB,
} from 'pdf-lib'
import {
  ArrowDown,
  ArrowUp,
  ChevronLeft,
  ChevronRight,
  Circle,
  Download,
  FileOutput,
  FilePlus2,
  FileText,
  FolderOpen,
  Hash,
  Highlighter,
  List,
  Maximize,
  MessageSquare,
  MoveHorizontal,
  PanelLeft,
  Pen,
  RotateCw,
  Scissors,
  Search,
  Square,
  Stamp,
  Strikethrough,
  Trash2,
  Type,
  Underline,
  Undo2,
  Redo2,
  X,
  ZoomIn,
  ZoomOut,
} from 'lucide-react'

pdfjsLib.GlobalWorkerOptions.workerSrc = workerUrl

export interface PdfWorkspaceProps {
  initialFile?: File | null
  onClose: () => void
}

type AnnotationType =
  | 'highlight'
  | 'underline'
  | 'strike'
  | 'comment'
  | 'textbox'
  | 'pen'
  | 'rect'
  | 'ellipse'

interface Annotation {
  id: string
  page: number // 0-based page index
  type: AnnotationType
  x: number // fractions of page width/height (displayed orientation)
  y: number
  w: number
  h: number
  points?: { x: number; y: number }[]
  text?: string
  color: string
}

type Tool = 'select' | AnnotationType

const MAX_FILE_BYTES = 50 * 1024 * 1024
const PALETTE = ['#facc15', '#22c55e', '#3b82f6', '#ef4444', '#f97316', '#a855f7', '#111827']

const TYPE_LABELS: Record<AnnotationType, string> = {
  highlight: 'Highlight',
  underline: 'Underline',
  strike: 'Strikethrough',
  comment: 'Comment',
  textbox: 'Text box',
  pen: 'Freehand',
  rect: 'Rectangle',
  ellipse: 'Ellipse',
}

let uidCounter = 0
function uid(): string {
  uidCounter += 1
  return `ann-${Date.now().toString(36)}-${uidCounter}`
}

function hexToRgb(hex: string): RGB {
  const h = hex.replace('#', '')
  const n = parseInt(h.length === 3 ? h.split('').map((c) => c + c).join('') : h, 16)
  return rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255)
}

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  document.body.appendChild(a)
  a.click()
  a.remove()
  window.setTimeout(() => URL.revokeObjectURL(url), 1000)
}

/** Convert a displayed-space fraction rect to unrotated PDF points (bottom-left origin). */
function fracRectToPdf(
  fx: number,
  fy: number,
  fw: number,
  fh: number,
  pageW: number,
  pageH: number,
  rotation: number,
): { x: number; y: number; w: number; h: number } {
  const rot = ((rotation % 360) + 360) % 360
  if (rot === 90) return { x: fy * pageW, y: fx * pageH, w: fh * pageW, h: fw * pageH }
  if (rot === 180) return { x: (1 - fx - fw) * pageW, y: fy * pageH, w: fw * pageW, h: fh * pageH }
  if (rot === 270) return { x: (1 - fy - fh) * pageW, y: (1 - fx - fw) * pageH, w: fh * pageW, h: fw * pageH }
  return { x: fx * pageW, y: (1 - fy - fh) * pageH, w: fw * pageW, h: fh * pageH }
}

function fracPointToPdf(
  fx: number,
  fy: number,
  pageW: number,
  pageH: number,
  rotation: number,
): { x: number; y: number } {
  const r = fracRectToPdf(fx, fy, 0, 0, pageW, pageH, rotation)
  return { x: r.x, y: r.y }
}

/** Rotate an annotation's fractions 90° clockwise (page rotated 90° CW). */
function rotateAnnotation90CW(a: Annotation): Annotation {
  const pts = a.points?.map((p) => ({ x: 1 - p.y, y: p.x }))
  return { ...a, x: 1 - (a.y + a.h), y: a.x, w: a.h, h: a.w, points: pts }
}

function wrapText(text: string, font: { widthOfTextAtSize(t: string, s: number): number }, size: number, maxWidth: number): string[] {
  const lines: string[] = []
  for (const raw of text.split('\n')) {
    const words = raw.split(/\s+/).filter(Boolean)
    if (words.length === 0) {
      lines.push('')
      continue
    }
    let line = words[0]
    for (const word of words.slice(1)) {
      const trial = `${line} ${word}`
      if (font.widthOfTextAtSize(trial, size) <= maxWidth) line = trial
      else {
        lines.push(line)
        line = word
      }
    }
    lines.push(line)
  }
  return lines.length ? lines : ['']
}

export default function PdfWorkspace({ initialFile, onClose }: PdfWorkspaceProps) {
  const [pdfBytes, setPdfBytes] = useState<Uint8Array | null>(null)
  const [fileName, setFileName] = useState('')
  const [pdfDoc, setPdfDoc] = useState<pdfjsLib.PDFDocumentProxy | null>(null)
  const [numPages, setNumPages] = useState(0)
  const [currentPage, setCurrentPage] = useState(1)
  const [pageRotations, setPageRotations] = useState<number[]>([])
  const [scale, setScale] = useState(1)
  const [pageSize, setPageSize] = useState<{ w: number; h: number } | null>(null)
  const [error, setError] = useState('')
  const [busy, setBusy] = useState('')
  const [tool, setTool] = useState<Tool>('select')
  const [color, setColor] = useState('#facc15')
  const [annotations, setAnnotations] = useState<Annotation[]>([])
  const [undoStack, setUndoStack] = useState<Annotation[][]>([])
  const [redoStack, setRedoStack] = useState<Annotation[][]>([])
  const [draft, setDraft] = useState<Annotation | null>(null)
  const [selectedPages, setSelectedPages] = useState<Set<number>>(new Set())
  const [showThumbs, setShowThumbs] = useState(true)
  const [showAnnotPanel, setShowAnnotPanel] = useState(false)
  const [searchQuery, setSearchQuery] = useState('')
  const [searchInfo, setSearchInfo] = useState('')
  const [jumpValue, setJumpValue] = useState('1')
  const [splitStart, setSplitStart] = useState('1')
  const [splitEnd, setSplitEnd] = useState('1')
  const [watermarkText, setWatermarkText] = useState('')

  const fileInputRef = useRef<HTMLInputElement>(null)
  const mergeInputRef = useRef<HTMLInputElement>(null)
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const overlayRef = useRef<HTMLCanvasElement>(null)
  const containerRef = useRef<HTMLDivElement>(null)
  const renderTaskRef = useRef<pdfjsLib.RenderTask | null>(null)
  const loadTokenRef = useRef(0)
  const dragRef = useRef<{ startX: number; startY: number; id: string } | null>(null)
  const pageTextsRef = useRef<Map<number, string>>(new Map())
  const openedInitialRef = useRef(false)

  const baseName = fileName.replace(/\.pdf$/i, '') || 'document'
  const pageIndex = currentPage - 1

  // ---------- file loading ----------

  const validatePdfFile = (file: File): string | null => {
    const isPdf = file.type === 'application/pdf' || /\.pdf$/i.test(file.name)
    if (!isPdf) return 'That file is not a PDF. Please choose a .pdf file.'
    if (file.size > MAX_FILE_BYTES) return 'That PDF is over 50 MB. Please choose a smaller file.'
    if (file.size === 0) return 'That file is empty.'
    return null
  }

  const openFile = useCallback(async (file: File) => {
    const problem = validatePdfFile(file)
    if (problem) {
      setError(problem)
      return
    }
    setError('')
    setBusy('Opening PDF…')
    try {
      const buf = new Uint8Array(await file.arrayBuffer())
      setPdfBytes(buf)
      setFileName(file.name)
      setAnnotations([])
      setUndoStack([])
      setRedoStack([])
      setSelectedPages(new Set())
      setCurrentPage(1)
      setJumpValue('1')
      setSplitStart('1')
      setSplitEnd('1')
      setSearchInfo('')
      setSearchQuery('')
      pageTextsRef.current = new Map()
    } catch {
      setError('Could not read that file. Please try a different PDF.')
    } finally {
      setBusy('')
    }
  }, [])

  useEffect(() => {
    if (!openedInitialRef.current && initialFile) {
      openedInitialRef.current = true
      void openFile(initialFile)
    }
  }, [initialFile, openFile])

  // Load pdf.js document whenever the working bytes change.
  useEffect(() => {
    if (!pdfBytes) {
      setPdfDoc(null)
      setNumPages(0)
      setPageRotations([])
      return
    }
    const token = ++loadTokenRef.current
    let cancelled = false
    let doc: pdfjsLib.PDFDocumentProxy | null = null
    setBusy('Loading pages…')
    pdfjsLib
      .getDocument({ data: pdfBytes.slice() })
      .promise.then(async (d) => {
        doc = d
        const rotations: number[] = []
        for (let i = 1; i <= d.numPages; i++) {
          const p = await d.getPage(i)
          rotations.push(p.rotate)
        }
        if (cancelled || token !== loadTokenRef.current) return
        setPdfDoc(d)
        setNumPages(d.numPages)
        setPageRotations(rotations)
        setCurrentPage((prev) => Math.min(Math.max(prev, 1), d.numPages))
        setSplitEnd(String(d.numPages))
        pageTextsRef.current = new Map()
        setBusy('')
      })
      .catch(() => {
        if (cancelled || token !== loadTokenRef.current) return
        setError('That file could not be opened as a PDF. It may be damaged or protected.')
        setPdfDoc(null)
        setBusy('')
      })
    return () => {
      cancelled = true
      if (doc) void doc.destroy()
    }
  }, [pdfBytes])

  // ---------- main page render ----------

  useEffect(() => {
    if (!pdfDoc || !canvasRef.current) return
    const canvas = canvasRef.current
    const ctx = canvas.getContext('2d')
    if (!ctx) return
    let cancelled = false
    pdfDoc
      .getPage(currentPage)
      .then((page) => {
        if (cancelled) return
        const viewport = page.getViewport({ scale })
        setPageSize({ w: viewport.width / scale, h: viewport.height / scale })
        canvas.width = Math.max(1, Math.floor(viewport.width))
        canvas.height = Math.max(1, Math.floor(viewport.height))
        canvas.style.width = `${viewport.width}px`
        canvas.style.height = `${viewport.height}px`
        const overlay = overlayRef.current
        if (overlay) {
          overlay.width = canvas.width
          overlay.height = canvas.height
          overlay.style.width = canvas.style.width
          overlay.style.height = canvas.style.height
        }
        renderTaskRef.current?.cancel()
        const task = page.render({ canvas, canvasContext: ctx, viewport })
        renderTaskRef.current = task
        return task.promise.catch((e: unknown) => {
          if (e instanceof Error && e.name === 'RenderingCancelledException') return
          if (typeof e === 'object' && e !== null && 'name' in e && (e as { name?: string }).name === 'RenderingCancelledException') return
        })
      })
      .catch(() => {
        /* page no longer exists (e.g. after delete); load effect clamps the page */
      })
    return () => {
      cancelled = true
      renderTaskRef.current?.cancel()
    }
  }, [pdfDoc, currentPage, scale])

  // ---------- overlay annotation drawing ----------

  const drawAnnotationsOn = useCallback((ctx: CanvasRenderingContext2D, w: number, h: number) => {
    const list = annotations.filter((a) => a.page === pageIndex)
    const all = draft && draft.page === pageIndex ? [...list, draft] : list
    for (const a of all) {
      const x = a.x * w
      const y = a.y * h
      const aw = a.w * w
      const ah = a.h * h
      ctx.save()
      if (a.type === 'highlight') {
        ctx.globalAlpha = 0.4
        ctx.fillStyle = a.color
        ctx.fillRect(x, y, Math.max(aw, 2), Math.max(ah, 4))
      } else if (a.type === 'underline' || a.type === 'strike') {
        ctx.strokeStyle = a.color
        ctx.lineWidth = Math.max(2, h * 0.004)
        ctx.beginPath()
        const ly = a.type === 'underline' ? y + ah - 1 : y + ah / 2
        ctx.moveTo(x, ly)
        ctx.lineTo(x + aw, ly)
        ctx.stroke()
      } else if (a.type === 'pen') {
        const pts = a.points ?? []
        if (pts.length > 1) {
          ctx.strokeStyle = a.color
          ctx.lineWidth = Math.max(2, w * 0.004)
          ctx.lineJoin = 'round'
          ctx.lineCap = 'round'
          ctx.beginPath()
          ctx.moveTo(pts[0].x * w, pts[0].y * h)
          for (const p of pts.slice(1)) ctx.lineTo(p.x * w, p.y * h)
          ctx.stroke()
        }
      } else if (a.type === 'rect') {
        ctx.strokeStyle = a.color
        ctx.lineWidth = Math.max(2, w * 0.003)
        ctx.strokeRect(x, y, aw, ah)
      } else if (a.type === 'ellipse') {
        ctx.strokeStyle = a.color
        ctx.lineWidth = Math.max(2, w * 0.003)
        ctx.beginPath()
        ctx.ellipse(x + aw / 2, y + ah / 2, Math.abs(aw) / 2, Math.abs(ah) / 2, 0, 0, Math.PI * 2)
        ctx.stroke()
      } else if (a.type === 'textbox') {
        ctx.fillStyle = 'rgba(255,255,255,0.85)'
        ctx.fillRect(x, y, aw, ah)
        ctx.strokeStyle = a.color
        ctx.lineWidth = 1.5
        ctx.strokeRect(x, y, aw, ah)
        if (a.text) {
          ctx.fillStyle = '#111827'
          const fontSize = Math.max(10, ah / Math.max(1, a.text.split('\n').length) * 0.7)
          ctx.font = `${fontSize}px sans-serif`
          const lines = wrapText(a.text, { widthOfTextAtSize: (t: string, s: number) => ctx.measureText(t).width * (s / fontSize) }, fontSize, aw - 6)
          lines.forEach((line, i) => ctx.fillText(line, x + 3, y + fontSize + i * fontSize * 1.2, aw - 6))
        }
      } else if (a.type === 'comment') {
        ctx.fillStyle = '#fef08a'
        ctx.fillRect(x, y, Math.max(aw, 18), Math.max(ah, 18))
        ctx.strokeStyle = '#a16207'
        ctx.lineWidth = 1
        ctx.strokeRect(x, y, Math.max(aw, 18), Math.max(ah, 18))
        ctx.fillStyle = '#713f12'
        ctx.font = 'bold 12px sans-serif'
        ctx.fillText('!', x + 6, y + 14)
        if (a.text) {
          ctx.font = '11px sans-serif'
          ctx.fillStyle = '#111827'
          const lines = wrapText(a.text, { widthOfTextAtSize: (t: string) => ctx.measureText(t).width }, 11, Math.max(aw - 6, 40))
          lines.slice(0, 3).forEach((line, i) => ctx.fillText(line, x + 3, y + 28 + i * 13, Math.max(aw - 6, 40)))
        }
      }
      ctx.restore()
    }
  }, [annotations, draft, pageIndex])

  useEffect(() => {
    const overlay = overlayRef.current
    if (!overlay) return
    const ctx = overlay.getContext('2d')
    if (!ctx) return
    ctx.clearRect(0, 0, overlay.width, overlay.height)
    drawAnnotationsOn(ctx, overlay.width, overlay.height)
  }, [drawAnnotationsOn, scale, currentPage])

  // ---------- annotation history ----------

  const pushAnnotations = (next: Annotation[]) => {
    setUndoStack((prev) => [...prev.slice(-99), annotations])
    setRedoStack([])
    setAnnotations(next)
  }

  const undo = () => {
    if (undoStack.length === 0) return
    const prev = undoStack[undoStack.length - 1]
    setUndoStack(undoStack.slice(0, -1))
    setRedoStack((r) => [...r, annotations])
    setAnnotations(prev)
  }

  const redo = () => {
    if (redoStack.length === 0) return
    const next = redoStack[redoStack.length - 1]
    setRedoStack(redoStack.slice(0, -1))
    setUndoStack((u) => [...u, annotations])
    setAnnotations(next)
  }

  const deleteAnnotation = (id: string) => pushAnnotations(annotations.filter((a) => a.id !== id))

  // ---------- overlay pointer interaction ----------

  const fractionsFromEvent = (e: React.PointerEvent<HTMLCanvasElement>) => {
    const rect = e.currentTarget.getBoundingClientRect()
    return {
      fx: Math.min(1, Math.max(0, (e.clientX - rect.left) / rect.width)),
      fy: Math.min(1, Math.max(0, (e.clientY - rect.top) / rect.height)),
    }
  }

  const onOverlayPointerDown = (e: React.PointerEvent<HTMLCanvasElement>) => {
    if (tool === 'select' || !pdfDoc) return
    e.preventDefault()
    e.currentTarget.setPointerCapture(e.pointerId)
    const { fx, fy } = fractionsFromEvent(e)
    if (tool === 'textbox' || tool === 'comment') {
      const label = tool === 'textbox' ? 'Text for the text box:' : 'Comment text:'
      const text = window.prompt(label, '')
      if (text === null || text.trim() === '') return
      const ann: Annotation = {
        id: uid(),
        page: pageIndex,
        type: tool,
        x: fx,
        y: fy,
        w: tool === 'textbox' ? 0.3 : 0.22,
        h: tool === 'textbox' ? 0.09 : 0.12,
        text: text.trim(),
        color,
      }
      pushAnnotations([...annotations, ann])
      setShowAnnotPanel(true)
      return
    }
    const id = uid()
    dragRef.current = { startX: fx, startY: fy, id }
    if (tool === 'pen') {
      setDraft({ id, page: pageIndex, type: 'pen', x: fx, y: fy, w: 0, h: 0, points: [{ x: fx, y: fy }], color })
    } else {
      setDraft({ id, page: pageIndex, type: tool, x: fx, y: fy, w: 0, h: 0, color })
    }
  }

  const onOverlayPointerMove = (e: React.PointerEvent<HTMLCanvasElement>) => {
    const drag = dragRef.current
    if (!drag || !draft) return
    const { fx, fy } = fractionsFromEvent(e)
    if (draft.type === 'pen') {
      setDraft({ ...draft, points: [...(draft.points ?? []), { x: fx, y: fy }] })
    } else {
      setDraft({
        ...draft,
        x: Math.min(drag.startX, fx),
        y: Math.min(drag.startY, fy),
        w: Math.abs(fx - drag.startX),
        h: Math.abs(fy - drag.startY),
      })
    }
  }

  const onOverlayPointerUp = () => {
    const drag = dragRef.current
    dragRef.current = null
    if (!draft || !drag) return
    const finished = draft
    setDraft(null)
    if (finished.type === 'pen') {
      if ((finished.points?.length ?? 0) < 2) return
      const xs = finished.points!.map((p) => p.x)
      const ys = finished.points!.map((p) => p.y)
      const ann: Annotation = {
        ...finished,
        x: Math.min(...xs),
        y: Math.min(...ys),
        w: Math.max(...xs) - Math.min(...xs),
        h: Math.max(...ys) - Math.min(...ys),
      }
      pushAnnotations([...annotations, ann])
    } else {
      if (finished.w < 0.005 || finished.h < 0.003) return
      pushAnnotations([...annotations, finished])
    }
    setShowAnnotPanel(true)
  }

  // ---------- pdf-lib helpers ----------

  /** Re-serialize the working bytes through pdf-lib, optionally remapping annotation page indexes. */
  const applyPdfLib = async (
    label: string,
    mutator: (doc: PDFDocument) => Promise<void> | void,
    remapAnnotations?: (anns: Annotation[]) => Annotation[],
  ) => {
    if (!pdfBytes) return
    setBusy(`${label}…`)
    setError('')
    try {
      const doc = await PDFDocument.load(pdfBytes.slice())
      await mutator(doc)
      const next = await doc.save()
      if (remapAnnotations) setAnnotations((prev) => remapAnnotations(prev))
      setPdfBytes(new Uint8Array(next))
    } catch {
      setError(`Could not ${label.toLowerCase()}. The PDF may be protected.`)
    } finally {
      setBusy('')
    }
  }

  const remapAfterDelete = (deleted: number[]) => (anns: Annotation[]) => {
    const del = new Set(deleted)
    return anns
      .filter((a) => !del.has(a.page))
      .map((a) => ({ ...a, page: a.page - deleted.filter((d) => d < a.page).length }))
  }

  // ---------- downloads ----------

  const downloadCurrent = () => {
    if (!pdfBytes) return
    downloadBlob(new Blob([pdfBytes.slice().buffer as ArrayBuffer], { type: 'application/pdf' }), fileName || 'document.pdf')
  }

  const bakeAnnotations = async (doc: PDFDocument) => {
    const font = await doc.embedFont(StandardFonts.Helvetica)
    for (const a of annotations) {
      if (a.page >= doc.getPageCount()) continue
      const page: PDFPage = doc.getPage(a.page)
      const { width: W, height: H } = page.getSize()
      const rotation = page.getRotation().angle
      const c = hexToRgb(a.color)
      if (a.type === 'highlight') {
        const r = fracRectToPdf(a.x, a.y, a.w, a.h, W, H, rotation)
        page.drawRectangle({ x: r.x, y: r.y, width: r.w, height: Math.max(r.h, 4), color: c, opacity: 0.4 })
      } else if (a.type === 'underline' || a.type === 'strike') {
        const r = fracRectToPdf(a.x, a.y, a.w, a.h, W, H, rotation)
        const ly = a.type === 'underline' ? r.y : r.y + r.h / 2
        page.drawLine({ start: { x: r.x, y: ly }, end: { x: r.x + r.w, y: ly }, thickness: 1.5, color: c, opacity: 1 })
      } else if (a.type === 'pen') {
        const pts = (a.points ?? []).map((p) => fracPointToPdf(p.x, p.y, W, H, rotation))
        for (let i = 1; i < pts.length; i++) {
          page.drawLine({ start: pts[i - 1], end: pts[i], thickness: 1.5, color: c, opacity: 1 })
        }
      } else if (a.type === 'rect') {
        const r = fracRectToPdf(a.x, a.y, a.w, a.h, W, H, rotation)
        page.drawRectangle({ x: r.x, y: r.y, width: r.w, height: r.h, borderColor: c, borderWidth: 1.5, opacity: 0 })
      } else if (a.type === 'ellipse') {
        const r = fracRectToPdf(a.x, a.y, a.w, a.h, W, H, rotation)
        page.drawEllipse({ x: r.x + r.w / 2, y: r.y + r.h / 2, xScale: Math.abs(r.w) / 2, yScale: Math.abs(r.h) / 2, borderColor: c, borderWidth: 1.5, opacity: 0 })
      } else if (a.type === 'textbox') {
        const r = fracRectToPdf(a.x, a.y, a.w, a.h, W, H, rotation)
        page.drawRectangle({ x: r.x, y: r.y, width: r.w, height: r.h, color: rgb(1, 1, 1), opacity: 0.85, borderColor: c, borderWidth: 1 })
        if (a.text) {
          const size = 11
          const lines = wrapText(a.text, font, size, Math.max(r.w - 6, 10))
          lines.forEach((line, i) => {
            page.drawText(line, { x: r.x + 3, y: r.y + r.h - 14 - i * size * 1.2, size, font, color: rgb(0.07, 0.1, 0.15) })
          })
        }
      } else if (a.type === 'comment') {
        // pdf-lib has no public TextAnnotation helper, so a comment is baked as a
        // visible yellow note box with its text (see report).
        const r = fracRectToPdf(a.x, a.y, a.w, a.h, W, H, rotation)
        const boxW = Math.max(r.w, 90)
        const boxH = Math.max(r.h, 28)
        page.drawRectangle({ x: r.x, y: r.y + r.h - boxH, width: boxW, height: boxH, color: rgb(1, 0.94, 0.54), opacity: 0.95, borderColor: rgb(0.63, 0.38, 0.03), borderWidth: 1 })
        page.drawText('Note:', { x: r.x + 4, y: r.y + r.h - 14, size: 9, font, color: rgb(0.44, 0.25, 0.07) })
        if (a.text) {
          const lines = wrapText(a.text, font, 9, Math.max(boxW - 8, 10))
          lines.slice(0, 3).forEach((line, i) => {
            page.drawText(line, { x: r.x + 4, y: r.y + r.h - 26 - i * 11, size: 9, font, color: rgb(0.07, 0.1, 0.15) })
          })
        }
      }
    }
  }

  const downloadModified = async () => {
    if (!pdfBytes) return
    setBusy('Preparing download…')
    setError('')
    try {
      const doc = await PDFDocument.load(pdfBytes.slice())
      await bakeAnnotations(doc)
      const bytes = await doc.save()
      downloadBlob(new Blob([bytes.buffer as ArrayBuffer], { type: 'application/pdf' }), `${baseName}-modified.pdf`)
    } catch {
      setError('Could not build the modified PDF. The PDF may be protected.')
    } finally {
      setBusy('')
    }
  }

  // ---------- search ----------

  const extractAllText = async (): Promise<Map<number, string>> => {
    if (!pdfDoc) return new Map()
    if (pageTextsRef.current.size === numPages && numPages > 0) return pageTextsRef.current
    const map = new Map<number, string>()
    for (let i = 1; i <= numPages; i++) {
      const page = await pdfDoc.getPage(i)
      const tc = await page.getTextContent()
      const text = tc.items.map((it) => ('str' in it ? (it as { str: string }).str : '')).join(' ')
      map.set(i, text)
    }
    pageTextsRef.current = map
    return map
  }

  const runSearch = async () => {
    const q = searchQuery.trim().toLowerCase()
    if (!q || !pdfDoc) {
      setSearchInfo('')
      return
    }
    setBusy('Searching…')
    try {
      const texts = await extractAllText()
      let total = 0
      let pages = 0
      let first = -1
      for (let i = 1; i <= numPages; i++) {
        const text = (texts.get(i) ?? '').toLowerCase()
        let count = 0
        let idx = text.indexOf(q)
        while (idx !== -1) {
          count += 1
          idx = text.indexOf(q, idx + q.length)
        }
        if (count > 0) {
          pages += 1
          total += count
          if (first === -1) first = i
        }
      }
      if (first === -1) {
        setSearchInfo(`No matches for "${searchQuery.trim()}"`)
      } else {
        setSearchInfo(`${total} match${total === 1 ? '' : 'es'} on ${pages} page${pages === 1 ? '' : 's'}`)
        goToPage(first)
      }
    } finally {
      setBusy('')
    }
  }

  // ---------- navigation / zoom ----------

  const goToPage = (n: number) => {
    if (numPages === 0) return
    const clamped = Math.min(Math.max(1, Math.floor(n) || 1), numPages)
    setCurrentPage(clamped)
    setJumpValue(String(clamped))
  }

  const fitToWidth = () => {
    if (!pageSize || !containerRef.current) return
    const avail = containerRef.current.clientWidth - 48
    if (avail > 0 && pageSize.w > 0) setScale(avail / pageSize.w)
  }

  const fitToPage = () => {
    if (!pageSize || !containerRef.current) return
    const availW = containerRef.current.clientWidth - 48
    const availH = containerRef.current.clientHeight - 48
    if (availW > 0 && availH > 0 && pageSize.w > 0 && pageSize.h > 0) {
      setScale(Math.min(availW / pageSize.w, availH / pageSize.h))
    }
  }

  // ---------- page management ----------

  const rotateCurrentPage = () => {
    if (!pdfBytes || numPages === 0) return
    const idx = pageIndex
    void applyPdfLib(
      'Rotating page',
      async (doc) => {
        const page = doc.getPage(idx)
        const current = page.getRotation().angle
        page.setRotation(degrees((current + 90) % 360))
      },
      (anns) => anns.map((a) => (a.page === idx ? rotateAnnotation90CW(a) : a)),
    )
  }

  const deletePages = () => {
    if (!pdfBytes || numPages === 0) return
    const targets = selectedPages.size > 0 ? [...selectedPages].sort((a, b) => a - b) : [pageIndex]
    if (targets.length >= numPages) {
      setError('Cannot delete every page. At least one page must remain.')
      return
    }
    const label = targets.length === 1 ? `Delete page ${targets[0] + 1}?` : `Delete ${targets.length} selected pages?`
    if (!window.confirm(label)) return
    void applyPdfLib(
      'Deleting pages',
      async (doc) => {
        for (const idx of [...targets].sort((a, b) => b - a)) doc.removePage(idx)
      },
      remapAfterDelete(targets),
    ).then(() => {
      setSelectedPages(new Set())
      setCurrentPage((prev) => Math.min(prev, numPages - targets.length))
    })
  }

  const moveCurrentPage = (dir: -1 | 1) => {
    if (!pdfBytes || numPages < 2) return
    const from = pageIndex
    const to = from + dir
    if (to < 0 || to >= numPages) return
    const order = Array.from({ length: numPages }, (_, i) => i)
    ;[order[from], order[to]] = [order[to], order[from]]
    setBusy('Reordering pages…')
    setError('')
    PDFDocument.load(pdfBytes.slice())
      .then(async (src) => {
        const fresh = await PDFDocument.create()
        const copies = await fresh.copyPages(src, order)
        for (const p of copies) fresh.addPage(p)
        const next = await fresh.save()
        setAnnotations((prev) =>
          prev.map((a) => (a.page === from ? { ...a, page: to } : a.page === to ? { ...a, page: from } : a)),
        )
        setPdfBytes(new Uint8Array(next))
        setCurrentPage(to + 1)
      })
      .catch(() => setError('Could not reorder the pages.'))
      .finally(() => setBusy(''))
  }

  const mergeFiles = async (files: FileList | null) => {
    if (!files || files.length === 0 || !pdfBytes) return
    setBusy('Merging PDFs…')
    setError('')
    try {
      const doc = await PDFDocument.load(pdfBytes.slice())
      for (const file of Array.from(files)) {
        const problem = validatePdfFile(file)
        if (problem) {
          setError(`${file.name}: ${problem}`)
          continue
        }
        const other = await PDFDocument.load(new Uint8Array(await file.arrayBuffer()))
        const pages = await doc.copyPages(other, other.getPageIndices())
        for (const p of pages) doc.addPage(p)
      }
      const next = await doc.save()
      setPdfBytes(new Uint8Array(next))
    } catch {
      setError('Could not merge. One of the files may be damaged or protected.')
    } finally {
      setBusy('')
    }
  }

  const splitRange = async () => {
    if (!pdfBytes || numPages === 0) return
    const start = Math.min(Math.max(1, parseInt(splitStart, 10) || 1), numPages)
    const end = Math.min(Math.max(start, parseInt(splitEnd, 10) || start), numPages)
    setBusy('Splitting PDF…')
    try {
      const src = await PDFDocument.load(pdfBytes.slice())
      const out = await PDFDocument.create()
      const indices = Array.from({ length: end - start + 1 }, (_, i) => start - 1 + i)
      const pages = await out.copyPages(src, indices)
      for (const p of pages) out.addPage(p)
      const bytes = await out.save()
      downloadBlob(new Blob([bytes.buffer as ArrayBuffer], { type: 'application/pdf' }), `${baseName}-pages-${start}-${end}.pdf`)
    } catch {
      setError('Could not split the PDF.')
    } finally {
      setBusy('')
    }
  }

  const extractSelected = async () => {
    if (!pdfBytes || numPages === 0) return
    if (selectedPages.size === 0) {
      setError('Tick the checkboxes on thumbnails to choose pages to extract.')
      return
    }
    const indices = [...selectedPages].sort((a, b) => a - b)
    setBusy('Extracting pages…')
    try {
      const src = await PDFDocument.load(pdfBytes.slice())
      const out = await PDFDocument.create()
      const pages = await out.copyPages(src, indices)
      for (const p of pages) out.addPage(p)
      const bytes = await out.save()
      downloadBlob(new Blob([bytes.buffer as ArrayBuffer], { type: 'application/pdf' }), `${baseName}-extracted.pdf`)
    } catch {
      setError('Could not extract the selected pages.')
    } finally {
      setBusy('')
    }
  }

  const addPageNumbers = () => {
    if (!pdfBytes) return
    void applyPdfLib('Adding page numbers', async (doc) => {
      const font = await doc.embedFont(StandardFonts.Helvetica)
      doc.getPages().forEach((page, i) => {
        const { width } = page.getSize()
        const text = `Page ${i + 1}`
        const size = 10
        const textWidth = font.widthOfTextAtSize(text, size)
        page.drawText(text, { x: (width - textWidth) / 2, y: 20, size, font, color: rgb(0.3, 0.3, 0.3) })
      })
    })
  }

  const addWatermark = () => {
    const text = watermarkText.trim()
    if (!text || !pdfBytes) {
      if (!text) setError('Type the watermark text first.')
      return
    }
    void applyPdfLib('Adding watermark', async (doc) => {
      const font = await doc.embedFont(StandardFonts.Helvetica)
      doc.getPages().forEach((page) => {
        const { width, height } = page.getSize()
        const size = 48
        const textWidth = font.widthOfTextAtSize(text, size)
        page.drawText(text, {
          x: width / 2 - textWidth / 2,
          y: height / 2,
          size,
          font,
          color: rgb(0.5, 0.5, 0.5),
          opacity: 0.15,
          rotate: degrees(45),
        })
      })
    })
  }

  const clearPageAnnotations = () => {
    if (!window.confirm(`Remove all annotations on page ${currentPage}?`)) return
    pushAnnotations(annotations.filter((a) => a.page !== pageIndex))
  }

  const toggleSelectPage = (idx: number) => {
    setSelectedPages((prev) => {
      const next = new Set(prev)
      if (next.has(idx)) next.delete(idx)
      else next.add(idx)
      return next
    })
  }

  const onDrop = (e: React.DragEvent) => {
    e.preventDefault()
    const file = e.dataTransfer.files?.[0]
    if (file) void openFile(file)
  }

  const pageAnnotations = annotations.filter((a) => a.page === pageIndex)
  const annotating = tool !== 'select'

  // ---------- render ----------

  const toolBtn = (t: Tool, icon: React.ReactNode, label: string) => (
    <button
      key={t}
      type="button"
      onClick={() => {
        setTool(t)
        if (t !== 'select') setShowAnnotPanel(true)
      }}
      className={`ribbon-btn ${tool === t ? 'is-active' : ''}`}
      title={label}
      aria-pressed={tool === t}
    >
      {icon}
      <span>{label}</span>
    </button>
  )

  return (
    <div className="flex h-full flex-col bg-[#e8ebee]" onDragOver={(e) => e.preventDefault()} onDrop={onDrop}>
      <input
        ref={fileInputRef}
        type="file"
        accept="application/pdf,.pdf"
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0]
          if (f) void openFile(f)
          e.target.value = ''
        }}
      />
      <input
        ref={mergeInputRef}
        type="file"
        accept="application/pdf,.pdf"
        multiple
        className="hidden"
        onChange={(e) => {
          void mergeFiles(e.target.files)
          e.target.value = ''
        }}
      />

      {/* Title bar */}
      <header className="titlebar flex items-center gap-3 px-3 py-2 text-white">
        <FileText className="h-5 w-5 shrink-0" />
        <div className="min-w-0 flex-1">
          <div className="truncate text-sm font-semibold">My Word — PDF Tools</div>
          <div className="truncate text-xs text-blue-100">
            {pdfBytes ? `${fileName} • ${numPages} page${numPages === 1 ? '' : 's'}` : 'Open a PDF to view, annotate and organise it'}
          </div>
        </div>
        {busy && <span className="shrink-0 rounded bg-white/20 px-2 py-1 text-xs">{busy}</span>}
        <button
          type="button"
          onClick={onClose}
          className="flex shrink-0 items-center gap-1 rounded bg-white/15 px-3 py-1.5 text-sm font-semibold hover:bg-white/25"
          title="Back to dashboard"
        >
          <X className="h-4 w-4" /> Close
        </button>
      </header>

      {error && (
        <div className="flex items-center justify-between gap-3 bg-red-50 px-3 py-2 text-sm text-red-800">
          <span>{error}</span>
          <button type="button" onClick={() => setError('')} className="font-semibold underline">Dismiss</button>
        </div>
      )}

      {/* Toolbar */}
      <div className="flex items-stretch gap-1 overflow-x-auto border-b border-[#b6c4d6] bg-white px-2 py-1.5">
        <div className="flex shrink-0 items-stretch gap-0.5 border-r border-slate-200 pr-2">
          <button type="button" onClick={() => fileInputRef.current?.click()} className="ribbon-btn" title="Open a PDF from your device">
            <FolderOpen className="h-5 w-5" /><span>Open PDF</span>
          </button>
          <button type="button" onClick={downloadCurrent} disabled={!pdfBytes} className="ribbon-btn" title="Download the current PDF as-is">
            <Download className="h-5 w-5" /><span>Download</span>
          </button>
          <button type="button" onClick={() => void downloadModified()} disabled={!pdfBytes} className="ribbon-btn" title="Download with annotations baked in">
            <FileOutput className="h-5 w-5" /><span>Download modified</span>
          </button>
        </div>

        <div className="flex shrink-0 items-center gap-0.5 border-r border-slate-200 pr-2">
          <button type="button" onClick={() => goToPage(currentPage - 1)} disabled={!pdfBytes || currentPage <= 1} className="ribbon-btn" title="Previous page">
            <ChevronLeft className="h-5 w-5" /><span>Prev</span>
          </button>
          <div className="flex items-center gap-1 px-1 text-xs whitespace-nowrap">
            <span>Page</span>
            <input
              type="number"
              min={1}
              max={numPages || 1}
              value={jumpValue}
              disabled={!pdfBytes}
              onChange={(e) => setJumpValue(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') goToPage(parseInt(jumpValue, 10))
              }}
              onBlur={() => goToPage(parseInt(jumpValue, 10))}
              className="w-14 rounded border border-slate-300 px-1 py-1 text-center"
              aria-label="Go to page"
            />
            <span>of {numPages || '—'}</span>
          </div>
          <button type="button" onClick={() => goToPage(currentPage + 1)} disabled={!pdfBytes || currentPage >= numPages} className="ribbon-btn" title="Next page">
            <ChevronRight className="h-5 w-5" /><span>Next</span>
          </button>
          <button type="button" onClick={() => setScale((s) => Math.max(0.25, +(s - 0.1).toFixed(2)))} disabled={!pdfBytes} className="ribbon-btn" title="Zoom out">
            <ZoomOut className="h-5 w-5" /><span>Zoom out</span>
          </button>
          <span className="w-11 text-center text-xs font-semibold">{Math.round(scale * 100)}%</span>
          <button type="button" onClick={() => setScale((s) => Math.min(4, +(s + 0.1).toFixed(2)))} disabled={!pdfBytes} className="ribbon-btn" title="Zoom in">
            <ZoomIn className="h-5 w-5" /><span>Zoom in</span>
          </button>
          <button type="button" onClick={fitToWidth} disabled={!pdfBytes} className="ribbon-btn" title="Fit page width to the window">
            <MoveHorizontal className="h-5 w-5" /><span>Fit width</span>
          </button>
          <button type="button" onClick={fitToPage} disabled={!pdfBytes} className="ribbon-btn" title="Fit whole page in the window">
            <Maximize className="h-5 w-5" /><span>Fit page</span>
          </button>
          <button
            type="button"
            onClick={() => setShowThumbs((v) => !v)}
            className={`ribbon-btn ${showThumbs ? 'is-active' : ''}`}
            title="Show or hide page thumbnails"
          >
            <PanelLeft className="h-5 w-5" /><span>Thumbnails</span>
          </button>
        </div>

        <div className="flex shrink-0 items-center gap-1 border-r border-slate-200 pr-2">
          <input
            type="search"
            value={searchQuery}
            disabled={!pdfBytes}
            onChange={(e) => setSearchQuery(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') void runSearch()
            }}
            placeholder="Search text in PDF…"
            className="w-36 rounded border border-slate-300 px-2 py-1.5 text-xs"
            aria-label="Search text in PDF"
          />
          <button type="button" onClick={() => void runSearch()} disabled={!pdfBytes || !searchQuery.trim()} className="ribbon-btn" title="Search (case-insensitive)">
            <Search className="h-5 w-5" /><span>Search</span>
          </button>
          {searchInfo && <span className="max-w-40 truncate text-xs text-slate-600" title={searchInfo}>{searchInfo}</span>}
        </div>

        <div className="flex shrink-0 items-stretch gap-0.5 border-r border-slate-200 pr-2">
          {toolBtn('highlight', <Highlighter className="h-5 w-5" />, 'Highlight')}
          {toolBtn('underline', <Underline className="h-5 w-5" />, 'Underline')}
          {toolBtn('strike', <Strikethrough className="h-5 w-5" />, 'Strike')}
          {toolBtn('textbox', <Type className="h-5 w-5" />, 'Text box')}
          {toolBtn('comment', <MessageSquare className="h-5 w-5" />, 'Comment')}
          {toolBtn('pen', <Pen className="h-5 w-5" />, 'Pen')}
          {toolBtn('rect', <Square className="h-5 w-5" />, 'Rectangle')}
          {toolBtn('ellipse', <Circle className="h-5 w-5" />, 'Ellipse')}
          <div className="flex flex-col items-center justify-center gap-1 px-1">
            <div className="flex gap-1">
              {PALETTE.map((c) => (
                <button
                  key={c}
                  type="button"
                  onClick={() => setColor(c)}
                  className={`h-4 w-4 rounded-full border ${color === c ? 'ring-2 ring-blue-600 ring-offset-1' : 'border-slate-300'}`}
                  style={{ backgroundColor: c }}
                  title={`Colour ${c}`}
                  aria-label={`Use colour ${c}`}
                />
              ))}
              <input
                type="color"
                value={color}
                onChange={(e) => setColor(e.target.value)}
                className="h-4 w-8 cursor-pointer rounded border border-slate-300 p-0"
                title="Custom colour"
                aria-label="Custom colour"
              />
            </div>
            <span className="text-[11px] leading-none text-slate-600">Colour</span>
          </div>
          <button type="button" onClick={undo} disabled={undoStack.length === 0} className="ribbon-btn" title="Undo last annotation">
            <Undo2 className="h-5 w-5" /><span>Undo</span>
          </button>
          <button type="button" onClick={redo} disabled={redoStack.length === 0} className="ribbon-btn" title="Redo annotation">
            <Redo2 className="h-5 w-5" /><span>Redo</span>
          </button>
          <button type="button" onClick={clearPageAnnotations} disabled={!pdfBytes || pageAnnotations.length === 0} className="ribbon-btn" title="Remove all annotations on this page">
            <Trash2 className="h-5 w-5" /><span>Clear page</span>
          </button>
          <button
            type="button"
            onClick={() => setShowAnnotPanel((v) => !v)}
            className={`ribbon-btn ${showAnnotPanel ? 'is-active' : ''}`}
            title="Show annotation list for this page"
          >
            <List className="h-5 w-5" /><span>Annotations</span>
          </button>
        </div>

        <div className="flex shrink-0 items-stretch gap-0.5">
          <button type="button" onClick={rotateCurrentPage} disabled={!pdfBytes} className="ribbon-btn" title="Rotate this page 90° clockwise">
            <RotateCw className="h-5 w-5" /><span>Rotate</span>
          </button>
          <button type="button" onClick={() => moveCurrentPage(-1)} disabled={!pdfBytes || currentPage <= 1} className="ribbon-btn" title="Move this page up">
            <ArrowUp className="h-5 w-5" /><span>Move up</span>
          </button>
          <button type="button" onClick={() => moveCurrentPage(1)} disabled={!pdfBytes || currentPage >= numPages} className="ribbon-btn" title="Move this page down">
            <ArrowDown className="h-5 w-5" /><span>Move down</span>
          </button>
          <button type="button" onClick={deletePages} disabled={!pdfBytes} className="ribbon-btn" title="Delete this page (or the ticked pages)">
            <Trash2 className="h-5 w-5" /><span>Delete page</span>
          </button>
          <button type="button" onClick={() => mergeInputRef.current?.click()} disabled={!pdfBytes} className="ribbon-btn" title="Append more PDFs to this document">
            <FilePlus2 className="h-5 w-5" /><span>Merge PDFs</span>
          </button>
          <div className="flex flex-col items-center justify-center gap-0.5 px-1">
            <div className="flex items-center gap-1 text-[11px]">
              <span>Split</span>
              <input
                type="number" min={1} max={numPages || 1} value={splitStart}
                onChange={(e) => setSplitStart(e.target.value)}
                className="w-12 rounded border border-slate-300 px-1 py-0.5 text-center" aria-label="Split start page"
              />
              <span>–</span>
              <input
                type="number" min={1} max={numPages || 1} value={splitEnd}
                onChange={(e) => setSplitEnd(e.target.value)}
                className="w-12 rounded border border-slate-300 px-1 py-0.5 text-center" aria-label="Split end page"
              />
            </div>
            <button type="button" onClick={() => void splitRange()} disabled={!pdfBytes} className="flex items-center gap-1 rounded bg-blue-100 px-2 py-0.5 text-[11px] font-semibold text-blue-800 hover:bg-blue-200" title="Download pages in this range as a separate PDF">
              <Scissors className="h-3.5 w-3.5" /> Split &amp; download
            </button>
          </div>
          <button type="button" onClick={() => void extractSelected()} disabled={!pdfBytes || selectedPages.size === 0} className="ribbon-btn" title="Download only the ticked pages as a new PDF">
            <FileOutput className="h-5 w-5" /><span>Extract ticked</span>
          </button>
          <button type="button" onClick={addPageNumbers} disabled={!pdfBytes} className="ribbon-btn" title="Add “Page N” at the bottom of every page">
            <Hash className="h-5 w-5" /><span>Page numbers</span>
          </button>
          <div className="flex flex-col items-center justify-center gap-0.5 px-1">
            <input
              type="text"
              value={watermarkText}
              onChange={(e) => setWatermarkText(e.target.value)}
              placeholder="Watermark text…"
              className="w-28 rounded border border-slate-300 px-2 py-1 text-xs"
              aria-label="Watermark text"
            />
            <button type="button" onClick={addWatermark} disabled={!pdfBytes || !watermarkText.trim()} className="flex items-center gap-1 rounded bg-blue-100 px-2 py-0.5 text-[11px] font-semibold text-blue-800 hover:bg-blue-200" title="Draw the watermark diagonally on every page">
              <Stamp className="h-3.5 w-3.5" /> Add watermark
            </button>
          </div>
        </div>
      </div>

      {/* Body */}
      {!pdfBytes ? (
        <div className="flex flex-1 items-center justify-center p-6">
          <div className="w-full max-w-md rounded-lg border border-slate-300 bg-white p-8 text-center shadow">
            <FileText className="mx-auto h-12 w-12 text-blue-700" />
            <h2 className="mt-3 text-lg font-semibold text-slate-800">No PDF open yet</h2>
            <p className="mt-1 text-sm text-slate-600">
              Open a PDF to view its pages, highlight and annotate, and rearrange, merge, split or watermark it.
              You can also drop a PDF file anywhere on this screen.
            </p>
            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              className="mt-5 inline-flex items-center gap-2 rounded bg-blue-700 px-5 py-2.5 text-sm font-semibold text-white shadow hover:bg-blue-600"
            >
              <FolderOpen className="h-4 w-4" /> Open PDF
            </button>
            <p className="mt-3 text-xs text-slate-500">PDF files only, up to 50 MB. Your file stays on your device until you download the result.</p>
          </div>
        </div>
      ) : (
        <div className="flex min-h-0 flex-1">
          {showThumbs && pdfDoc && (
            <aside className="w-40 shrink-0 overflow-y-auto border-r border-slate-300 bg-[#f4f6f8] p-2">
              <p className="mb-2 text-xs font-semibold text-slate-600">
                Pages ({numPages}){selectedPages.size > 0 ? ` • ${selectedPages.size} ticked` : ''}
              </p>
              <div className="flex flex-col gap-3">
                {Array.from({ length: numPages }, (_, i) => (
                  <Thumbnail
                    key={`${i}-${pageRotations[i] ?? 0}`}
                    pdfDoc={pdfDoc}
                    pageNumber={i + 1}
                    active={currentPage === i + 1}
                    selected={selectedPages.has(i)}
                    annotationCount={annotations.filter((a) => a.page === i).length}
                    onOpen={() => goToPage(i + 1)}
                    onToggleSelect={() => toggleSelectPage(i)}
                  />
                ))}
              </div>
            </aside>
          )}

          <div ref={containerRef} className="min-w-0 flex-1 overflow-auto bg-[#525659] p-4">
            <div className="mx-auto w-fit">
              <div className="relative shadow-2xl" style={{ touchAction: tool === 'select' ? undefined : 'none' }}>
                <canvas ref={canvasRef} className="block bg-white" />
                <canvas
                  ref={overlayRef}
                  className="absolute top-0 left-0 block"
                  style={{ cursor: tool === 'select' ? 'default' : 'crosshair', touchAction: 'none' }}
                  onPointerDown={onOverlayPointerDown}
                  onPointerMove={onOverlayPointerMove}
                  onPointerUp={onOverlayPointerUp}
                  onPointerLeave={onOverlayPointerUp}
                />
              </div>
              <p className="mt-2 text-center text-xs text-white/80">
                Page {currentPage} of {numPages}
                {annotating ? ' • Drag on the page to draw with the selected tool' : ' • Pick an annotation tool above to mark up this page'}
              </p>
            </div>
          </div>

          {showAnnotPanel && (
            <aside className="w-64 shrink-0 overflow-y-auto border-l border-slate-300 bg-white p-3">
              <div className="flex items-center justify-between">
                <h3 className="text-sm font-semibold text-slate-800">Annotations — page {currentPage}</h3>
                <button type="button" onClick={() => setShowAnnotPanel(false)} className="rounded p-1 hover:bg-slate-100" title="Hide annotation list">
                  <X className="h-4 w-4" />
                </button>
              </div>
              {pageAnnotations.length === 0 ? (
                <p className="mt-3 text-xs text-slate-500">No annotations on this page yet. Choose a tool above and drag on the page.</p>
              ) : (
                <ul className="mt-2 flex flex-col gap-2">
                  {pageAnnotations.map((a, i) => (
                    <li key={a.id} className="flex items-start gap-2 rounded border border-slate-200 p-2 text-xs">
                      <span className="mt-0.5 h-3 w-3 shrink-0 rounded-full border border-slate-300" style={{ backgroundColor: a.color }} />
                      <span className="min-w-0 flex-1">
                        <span className="font-semibold">{i + 1}. {TYPE_LABELS[a.type]}</span>
                        {a.text && <span className="block truncate text-slate-600" title={a.text}>“{a.text}”</span>}
                      </span>
                      <button type="button" onClick={() => deleteAnnotation(a.id)} className="rounded p-1 text-red-600 hover:bg-red-50" title="Delete this annotation">
                        <Trash2 className="h-3.5 w-3.5" />
                      </button>
                    </li>
                  ))}
                </ul>
              )}
              <p className="mt-3 text-[11px] leading-relaxed text-slate-500">
                Annotations are saved into the PDF when you use “Download modified”. Undo and Redo only affect annotations, not page changes.
              </p>
            </aside>
          )}
        </div>
      )}
    </div>
  )
}

function Thumbnail({
  pdfDoc,
  pageNumber,
  active,
  selected,
  annotationCount,
  onOpen,
  onToggleSelect,
}: {
  pdfDoc: pdfjsLib.PDFDocumentProxy
  pageNumber: number
  active: boolean
  selected: boolean
  annotationCount: number
  onOpen: () => void
  onToggleSelect: () => void
}) {
  const canvasRef = useRef<HTMLCanvasElement>(null)

  useEffect(() => {
    let cancelled = false
    let task: pdfjsLib.RenderTask | null = null
    pdfDoc
      .getPage(pageNumber)
      .then((page) => {
        if (cancelled || !canvasRef.current) return
        const viewport = page.getViewport({ scale: 1 })
        const thumbScale = 128 / viewport.width
        const tv = page.getViewport({ scale: thumbScale })
        const canvas = canvasRef.current
        canvas.width = Math.max(1, Math.floor(tv.width))
        canvas.height = Math.max(1, Math.floor(tv.height))
        const ctx = canvas.getContext('2d')
        if (!ctx) return
        task = page.render({ canvas, canvasContext: ctx, viewport: tv })
        return task.promise.catch(() => undefined)
      })
      .catch(() => undefined)
    return () => {
      cancelled = true
      task?.cancel()
    }
  }, [pdfDoc, pageNumber])

  return (
    <div
      className={`rounded border bg-white p-1.5 ${active ? 'border-blue-600 ring-2 ring-blue-300' : 'border-slate-300'}`}
    >
      <button type="button" onClick={onOpen} className="block w-full" title={`Go to page ${pageNumber}`}>
        <canvas ref={canvasRef} className="mx-auto block max-w-full bg-white" />
      </button>
      <div className="mt-1 flex items-center justify-between text-[11px]">
        <label className="flex cursor-pointer items-center gap-1 text-slate-700">
          <input type="checkbox" checked={selected} onChange={onToggleSelect} aria-label={`Select page ${pageNumber}`} />
          {pageNumber}
        </label>
        {annotationCount > 0 && (
          <span className="rounded bg-amber-100 px-1.5 py-0.5 font-semibold text-amber-800" title={`${annotationCount} annotation(s) on this page`}>
            {annotationCount} ✎
          </span>
        )}
      </div>
    </div>
  )
}
