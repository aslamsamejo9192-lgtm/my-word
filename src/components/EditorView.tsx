import { useCallback, useEffect, useRef, useState } from 'react'
import type { ReactNode } from 'react'
import { EditorContent, useEditor } from '@tiptap/react'
import StarterKit from '@tiptap/starter-kit'
import { Table } from '@tiptap/extension-table'
import TableRow from '@tiptap/extension-table-row'
import TableHeader from '@tiptap/extension-table-header'
import TableCell from '@tiptap/extension-table-cell'
import Image from '@tiptap/extension-image'
import Link from '@tiptap/extension-link'
import TextAlign from '@tiptap/extension-text-align'
import Highlight from '@tiptap/extension-highlight'
import Underline from '@tiptap/extension-underline'
import Subscript from '@tiptap/extension-subscript'
import Superscript from '@tiptap/extension-superscript'
import { FontSize, TextStyle } from '@tiptap/extension-text-style'
import Color from '@tiptap/extension-color'
import FontFamily from '@tiptap/extension-font-family'
import Placeholder from '@tiptap/extension-placeholder'
import CharacterCount from '@tiptap/extension-character-count'
import Paragraph from '@tiptap/extension-paragraph'
import Heading from '@tiptap/extension-heading'
import { Mark, Node, mergeAttributes } from '@tiptap/core'
import type { Content, Editor } from '@tiptap/core'
import { Plugin, PluginKey } from 'prosemirror-state'
import {
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  Bold,
  BookOpen,
  ClipboardPaste,
  Columns3,
  Copy,
  FilePlus,
  FileText,
  FolderOpen,
  Hash,
  Highlighter,
  Image as ImageIcon,
  Info,
  Italic,
  Link2,
  Link2Off,
  List,
  ListOrdered,
  Maximize,
  MessageSquare,
  Minus,
  Palette,
  Plus,
  Printer,
  Redo2,
  Save,
  Scissors,
  Search,
  Strikethrough,
  Subscript as SubscriptIcon,
  Superscript as SuperscriptIcon,
  Table as TableIcon,
  Type,
  Underline as UnderlineIcon,
  Undo2,
  X,
  ZoomIn,
  ZoomOut,
} from 'lucide-react'
import { exportDocx } from '../lib/docx'
import { marginPx, pageDims, uid } from '../lib/types'
import type { MyDocument, PageSettings } from '../lib/types'

export interface EditorViewProps {
  doc: MyDocument
  recentDocs: { id: string; title: string; updatedAt: number }[]
  onSave: (doc: MyDocument) => void
  onSaveAs: (doc: MyDocument, newTitle: string) => void
  onClose: () => void
  onNew: () => void
  onOpenRecent: (id: string) => void
  onImportFile: (file: File) => void
  onOpenPdf: () => void
}

/* ------------------------------------------------------------------ */
/* Command type merging for the custom extensions                      */
/* ------------------------------------------------------------------ */

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    // NOTE: the `paragraph` command group cannot be augmented here — the
    // Paragraph extension already declares it as exactly { setParagraph },
    // and TS requires merged interface properties to be identical (TS2717).
    // The custom paragraph commands (setLineHeight / setIndent /
    // toggleParaBorder / setParaSpacing) still exist at runtime via
    // CustomParagraph.addCommands; call sites below use the already-typed
    // updateAttributes('paragraph', …) instead.
    tracked: {
      setTrackChanges: (enabled: boolean) => ReturnType
    }
    pageBreak: {
      insertPageBreak: () => ReturnType
    }
  }
}

/* ------------------------------------------------------------------ */
/* Custom paragraph: line height / indent / border / spacing           */
/* ------------------------------------------------------------------ */

const CustomParagraph = Paragraph.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      lineHeight: {
        default: null,
        parseHTML: (element: HTMLElement) => element.style.lineHeight || null,
        renderHTML: (attributes: Record<string, unknown>) =>
          attributes.lineHeight ? { style: `line-height: ${attributes.lineHeight}` } : {},
      },
      indent: {
        default: 0,
        parseHTML: (element: HTMLElement) => {
          const px = parseFloat(element.style.marginLeft || '')
          return Number.isNaN(px) ? 0 : Math.round(px / 24)
        },
        renderHTML: (attributes: Record<string, unknown>) =>
          typeof attributes.indent === 'number' && attributes.indent > 0
            ? { style: `margin-left: ${attributes.indent * 24}px` }
            : {},
      },
      paraBorder: {
        default: false,
        parseHTML: (element: HTMLElement) =>
          element.style.borderStyle === 'solid' || element.style.borderWidth !== '',
        renderHTML: (attributes: Record<string, unknown>) =>
          attributes.paraBorder ? { style: 'border: 1px solid #9ca3af; padding: 4px' } : {},
      },
      paraSpacing: {
        default: 'normal',
        parseHTML: (element: HTMLElement) => {
          const mb = parseFloat(element.style.marginBottom || '')
          if (Number.isNaN(mb)) return 'normal'
          if (mb <= 4) return 'compact'
          if (mb >= 14) return 'relaxed'
          return 'normal'
        },
        renderHTML: (attributes: Record<string, unknown>) => {
          const mb =
            attributes.paraSpacing === 'compact' ? 2 : attributes.paraSpacing === 'relaxed' ? 16 : 8
          return { style: `margin-bottom: ${mb}px` }
        },
      },
    }
  },
  addCommands() {
    return {
      ...this.parent?.(),
      setLineHeight:
        (lineHeight: string) =>
        ({ commands }: { commands: { updateAttributes: (t: string, a: Record<string, unknown>) => boolean } }) =>
          commands.updateAttributes('paragraph', { lineHeight }),
      setIndent:
        (indent: number) =>
        ({ commands }: { commands: { updateAttributes: (t: string, a: Record<string, unknown>) => boolean } }) =>
          commands.updateAttributes('paragraph', { indent: Math.max(0, Math.min(5, indent)) }),
      toggleParaBorder:
        () =>
        ({
          commands,
          editor,
        }: {
          commands: { updateAttributes: (t: string, a: Record<string, unknown>) => boolean }
          editor: Editor
        }) => {
          const current = editor.getAttributes('paragraph').paraBorder as boolean | undefined
          return commands.updateAttributes('paragraph', { paraBorder: !current })
        },
      setParaSpacing:
        (spacing: 'compact' | 'normal' | 'relaxed') =>
        ({ commands }: { commands: { updateAttributes: (t: string, a: Record<string, unknown>) => boolean } }) =>
          commands.updateAttributes('paragraph', { paraSpacing: spacing }),
    }
  },
})

/* ------------------------------------------------------------------ */
/* Custom nodes: page break, title, subtitle, page number              */
/* ------------------------------------------------------------------ */

const PageBreakNode = Node.create({
  name: 'pageBreak',
  group: 'block',
  atom: true,
  parseHTML() {
    return [{ tag: 'div[data-page-break]' }]
  },
  renderHTML() {
    return ['div', { 'data-page-break': '', class: 'page-break-node' }]
  },
  addCommands() {
    return {
      insertPageBreak:
        () =>
        ({ commands }: { commands: { insertContent: (c: Record<string, unknown>) => boolean } }) =>
          commands.insertContent({ type: 'pageBreak' }),
    }
  },
})

const TitleNode = Node.create({
  name: 'title',
  content: 'inline*',
  group: 'block',
  defining: true,
  parseHTML() {
    return [{ tag: 'h1.title-node' }]
  },
  renderHTML({ HTMLAttributes }) {
    return [
      'h1',
      mergeAttributes(HTMLAttributes, {
        class: 'title-node',
        style: 'font-size:32px;text-align:center;color:#1f4e9c;margin:8px 0',
      }),
      0,
    ]
  },
})

const SubtitleNode = Node.create({
  name: 'subtitle',
  content: 'inline*',
  group: 'block',
  defining: true,
  parseHTML() {
    return [{ tag: 'h2.subtitle-node' }]
  },
  renderHTML({ HTMLAttributes }) {
    return [
      'h2',
      mergeAttributes(HTMLAttributes, {
        class: 'subtitle-node',
        style: 'font-size:18px;text-align:center;color:#6b7280;font-weight:400',
      }),
      0,
    ]
  },
})

const PageNumberNode = Node.create({
  name: 'pageNumber',
  inline: true,
  group: 'inline',
  atom: true,
  selectable: true,
  parseHTML() {
    return [{ tag: 'span.page-number' }]
  },
  renderHTML() {
    return ['span', { class: 'page-number' }]
  },
})

/* ------------------------------------------------------------------ */
/* Track Changes mark: typed text is inserted blue + underlined        */
/* ------------------------------------------------------------------ */

const TrackedMark = Mark.create({
  name: 'tracked',
  addStorage() {
    return { enabled: false }
  },
  parseHTML() {
    return [{ tag: 'span[data-tracked]' }]
  },
  renderHTML({ HTMLAttributes }) {
    return [
      'span',
      mergeAttributes(HTMLAttributes, {
        'data-tracked': '',
        style: 'color: #2563eb; text-decoration: underline;',
      }),
      0,
    ]
  },
  addCommands() {
    return {
      setTrackChanges:
        (enabled: boolean) =>
        () => {
          ;(this.storage as { enabled: boolean }).enabled = enabled
          return true
        },
    }
  },
  addProseMirrorPlugins() {
    const storage = this.storage as { enabled: boolean }
    return [
      new Plugin({
        key: new PluginKey('tracked-input'),
        props: {
          handleTextInput(view, from, to, text) {
            if (!storage.enabled) return false
            const markType = view.state.schema.marks.tracked
            if (!markType) return false
            const textNode = view.state.schema.text(text, [markType.create()])
            view.dispatch(view.state.tr.replaceRangeWith(from, to, textNode))
            return true
          },
        },
      }),
    ]
  },
})

/* ------------------------------------------------------------------ */
/* Small UI helpers                                                    */
/* ------------------------------------------------------------------ */

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  document.body.appendChild(a)
  a.click()
  a.remove()
  setTimeout(() => URL.revokeObjectURL(url), 1000)
}

function Dialog({
  title,
  onClose,
  children,
}: {
  title: string
  onClose: () => void
  children: ReactNode
}) {
  return (
    <div
      className="no-print fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4"
      onClick={onClose}
    >
      <div
        className="max-h-[85vh] w-full max-w-lg overflow-auto rounded bg-white p-4 shadow-xl"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="mb-3 flex items-center justify-between">
          <h3 className="text-sm font-semibold text-slate-800">{title}</h3>
          <button type="button" onClick={onClose} className="rounded p-1 hover:bg-slate-100" aria-label="Close">
            <X size={16} />
          </button>
        </div>
        {children}
      </div>
    </div>
  )
}

function RB({
  label,
  onClick,
  active,
  disabled,
  title,
  children,
}: {
  label: string
  onClick: () => void
  active?: boolean
  disabled?: boolean
  title?: string
  children: ReactNode
}) {
  return (
    <button
      type="button"
      className={`ribbon-btn${active ? ' is-active' : ''}`}
      onClick={onClick}
      disabled={disabled}
      title={title ?? label}
    >
      {children}
      <span>{label}</span>
    </button>
  )
}

function RGroup({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div className="flex shrink-0 flex-col items-center justify-between border-r border-slate-300 px-1.5 last:border-r-0">
      <div className="flex items-start gap-0.5">{children}</div>
      <span className="mt-0.5 text-[10px] text-slate-500">{label}</span>
    </div>
  )
}

const FONT_FAMILIES = ['Arial', 'Calibri', 'Times New Roman', 'Georgia', 'Courier New', 'Verdana', 'Tahoma']
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]
const SYMBOLS = ['©', '®', '™', '€', '£', '¥', '§', '¶', '•', '…', '–', '—', '°', '±', '×', '÷', '←', '→', '↑', '↓', '✓', '✔', '★', '☆', '♥', '♦']

const selectCls = 'h-7 rounded border border-slate-300 bg-white px-1 text-xs text-slate-800'
const inputCls = 'w-full rounded border border-slate-300 px-2 py-1.5 text-sm text-slate-800'
const btnCls = 'rounded bg-[#2f6cb5] px-3 py-1.5 text-xs font-medium text-white hover:bg-[#1f4e9c] disabled:opacity-50'
const btnGhostCls = 'rounded border border-slate-300 bg-white px-3 py-1.5 text-xs text-slate-700 hover:bg-slate-100'

type TabId = 'FILE' | 'HOME' | 'INSERT' | 'PAGE LAYOUT' | 'REFERENCES' | 'REVIEW' | 'VIEW'
const TABS: TabId[] = ['FILE', 'HOME', 'INSERT', 'PAGE LAYOUT', 'REFERENCES', 'REVIEW', 'VIEW']

interface HeadingEntry {
  pos: number
  level: number
  text: string
}

export default function EditorView(props: EditorViewProps) {
  const [docState, setDocState] = useState<MyDocument>(props.doc)
  const docStateRef = useRef(docState)
  const [activeTab, setActiveTab] = useState<TabId>('HOME')
  const [fileOpen, setFileOpen] = useState(false)
  const [dialog, setDialog] = useState<
    | null
    | 'docInfo'
    | 'toc'
    | 'wordCount'
    | 'find'
    | 'table'
    | 'symbols'
    | 'review'
    | 'citation'
    | 'nav'
  >(null)
  const [status, setStatus] = useState<'saved' | 'saving' | 'dirty'>('saved')
  const [counts, setCounts] = useState({ words: 0, chars: 0 })
  const [estPages, setEstPages] = useState(1)
  const [navPane, setNavPane] = useState(false)
  const [commentsPane, setCommentsPane] = useState(false)
  const [spellcheck, setSpellcheck] = useState(true)
  const [trackChanges, setTrackChanges] = useState(false)
  const [imageSelected, setImageSelected] = useState(false)
  const [, setSelTick] = useState(0)
  const [findText, setFindText] = useState('')
  const [replaceText, setReplaceText] = useState('')
  const [matches, setMatches] = useState<{ from: number; to: number }[]>([])
  const [matchIdx, setMatchIdx] = useState(-1)
  const [tableRows, setTableRows] = useState(3)
  const [tableCols, setTableCols] = useState(3)
  const [citAuthor, setCitAuthor] = useState('')
  const [citYear, setCitYear] = useState('')
  const [citTitle, setCitTitle] = useState('')

  const saveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  const workspaceRef = useRef<HTMLDivElement>(null)
  const contentWrapRef = useRef<HTMLDivElement>(null)
  const importInputRef = useRef<HTMLInputElement>(null)
  const pictureInputRef = useRef<HTMLInputElement>(null)
  const footnotePosRef = useRef<Record<string, number>>({})
  const docIdRef = useRef(props.doc.id)

  useEffect(() => {
    docStateRef.current = docState
  }, [docState])

  const updateDoc = useCallback((patch: Partial<MyDocument>) => {
    setDocState((prev) => {
      const next = { ...prev, ...patch }
      docStateRef.current = next
      return next
    })
  }, [])

  const updatePageSettings = useCallback(
    (patch: Partial<PageSettings>) => {
      setDocState((prev) => {
        const next = { ...prev, pageSettings: { ...prev.pageSettings, ...patch } }
        docStateRef.current = next
        return next
      })
    },
    [],
  )

  const editor = useEditor({
    immediatelyRender: false,
    extensions: [
      StarterKit.configure({ paragraph: false, heading: false }),
      CustomParagraph,
      Heading.configure({ levels: [1, 2, 3] }),
      TitleNode,
      SubtitleNode,
      PageBreakNode,
      PageNumberNode,
      TrackedMark,
      Table.configure({ resizable: true }),
      TableRow,
      TableHeader,
      TableCell,
      Image,
      Link.configure({ openOnClick: true }),
      TextAlign.configure({ types: ['heading', 'paragraph'] }),
      Highlight.configure({ multicolor: true }),
      Underline,
      Subscript,
      Superscript,
      TextStyle,
      FontSize,
      Color,
      FontFamily,
      Placeholder.configure({ placeholder: 'Start typing…' }),
      CharacterCount,
    ],
    content:
      props.doc.html && props.doc.html.trim()
        ? props.doc.html
        : (props.doc.content as unknown as Content),
    editorProps: {
      attributes: { spellcheck: 'true' },
    },
    onUpdate: ({ editor: ed }) => {
      setCounts({
        words: ed.storage.characterCount.words(),
        chars: ed.storage.characterCount.characters(),
      })
      setSelTick((t) => t + 1)
      setStatus('dirty')
      if (saveTimerRef.current) clearTimeout(saveTimerRef.current)
      saveTimerRef.current = setTimeout(() => void saveNowRef.current(), 2000)
      setTimeout(() => refreshPages(ed), 60)
    },
    onSelectionUpdate: ({ editor: ed }) => {
      setImageSelected(ed.isActive('image'))
      setSelTick((t) => t + 1)
    },
  })

  const editorRef = useRef<Editor | null>(null)
  useEffect(() => {
    editorRef.current = editor
  }, [editor])

  const buildDoc = useCallback((): MyDocument | null => {
    const ed = editorRef.current
    if (!ed) return null
    const base = docStateRef.current
    return {
      ...base,
      content: ed.getJSON() as Record<string, unknown>,
      html: ed.getHTML(),
      updatedAt: Date.now(),
    }
  }, [])

  const saveNow = useCallback(() => {
    const doc = buildDoc()
    if (!doc) return
    if (saveTimerRef.current) {
      clearTimeout(saveTimerRef.current)
      saveTimerRef.current = null
    }
    setStatus('saving')
    props.onSave(doc)
    setStatus('saved')
  }, [buildDoc, props])

  const saveNowRef = useRef(saveNow)
  useEffect(() => {
    saveNowRef.current = saveNow
  }, [saveNow])

  function refreshPages(ed: Editor) {
    const wrap = contentWrapRef.current
    if (!wrap) return
    const dims = pageDims(docStateRef.current.pageSettings.size, docStateRef.current.pageSettings.orientation)
    const h = wrap.scrollHeight || ed.view.dom.scrollHeight || dims.height
    setEstPages(Math.max(1, Math.ceil(h / dims.height)))
  }

  // If App hands us a different document, swap the editor content.
  useEffect(() => {
    if (props.doc.id === docIdRef.current || !editor) return
    docIdRef.current = props.doc.id
    setDocState(props.doc)
    docStateRef.current = props.doc
    editor.commands.setContent(
      props.doc.html && props.doc.html.trim()
        ? props.doc.html
        : (props.doc.content as unknown as Content),
    )
    setStatus('saved')
  }, [props.doc, editor])

  // Initial counts + page estimate once the editor exists.
  useEffect(() => {
    if (!editor) return
    setCounts({
      words: editor.storage.characterCount.words(),
      chars: editor.storage.characterCount.characters(),
    })
    refreshPages(editor)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editor])

  // Warn about unsaved changes.
  useEffect(() => {
    if (status !== 'dirty') return
    const handler = (e: BeforeUnloadEvent) => {
      e.preventDefault()
    }
    window.addEventListener('beforeunload', handler)
    return () => window.removeEventListener('beforeunload', handler)
  }, [status])

  // Browser spellcheck toggle (honest label: this is the browser's checker).
  useEffect(() => {
    if (editor) editor.view.dom.spellcheck = spellcheck
  }, [editor, spellcheck])

  // Keyboard shortcuts: Ctrl/Cmd+S save, Ctrl/Cmd+P print, Ctrl/Cmd+F find, Ctrl/Cmd+H replace.
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      const mod = e.ctrlKey || e.metaKey
      if (!mod) return
      const key = e.key.toLowerCase()
      if (key === 's') {
        e.preventDefault()
        saveNowRef.current()
      } else if (key === 'p') {
        e.preventDefault()
        window.print()
      } else if (key === 'f') {
        e.preventDefault()
        setDialog('find')
      } else if (key === 'h') {
        e.preventDefault()
        setDialog('find')
      }
    }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [])

  const dims = pageDims(docState.pageSettings.size, docState.pageSettings.orientation)
  const margin = marginPx(docState.pageSettings.margins)
  const zoom = docState.pageSettings.zoom
  const ps = docState.pageSettings

  const collectHeadings = useCallback((): HeadingEntry[] => {
    const ed = editorRef.current
    if (!ed) return []
    const out: HeadingEntry[] = []
    ed.state.doc.descendants((node, pos) => {
      if (node.type.name === 'heading' || node.type.name === 'title' || node.type.name === 'subtitle') {
        const level =
          node.type.name === 'heading'
            ? ((node.attrs.level as number) || 2)
            : node.type.name === 'title'
              ? 1
              : 2
        out.push({ pos, level, text: node.textContent || '(empty heading)' })
      }
    })
    return out
  }, [])

  const jumpToPos = useCallback((pos: number) => {
    const ed = editorRef.current
    if (!ed) return
    ed.chain().focus().setTextSelection(Math.min(pos, ed.state.doc.content.size)).run()
    ed.view.dom.scrollIntoView({ block: 'nearest' })
  }, [])

  /* ---------------- clipboard ---------------- */

  const selectionText = useCallback((): { from: number; to: number; text: string } | null => {
    const ed = editorRef.current
    if (!ed) return null
    const { from, to } = ed.state.selection
    if (from === to) return null
    return { from, to, text: ed.state.doc.textBetween(from, to, ' ') }
  }, [])

  const doCopy = useCallback(() => {
    const sel = selectionText()
    try {
      if (document.execCommand('copy')) return
    } catch {
      /* fall through */
    }
    if (sel) void navigator.clipboard.writeText(sel.text).catch(() => undefined)
  }, [selectionText])

  const doCut = useCallback(() => {
    const sel = selectionText()
    try {
      if (document.execCommand('cut')) return
    } catch {
      /* fall through */
    }
    if (sel) {
      void navigator.clipboard.writeText(sel.text).catch(() => undefined)
      editorRef.current?.chain().focus().deleteSelection().run()
    }
  }, [selectionText])

  const doPaste = useCallback(() => {
    const ed = editorRef.current
    if (!ed) return
    // execCommand('paste') is blocked by browsers; read the clipboard directly.
    void navigator.clipboard
      .readText()
      .then((text) => {
        if (text) ed.chain().focus().insertContent(text).run()
      })
      .catch(() => {
        try {
          document.execCommand('paste')
        } catch {
          alert('Paste is blocked by the browser — use Ctrl+V to paste.')
        }
      })
  }, [])

  /* ---------------- find & replace ---------------- */

  const findMatches = useCallback((query: string): { from: number; to: number }[] => {
    const ed = editorRef.current
    if (!ed || !query) return []
    const needle = query.toLowerCase()
    const out: { from: number; to: number }[] = []
    ed.state.doc.descendants((node, pos) => {
      if (!node.isText || !node.text) return
      const hay = node.text.toLowerCase()
      let i = hay.indexOf(needle)
      while (i !== -1) {
        out.push({ from: pos + i, to: pos + i + query.length })
        i = hay.indexOf(needle, i + 1)
      }
    })
    return out
  }, [])

  const runFind = useCallback(
    (dir: 1 | -1) => {
      const found = findMatches(findText)
      setMatches(found)
      if (found.length === 0) {
        setMatchIdx(-1)
        return
      }
      const next = dir === 1 ? (matchIdx + 1) % found.length : (matchIdx - 1 + found.length) % found.length
      setMatchIdx(next)
      const m = found[next]
      jumpToPos(m.from)
      editorRef.current?.chain().focus().setTextSelection({ from: m.from, to: m.to }).run()
    },
    [findMatches, findText, matchIdx, jumpToPos],
  )

  const replaceCurrent = useCallback(() => {
    const ed = editorRef.current
    if (!ed || matchIdx < 0 || !matches[matchIdx]) return
    const m = matches[matchIdx]
    ed.chain().focus().insertContentAt({ from: m.from, to: m.to }, replaceText).run()
    const found = findMatches(findText)
    setMatches(found)
    setMatchIdx(found.length ? 0 : -1)
  }, [matchIdx, matches, replaceText, findMatches, findText])

  const replaceAll = useCallback(() => {
    const ed = editorRef.current
    if (!ed) return
    const found = findMatches(findText)
    for (let i = found.length - 1; i >= 0; i--) {
      ed.chain().focus().insertContentAt({ from: found[i].from, to: found[i].to }, replaceText).run()
    }
    setMatches([])
    setMatchIdx(-1)
  }, [findMatches, findText, replaceText])

  /* ---------------- references helpers ---------------- */

  const insertFootnoteLike = useCallback(
    (kind: 'footnotes' | 'endnotes') => {
      const ed = editorRef.current
      if (!ed) return
      const text = window.prompt(kind === 'footnotes' ? 'Footnote text:' : 'Endnote text:')
      if (!text) return
      const list = docStateRef.current[kind]
      const n = list.length + 1
      const pos = ed.state.selection.from
      ed.chain()
        .focus()
        .insertContent({ type: 'text', text: `[${n}]`, marks: [{ type: 'superscript' }] })
        .run()
      const entry = { id: uid(), text }
      footnotePosRef.current[entry.id] = pos
      updateDoc({ [kind]: [...list, entry] } as Partial<MyDocument>)
      setStatus('dirty')
    },
    [updateDoc],
  )

  const addCitation = useCallback(() => {
    const ed = editorRef.current
    if (!citAuthor.trim() || !citYear.trim()) {
      alert('Author and year are required for a citation.')
      return
    }
    const entry = { id: uid(), author: citAuthor.trim(), year: citYear.trim(), title: citTitle.trim() }
    ed?.chain().focus().insertContent(`(${entry.author}, ${entry.year})`).run()
    updateDoc({ citations: [...docStateRef.current.citations, entry] })
    setCitAuthor('')
    setCitYear('')
    setCitTitle('')
    setStatus('dirty')
    setDialog(null)
  }, [citAuthor, citYear, citTitle, updateDoc])

  const insertBibliography = useCallback(() => {
    const ed = editorRef.current
    if (!ed) return
    const cites = docStateRef.current.citations
    if (cites.length === 0) {
      alert('No citations yet — add one first.')
      return
    }
    const blocks = [
      { type: 'heading', attrs: { level: 2 }, content: [{ type: 'text', text: 'Bibliography' }] },
      ...cites.map((c) => ({
        type: 'paragraph',
        content: [{ type: 'text', text: `${c.author} (${c.year}). ${c.title}.` }],
      })),
    ]
    ed.chain().focus().insertContent(blocks).run()
    setStatus('dirty')
  }, [])

  const insertTocBlock = useCallback(() => {
    const ed = editorRef.current
    if (!ed) return
    const heads = collectHeadings()
    const blocks = [
      { type: 'heading', attrs: { level: 1 }, content: [{ type: 'text', text: 'Table of Contents' }] },
      ...(heads.length
        ? heads.map((h) => ({
            type: 'paragraph',
            content: [{ type: 'text', text: `${'  '.repeat(Math.max(0, h.level - 1))}${h.text}` }],
          }))
        : [{ type: 'paragraph', content: [{ type: 'text', text: '(No headings in this document yet.)' }] }]),
    ]
    ed.chain().focus().insertContent(blocks).run()
    setStatus('dirty')
    setDialog(null)
  }, [collectHeadings])

  const insertCaption = useCallback(() => {
    const ed = editorRef.current
    if (!ed) return
    const text = window.prompt('Caption text:')
    if (!text) return
    const count = (ed.getText().match(/Figure /g) || []).length
    const n = count + 1
    ed.chain()
      .focus()
      .insertContent({
        type: 'paragraph',
        content: [{ type: 'text', text: `Figure ${n}: ${text}` }],
      })
      .run()
    setStatus('dirty')
  }, [])

  const addComment = useCallback(() => {
    const ed = editorRef.current
    if (!ed) return
    const { from, to } = ed.state.selection
    if (from === to) {
      alert('Select some text first, then add a comment.')
      return
    }
    const text = window.prompt('Comment:')
    if (!text) return
    const quote = ed.state.doc.textBetween(from, to, ' ')
    ed.chain().focus().setHighlight({ color: '#fef08a' }).run()
    const comment = { id: uid(), text, quote, from, to }
    updateDoc({ comments: [...docStateRef.current.comments, comment] })
    setCommentsPane(true)
    setStatus('dirty')
  }, [updateDoc])

  const deleteComment = useCallback(
    (id: string) => {
      const ed = editorRef.current
      const c = docStateRef.current.comments.find((x) => x.id === id)
      if (c && ed) {
        ed
          .chain()
          .focus()
          .setTextSelection({ from: Math.min(c.from, ed.state.doc.content.size), to: Math.min(c.to, ed.state.doc.content.size) })
          .unsetHighlight()
          .run()
      }
      updateDoc({ comments: docStateRef.current.comments.filter((x) => x.id !== id) })
      setStatus('dirty')
    },
    [updateDoc],
  )

  const exportAsDocx = useCallback(() => {
    const doc = buildDoc()
    if (!doc) return
    void exportDocx(doc)
      .then((blob) => downloadBlob(blob, `${doc.title || 'document'}.docx`))
      .catch(() => alert('Could not export the .docx file.'))
  }, [buildDoc])

  const renameTitle = useCallback(() => {
    const name = window.prompt('Rename document:', docStateRef.current.title)
    if (name && name.trim()) {
      updateDoc({ title: name.trim() })
      setStatus('dirty')
      saveNowRef.current()
    }
  }, [updateDoc])

  const saveAsPrompt = useCallback(() => {
    const doc = buildDoc()
    if (!doc) return
    const name = window.prompt('Save a copy as:', doc.title)
    if (name && name.trim()) props.onSaveAs(doc, name.trim())
  }, [buildDoc, props])

  const fitZoom = useCallback(() => {
    const wrap = workspaceRef.current
    if (!wrap) return
    const sheetW = dims.width * (ps.twoPage ? 2 : 1) + 48
    const pct = Math.max(25, Math.min(200, Math.floor(((wrap.clientWidth - 32) / sheetW) * 100)))
    updatePageSettings({ zoom: pct })
  }, [dims.width, ps.twoPage, updatePageSettings])

  const applyStyle = useCallback(
    (style: string) => {
      const ed = editorRef.current
      if (!ed) return
      if (style === 'Normal') ed.chain().focus().setParagraph().run()
      else if (style === 'Heading 1') ed.chain().focus().setHeading({ level: 1 }).run()
      else if (style === 'Heading 2') ed.chain().focus().setHeading({ level: 2 }).run()
      else if (style === 'Title') ed.chain().focus().setNode('title').run()
      else if (style === 'Subtitle') ed.chain().focus().setNode('subtitle').run()
    },
    [],
  )

  const currentStyle = editor
    ? editor.isActive('title')
      ? 'Title'
      : editor.isActive('subtitle')
        ? 'Subtitle'
        : editor.isActive('heading', { level: 1 })
          ? 'Heading 1'
          : editor.isActive('heading', { level: 2 })
            ? 'Heading 2'
            : 'Normal'
    : 'Normal'

  const fontFamily = (editor?.getAttributes('textStyle').fontFamily as string | undefined) ?? 'Calibri'
  const fontSizeRaw = (editor?.getAttributes('textStyle').fontSize as string | undefined) ?? '16px'
  const fontSize = parseInt(fontSizeRaw, 10) || 16
  const lineHeight = (editor?.getAttributes('paragraph').lineHeight as string | undefined) ?? '1.15'
  const paraIndent = (editor?.getAttributes('paragraph').indent as number | undefined) ?? 0
  const paraBorder = (editor?.getAttributes('paragraph').paraBorder as boolean | undefined) ?? false

  const headings = collectHeadings()
  const docCounts = {
    headings: headings.length,
    images: (() => {
      let n = 0
      editor?.state.doc.descendants((node) => {
        if (node.type.name === 'image') n++
      })
      return n
    })(),
    tables: (() => {
      let n = 0
      editor?.state.doc.descendants((node) => {
        if (node.type.name === 'table') n++
      })
      return n
    })(),
  }

  const statusText = status === 'saved' ? 'Saved ✓' : status === 'saving' ? 'Saving…' : 'Unsaved changes'
  const inTable = editor?.isActive('table') ?? false
  const sheetWidth = ps.view === 'web' ? '100%' : ps.twoPage ? dims.width * 2 + 24 : dims.width

  if (!editor) {
    return (
      <div className="flex h-screen items-center justify-center bg-[#c8ccd0] text-sm text-slate-600">
        Opening editor…
      </div>
    )
  }

  const chain = () => editor.chain().focus()

  return (
    <div className="editor-root fixed inset-0 flex flex-col overflow-hidden bg-[#c8ccd0]">
      {/* ---------- Title bar ---------- */}
      <div className="titlebar no-print flex h-9 shrink-0 items-center gap-2 px-2 text-white">
        <span className="text-sm font-semibold">My Word</span>
        <button
          type="button"
          onClick={renameTitle}
          className="max-w-[40vw] truncate rounded px-1.5 py-0.5 text-xs hover:bg-white/20"
          title="Click to rename"
        >
          {docState.title}.docx
        </button>
        <span className="text-[11px] text-blue-100">{statusText}</span>
        <div className="ml-1 flex items-center gap-0.5">
          <button
            type="button"
            className="rounded p-1 hover:bg-white/20 disabled:opacity-40"
            title="Undo"
            disabled={!editor.can().undo()}
            onClick={() => chain().undo().run()}
          >
            <Undo2 size={15} />
          </button>
          <button
            type="button"
            className="rounded p-1 hover:bg-white/20 disabled:opacity-40"
            title="Redo"
            disabled={!editor.can().redo()}
            onClick={() => chain().redo().run()}
          >
            <Redo2 size={15} />
          </button>
          <button type="button" className="rounded p-1 hover:bg-white/20" title="Save" onClick={() => saveNow()}>
            <Save size={15} />
          </button>
        </div>
        <div className="flex-1" />
        <button
          type="button"
          className="rounded px-2 py-0.5 text-xs hover:bg-white/20"
          onClick={() => {
            if (status === 'dirty') saveNow()
            props.onClose()
          }}
          title="Close document"
        >
          ✕ Close
        </button>
      </div>

      {/* ---------- Ribbon tabs ---------- */}
      <div className="no-print relative shrink-0 bg-[#eef1f5]">
        <div className="flex items-end gap-0.5 overflow-x-auto px-2 pt-1">
          {TABS.map((t) => (
            <button
              key={t}
              type="button"
              className={`shrink-0 rounded-t px-3 py-1 text-[11px] font-medium ${
                activeTab === t && !fileOpen ? 'ribbon-tab-active' : 'text-slate-700 hover:bg-slate-200'
              }`}
              onClick={() => {
                if (t === 'FILE') {
                  setFileOpen((v) => !v)
                } else {
                  setFileOpen(false)
                  setActiveTab(t)
                }
              }}
            >
              {t}
            </button>
          ))}
          <div className="flex-1" />
          <button
            type="button"
            className="mb-0.5 shrink-0 rounded border border-slate-300 bg-white px-2 py-0.5 text-[11px] text-slate-700"
            onClick={props.onOpenPdf}
            title="Open the PDF workspace"
          >
            PDF
          </button>
        </div>

        {/* FILE dropdown panel */}
        {fileOpen && (
          <div className="absolute left-2 top-full z-40 w-72 rounded-b border border-slate-300 bg-white shadow-xl">
            <div className="p-1 text-sm">
              <button type="button" className="flex w-full items-center gap-2 rounded px-2 py-1.5 text-left hover:bg-blue-50" onClick={() => { setFileOpen(false); props.onNew() }}>
                <FilePlus size={15} /> New Document
              </button>
              <button type="button" className="flex w-full items-center gap-2 rounded px-2 py-1.5 text-left hover:bg-blue-50" onClick={() => { setFileOpen(false); importInputRef.current?.click() }}>
                <FolderOpen size={15} /> Open Document (.docx / .html)
              </button>
              <button type="button" className="flex w-full items-center gap-2 rounded px-2 py-1.5 text-left hover:bg-blue-50" onClick={() => { setFileOpen(false); saveNow() }}>
                <Save size={15} /> Save
              </button>
              <button type="button" className="flex w-full items-center gap-2 rounded px-2 py-1.5 text-left hover:bg-blue-50" onClick={() => { setFileOpen(false); saveAsPrompt() }}>
                <Copy size={15} /> Save As…
              </button>
              <button type="button" className="flex w-full items-center gap-2 rounded px-2 py-1.5 text-left hover:bg-blue-50" onClick={() => { setFileOpen(false); exportAsDocx() }}>
                <FileText size={15} /> Export as DOCX
              </button>
              <button type="button" className="flex w-full items-center gap-2 rounded px-2 py-1.5 text-left hover:bg-blue-50" onClick={() => { setFileOpen(false); window.print() }}>
                <Printer size={15} /> Export as PDF / Print…
              </button>
              <button type="button" className="flex w-full items-center gap-2 rounded px-2 py-1.5 text-left hover:bg-blue-50" onClick={() => { setFileOpen(false); window.print() }}>
                <Printer size={15} /> Print
              </button>
              <button type="button" className="flex w-full items-center gap-2 rounded px-2 py-1.5 text-left hover:bg-blue-50" onClick={() => { setFileOpen(false); setDialog('docInfo') }}>
                <Info size={15} /> Document Information
              </button>
            </div>
            <div className="border-t border-slate-200 p-2">
              <div className="mb-1 px-1 text-[11px] font-semibold uppercase text-slate-500">Recent Documents</div>
              {props.recentDocs.length === 0 && (
                <div className="px-2 py-1 text-xs text-slate-500">No recent documents.</div>
              )}
              {props.recentDocs.map((d) => (
                <button
                  key={d.id}
                  type="button"
                  className="block w-full truncate rounded px-2 py-1 text-left text-xs hover:bg-blue-50"
                  title={d.title}
                  onClick={() => {
                    setFileOpen(false)
                    props.onOpenRecent(d.id)
                  }}
                >
                  {d.title}
                  <span className="ml-2 text-[10px] text-slate-400">
                    {new Date(d.updatedAt).toLocaleDateString()}
                  </span>
                </button>
              ))}
            </div>
          </div>
        )}

        {/* Ribbon content row */}
        {!fileOpen && (
          <div className="flex items-stretch gap-1 overflow-x-auto border-t border-[#b6c4d6] bg-white px-2 py-1">
            {activeTab === 'HOME' && (
              <>
                <RGroup label="Clipboard">
                  <RB label="Cut" onClick={doCut} title="Cut (Ctrl+X)"><Scissors size={18} /></RB>
                  <RB label="Copy" onClick={doCopy} title="Copy (Ctrl+C)"><Copy size={18} /></RB>
                  <RB label="Paste" onClick={doPaste} title="Paste — browser may block clipboard read; use Ctrl+V"><ClipboardPaste size={18} /></RB>
                </RGroup>
                <RGroup label="Font">
                  <div className="flex flex-col gap-1">
                    <div className="flex items-center gap-1">
                      <select
                        className={selectCls}
                        style={{ width: 118 }}
                        value={fontFamily}
                        onChange={(e) => chain().setFontFamily(e.target.value).run()}
                        title="Font family"
                      >
                        {FONT_FAMILIES.map((f) => (
                          <option key={f} value={f}>{f}</option>
                        ))}
                      </select>
                      <select
                        className={selectCls}
                        style={{ width: 56 }}
                        value={fontSize}
                        onChange={(e) => chain().setFontSize(`${e.target.value}px`).run()}
                        title="Font size"
                      >
                        {FONT_SIZES.map((s) => (
                          <option key={s} value={s}>{s}</option>
                        ))}
                      </select>
                      <button type="button" className="rounded border border-slate-300 px-1.5 py-0.5 text-xs" title="Increase font size" onClick={() => chain().setFontSize(`${fontSize + 2}px`).run()}>A+</button>
                      <button type="button" className="rounded border border-slate-300 px-1.5 py-0.5 text-xs" title="Decrease font size" onClick={() => chain().setFontSize(`${Math.max(8, fontSize - 2)}px`).run()}>A−</button>
                    </div>
                    <div className="flex items-center gap-0.5">
                      <button type="button" className={`rounded p-1 hover:bg-slate-100 ${editor.isActive('bold') ? 'bg-blue-200' : ''}`} title="Bold (Ctrl+B)" onClick={() => chain().toggleBold().run()}><Bold size={15} /></button>
                      <button type="button" className={`rounded p-1 hover:bg-slate-100 ${editor.isActive('italic') ? 'bg-blue-200' : ''}`} title="Italic (Ctrl+I)" onClick={() => chain().toggleItalic().run()}><Italic size={15} /></button>
                      <button type="button" className={`rounded p-1 hover:bg-slate-100 ${editor.isActive('underline') ? 'bg-blue-200' : ''}`} title="Underline (Ctrl+U)" onClick={() => chain().toggleUnderline().run()}><UnderlineIcon size={15} /></button>
                      <button type="button" className={`rounded p-1 hover:bg-slate-100 ${editor.isActive('strike') ? 'bg-blue-200' : ''}`} title="Strikethrough" onClick={() => chain().toggleStrike().run()}><Strikethrough size={15} /></button>
                      <button type="button" className={`rounded p-1 hover:bg-slate-100 ${editor.isActive('subscript') ? 'bg-blue-200' : ''}`} title="Subscript" onClick={() => chain().toggleSubscript().run()}><SubscriptIcon size={15} /></button>
                      <button type="button" className={`rounded p-1 hover:bg-slate-100 ${editor.isActive('superscript') ? 'bg-blue-200' : ''}`} title="Superscript" onClick={() => chain().toggleSuperscript().run()}><SuperscriptIcon size={15} /></button>
                      <label className="flex cursor-pointer items-center gap-0.5 rounded p-1 hover:bg-slate-100" title="Highlight colour">
                        <Highlighter size={15} />
                        <input type="color" defaultValue="#fde047" className="h-4 w-5 cursor-pointer border-0 bg-transparent p-0" onChange={(e) => chain().toggleHighlight({ color: e.target.value }).run()} />
                      </label>
                      <label className="flex cursor-pointer items-center gap-0.5 rounded p-1 hover:bg-slate-100" title="Font colour">
                        <Palette size={15} />
                        <input type="color" defaultValue="#1f2937" className="h-4 w-5 cursor-pointer border-0 bg-transparent p-0" onChange={(e) => chain().setColor(e.target.value).run()} />
                      </label>
                      <button type="button" className="rounded px-1.5 py-0.5 text-[11px] hover:bg-slate-100" title="Clear formatting" onClick={() => chain().unsetAllMarks().clearNodes().run()}>Clear</button>
                    </div>
                  </div>
                </RGroup>
                <RGroup label="Paragraph">
                  <RB label="Bullets" onClick={() => chain().toggleBulletList().run()} active={editor.isActive('bulletList')}><List size={18} /></RB>
                  <RB label="Numbering" onClick={() => chain().toggleOrderedList().run()} active={editor.isActive('orderedList')}><ListOrdered size={18} /></RB>
                  <RB label="Left" onClick={() => chain().setTextAlign('left').run()} active={editor.isActive({ textAlign: 'left' })}><AlignLeft size={18} /></RB>
                  <RB label="Center" onClick={() => chain().setTextAlign('center').run()} active={editor.isActive({ textAlign: 'center' })}><AlignCenter size={18} /></RB>
                  <RB label="Right" onClick={() => chain().setTextAlign('right').run()} active={editor.isActive({ textAlign: 'right' })}><AlignRight size={18} /></RB>
                  <RB label="Justify" onClick={() => chain().setTextAlign('justify').run()} active={editor.isActive({ textAlign: 'justify' })}><AlignJustify size={18} /></RB>
                </RGroup>
                <RGroup label="Spacing">
                  <div className="flex flex-col gap-1">
                    <select className={selectCls} value={lineHeight} onChange={(e) => chain().updateAttributes('paragraph', { lineHeight: e.target.value }).run()} title="Line spacing">
                      <option value="1">1.0</option>
                      <option value="1.15">1.15</option>
                      <option value="1.5">1.5</option>
                      <option value="2">2.0</option>
                    </select>
                    <div className="flex gap-1">
                      <button type="button" className="rounded border border-slate-300 px-1.5 py-0.5 text-[11px]" title="Decrease indent" onClick={() => chain().updateAttributes('paragraph', { indent: Math.max(0, paraIndent - 1) }).run()}>⇤ Indent−</button>
                      <button type="button" className="rounded border border-slate-300 px-1.5 py-0.5 text-[11px]" title="Increase indent" onClick={() => chain().updateAttributes('paragraph', { indent: Math.min(5, paraIndent + 1) }).run()}>Indent+ ⇥</button>
                      <button type="button" className={`rounded border px-1.5 py-0.5 text-[11px] ${paraBorder ? 'border-blue-500 bg-blue-100' : 'border-slate-300'}`} title="Paragraph border" onClick={() => chain().updateAttributes('paragraph', { paraBorder: !paraBorder }).run()}>Border</button>
                    </div>
                  </div>
                </RGroup>
                <RGroup label="Styles">
                  <select className={selectCls} value={currentStyle} onChange={(e) => applyStyle(e.target.value)} title="Paragraph style">
                    <option>Normal</option>
                    <option>Heading 1</option>
                    <option>Heading 2</option>
                    <option>Title</option>
                    <option>Subtitle</option>
                  </select>
                </RGroup>
              </>
            )}

            {activeTab === 'INSERT' && (
              <>
                <RGroup label="Tables">
                  <RB label="Table" onClick={() => setDialog('table')}><TableIcon size={18} /></RB>
                </RGroup>
                <RGroup label="Illustrations">
                  <RB label="Picture" onClick={() => pictureInputRef.current?.click()}><ImageIcon size={18} /></RB>
                </RGroup>
                <RGroup label="Links">
                  <RB
                    label="Hyperlink"
                    onClick={() => {
                      const url = window.prompt('Link URL (https://…):', 'https://')
                      if (url) chain().setLink({ href: url }).run()
                    }}
                  >
                    <Link2 size={18} />
                  </RB>
                  <RB label="Unlink" onClick={() => chain().unsetLink().run()}><Link2Off size={18} /></RB>
                </RGroup>
                <RGroup label="Pages">
                  <RB label="Page Break" onClick={() => chain().insertPageBreak().run()}><FileText size={18} /></RB>
                  <RB label="Page Number" onClick={() => chain().insertContent({ type: 'pageNumber' }).run()}><Hash size={18} /></RB>
                </RGroup>
                <RGroup label="Header & Footer">
                  <RB
                    label={docState.showHeaderFooter ? 'Hide H/F' : 'Header/Footer'}
                    onClick={() => updateDoc({ showHeaderFooter: !docState.showHeaderFooter })}
                    active={docState.showHeaderFooter}
                  >
                    <Type size={18} />
                  </RB>
                  <div className="flex flex-col gap-1">
                    <input
                      className={selectCls}
                      style={{ width: 150 }}
                      placeholder="Header text"
                      value={docState.headerText}
                      disabled={!docState.showHeaderFooter}
                      onChange={(e) => updateDoc({ headerText: e.target.value })}
                    />
                    <input
                      className={selectCls}
                      style={{ width: 150 }}
                      placeholder="Footer text"
                      value={docState.footerText}
                      disabled={!docState.showHeaderFooter}
                      onChange={(e) => updateDoc({ footerText: e.target.value })}
                    />
                  </div>
                </RGroup>
                <RGroup label="Text">
                  <RB
                    label="Date & Time"
                    onClick={() => chain().insertContent(new Date().toLocaleString()).run()}
                  >
                    <BookOpen size={18} />
                  </RB>
                  <RB label="Symbols" onClick={() => setDialog('symbols')}><Plus size={18} /></RB>
                </RGroup>
              </>
            )}

            {activeTab === 'PAGE LAYOUT' && (
              <>
                <RGroup label="Margins">
                  {(['narrow', 'normal', 'wide'] as const).map((m) => (
                    <RB key={m} label={m[0].toUpperCase() + m.slice(1)} onClick={() => updatePageSettings({ margins: m })} active={ps.margins === m}>
                      <Columns3 size={18} />
                    </RB>
                  ))}
                </RGroup>
                <RGroup label="Orientation">
                  <RB label="Portrait" onClick={() => updatePageSettings({ orientation: 'portrait' })} active={ps.orientation === 'portrait'}><FileText size={18} /></RB>
                  <RB label="Landscape" onClick={() => updatePageSettings({ orientation: 'landscape' })} active={ps.orientation === 'landscape'}><FileText size={18} /></RB>
                </RGroup>
                <RGroup label="Page Size">
                  <select className={selectCls} value={ps.size} onChange={(e) => updatePageSettings({ size: e.target.value as PageSettings['size'] })} title="Page size">
                    {(['A4', 'A3', 'Letter', 'Legal'] as const).map((s) => (
                      <option key={s} value={s}>{s}</option>
                    ))}
                  </select>
                </RGroup>
                <RGroup label="Columns">
                  {([1, 2, 3] as const).map((c) => (
                    <RB key={c} label={`${c} col`} onClick={() => updatePageSettings({ columns: c })} active={ps.columns === c}>
                      <Columns3 size={18} />
                    </RB>
                  ))}
                </RGroup>
                <RGroup label="Page">
                  <RB label="Page Border" onClick={() => updatePageSettings({ border: !ps.border })} active={ps.border}><FileText size={18} /></RB>
                  <label className="flex cursor-pointer flex-col items-center gap-0.5" title="Page background colour">
                    <input type="color" value={ps.background} className="h-6 w-8 cursor-pointer border border-slate-300 bg-transparent p-0" onChange={(e) => updatePageSettings({ background: e.target.value })} />
                    <span className="text-[11px]">Background</span>
                  </label>
                </RGroup>
                <RGroup label="Paragraph">
                  <div className="flex flex-col gap-1">
                    <select
                      className={selectCls}
                      value={ps.paragraphSpacing}
                      onChange={(e) => {
                        const v = e.target.value as 'compact' | 'normal' | 'relaxed'
                        updatePageSettings({ paragraphSpacing: v })
                        chain().updateAttributes('paragraph', { paraSpacing: v }).run()
                      }}
                      title="Paragraph spacing"
                    >
                      <option value="compact">Compact</option>
                      <option value="normal">Normal</option>
                      <option value="relaxed">Relaxed</option>
                    </select>
                    <select
                      className={selectCls}
                      value={ps.indent}
                      onChange={(e) => {
                        const v = Number(e.target.value)
                        updatePageSettings({ indent: v })
                        chain().updateAttributes('paragraph', { indent: v }).run()
                      }}
                      title="Indentation"
                    >
                      <option value={0}>Indent: None</option>
                      <option value={1}>Indent: 1</option>
                      <option value={2}>Indent: 2</option>
                    </select>
                  </div>
                </RGroup>
              </>
            )}

            {activeTab === 'REFERENCES' && (
              <>
                <RGroup label="Contents">
                  <RB label="Table of Contents" onClick={() => setDialog('toc')}><BookOpen size={18} /></RB>
                </RGroup>
                <RGroup label="Notes">
                  <RB label="Footnote" onClick={() => insertFootnoteLike('footnotes')}><Plus size={18} /></RB>
                  <RB label="Endnote" onClick={() => insertFootnoteLike('endnotes')}><Plus size={18} /></RB>
                </RGroup>
                <RGroup label="Citations">
                  <RB label="Add Citation" onClick={() => setDialog('citation')}><BookOpen size={18} /></RB>
                  <RB label="Bibliography" onClick={insertBibliography}><List size={18} /></RB>
                </RGroup>
                <RGroup label="Captions">
                  <RB label="Add Caption" onClick={insertCaption}><Type size={18} /></RB>
                </RGroup>
              </>
            )}

            {activeTab === 'REVIEW' && (
              <>
                <RGroup label="Proofing">
                  <RB
                    label={spellcheck ? 'Spellcheck On' : 'Spellcheck Off'}
                    onClick={() => setSpellcheck((v) => !v)}
                    active={spellcheck}
                    title="Browser spellcheck (underlines misspelled words in the browser)"
                  >
                    <Search size={18} />
                  </RB>
                  <RB label="Word Count" onClick={() => setDialog('wordCount')}><Hash size={18} /></RB>
                </RGroup>
                <RGroup label="Comments">
                  <RB label="Add Comment" onClick={addComment}><MessageSquare size={18} /></RB>
                  <RB label="Comments" onClick={() => setCommentsPane((v) => !v)} active={commentsPane}>
                    <MessageSquare size={18} />
                  </RB>
                </RGroup>
                <RGroup label="Editing">
                  <RB label="Find & Replace" onClick={() => setDialog('find')}><Search size={18} /></RB>
                  <RB
                    label="Track Changes"
                    onClick={() => {
                      const next = !trackChanges
                      setTrackChanges(next)
                      editor.commands.setTrackChanges(next)
                    }}
                    active={trackChanges}
                    title="New typed text is shown blue and underlined"
                  >
                    <Type size={18} />
                  </RB>
                </RGroup>
                <RGroup label="Review">
                  <RB label="Document Review" onClick={() => setDialog('review')}><Info size={18} /></RB>
                </RGroup>
              </>
            )}

            {activeTab === 'VIEW' && (
              <>
                <RGroup label="Layout">
                  <RB label="Print Layout" onClick={() => updatePageSettings({ view: 'print' })} active={ps.view === 'print'}><FileText size={18} /></RB>
                  <RB label="Web Layout" onClick={() => updatePageSettings({ view: 'web' })} active={ps.view === 'web'}><FileText size={18} /></RB>
                </RGroup>
                <RGroup label="Show">
                  <RB label="Ruler" onClick={() => updatePageSettings({ showRuler: !ps.showRuler })} active={ps.showRuler}><Minus size={18} /></RB>
                  <RB label="Navigation" onClick={() => setNavPane((v) => !v)} active={navPane}><BookOpen size={18} /></RB>
                </RGroup>
                <RGroup label="Zoom">
                  <RB label="Zoom Out" onClick={() => updatePageSettings({ zoom: Math.max(25, zoom - 10) })}><ZoomOut size={18} /></RB>
                  <RB label={`${zoom}%`} onClick={() => updatePageSettings({ zoom: 100 })} title="Reset zoom to 100%"><ZoomIn size={18} /></RB>
                  <RB label="Zoom In" onClick={() => updatePageSettings({ zoom: Math.min(200, zoom + 10) })}><ZoomIn size={18} /></RB>
                  <RB label="Fit" onClick={fitZoom} title="Fit page to window width"><Maximize size={18} /></RB>
                </RGroup>
                <RGroup label="Window">
                  <RB
                    label="Two-Page"
                    onClick={() => updatePageSettings({ twoPage: !ps.twoPage })}
                    active={ps.twoPage}
                    title="Approximation: the sheet becomes double-width with two text columns"
                  >
                    <Columns3 size={18} />
                  </RB>
                  <RB label="Full Screen" onClick={() => void workspaceRef.current?.requestFullscreen()}><Maximize size={18} /></RB>
                </RGroup>
              </>
            )}
          </div>
        )}
      </div>

      {/* ---------- Image / table toolbars ---------- */}
      {(imageSelected || inTable) && (
        <div className="no-print flex shrink-0 flex-wrap items-center gap-2 border-b border-slate-300 bg-blue-50 px-3 py-1 text-xs">
          {imageSelected && (
            <>
              <span className="font-semibold text-slate-700">Image:</span>
              {(['25%', '50%', '75%', '100%'] as const).map((w) => (
                <button
                  key={w}
                  type="button"
                  className="rounded border border-slate-300 bg-white px-2 py-0.5 hover:bg-slate-100"
                  onClick={() => editor.commands.updateAttributes('image', { width: w })}
                >
                  {w}
                </button>
              ))}
              <span className="text-slate-500">Align:</span>
              <button type="button" className="rounded border border-slate-300 bg-white px-2 py-0.5 hover:bg-slate-100" onClick={() => chain().setTextAlign('left').run()}>Left</button>
              <button type="button" className="rounded border border-slate-300 bg-white px-2 py-0.5 hover:bg-slate-100" onClick={() => chain().setTextAlign('center').run()}>Center</button>
              <button type="button" className="rounded border border-slate-300 bg-white px-2 py-0.5 hover:bg-slate-100" onClick={() => chain().setTextAlign('right').run()}>Right</button>
            </>
          )}
          {inTable && (
            <>
              <span className="font-semibold text-slate-700">Table:</span>
              <button type="button" className="rounded border border-slate-300 bg-white px-2 py-0.5 hover:bg-slate-100" onClick={() => chain().addRowBefore().run()}>+ Row above</button>
              <button type="button" className="rounded border border-slate-300 bg-white px-2 py-0.5 hover:bg-slate-100" onClick={() => chain().addRowAfter().run()}>+ Row below</button>
              <button type="button" className="rounded border border-slate-300 bg-white px-2 py-0.5 hover:bg-slate-100" onClick={() => chain().deleteRow().run()}>Delete row</button>
              <button type="button" className="rounded border border-slate-300 bg-white px-2 py-0.5 hover:bg-slate-100" onClick={() => chain().addColumnBefore().run()}>+ Column left</button>
              <button type="button" className="rounded border border-slate-300 bg-white px-2 py-0.5 hover:bg-slate-100" onClick={() => chain().addColumnAfter().run()}>+ Column right</button>
              <button type="button" className="rounded border border-slate-300 bg-white px-2 py-0.5 hover:bg-slate-100" onClick={() => chain().deleteColumn().run()}>Delete column</button>
              <button type="button" className="rounded border border-slate-300 bg-white px-2 py-0.5 hover:bg-slate-100" onClick={() => chain().deleteTable().run()}>Delete table</button>
            </>
          )}
        </div>
      )}

      {/* ---------- Workspace ---------- */}
      <div className="flex min-h-0 flex-1">
        {navPane && (
          <div className="no-print w-60 shrink-0 overflow-auto border-r border-slate-300 bg-white p-3">
            <div className="mb-2 flex items-center justify-between">
              <span className="text-xs font-semibold text-slate-700">Navigation</span>
              <button type="button" onClick={() => setNavPane(false)} className="rounded p-0.5 hover:bg-slate-100"><X size={14} /></button>
            </div>
            {headings.length === 0 && <p className="text-xs text-slate-500">No headings yet.</p>}
            {headings.map((h, i) => (
              <button
                key={`${h.pos}-${i}`}
                type="button"
                className="block w-full truncate rounded px-1 py-1 text-left text-xs hover:bg-blue-50"
                style={{ paddingLeft: `${(h.level - 1) * 12 + 4}px`, fontWeight: h.level === 1 ? 600 : 400 }}
                title={h.text}
                onClick={() => jumpToPos(h.pos + 1)}
              >
                {h.text}
              </button>
            ))}
          </div>
        )}

        <div
          ref={workspaceRef}
          className="relative min-w-0 flex-1 overflow-auto"
          style={{ background: ps.background, touchAction: 'pan-y' }}
        >
          {ps.showRuler && ps.view === 'print' && (
            <div className="no-print sticky top-0 z-10 mx-auto" style={{ width: sheetWidth === '100%' ? '100%' : sheetWidth, maxWidth: '100%' }}>
              <div className="ruler-h h-5 w-full border-b border-slate-300" />
            </div>
          )}
          <div className="flex justify-center py-6" style={zoom === 100 ? undefined : { zoom: zoom / 100 }}>
            <div
              className="print-sheet relative bg-white"
              style={{
                width: sheetWidth,
                minHeight: ps.view === 'print' ? dims.height : undefined,
                padding: ps.view === 'print' ? margin : 28,
                border: ps.border ? '2px solid #111827' : undefined,
                boxShadow: '0 2px 12px rgba(0,0,0,0.25)',
              }}
            >
              {docState.showHeaderFooter && docState.headerText && (
                <div className="mb-3 border-b border-slate-300 pb-2 text-center text-xs text-slate-500">
                  {docState.headerText}
                </div>
              )}
              <div
                ref={contentWrapRef}
                style={{
                  columnCount: ps.twoPage ? 2 : ps.columns,
                  columnGap: ps.columns > 1 || ps.twoPage ? 32 : undefined,
                }}
              >
                <EditorContent editor={editor} />
              </div>

              {docState.footnotes.length > 0 && (
                <div className="mt-6 border-t border-slate-300 pt-2 text-xs text-slate-600">
                  <div className="mb-1 font-semibold">Footnotes</div>
                  {docState.footnotes.map((fn, i) => (
                    <button
                      key={fn.id}
                      type="button"
                      className="block w-full rounded px-1 py-0.5 text-left hover:bg-blue-50"
                      onClick={() => {
                        const pos = footnotePosRef.current[fn.id]
                        if (pos !== undefined) jumpToPos(pos)
                      }}
                      title="Jump to the footnote marker"
                    >
                      [{i + 1}] {fn.text}
                    </button>
                  ))}
                </div>
              )}
              {docState.endnotes.length > 0 && (
                <div className="mt-4 border-t border-slate-300 pt-2 text-xs text-slate-600">
                  <div className="mb-1 font-semibold">Endnotes</div>
                  {docState.endnotes.map((en, i) => (
                    <button
                      key={en.id}
                      type="button"
                      className="block w-full rounded px-1 py-0.5 text-left hover:bg-blue-50"
                      onClick={() => {
                        const pos = footnotePosRef.current[en.id]
                        if (pos !== undefined) jumpToPos(pos)
                      }}
                      title="Jump to the endnote marker"
                    >
                      [{i + 1}] {en.text}
                    </button>
                  ))}
                </div>
              )}
              {docState.showHeaderFooter && docState.footerText && (
                <div className="mt-6 border-t border-slate-300 pt-2 text-center text-xs text-slate-500">
                  {docState.footerText}
                </div>
              )}
            </div>
          </div>
        </div>

        {commentsPane && (
          <div className="no-print w-64 shrink-0 overflow-auto border-l border-slate-300 bg-white p-3">
            <div className="mb-2 flex items-center justify-between">
              <span className="text-xs font-semibold text-slate-700">Comments ({docState.comments.length})</span>
              <button type="button" onClick={() => setCommentsPane(false)} className="rounded p-0.5 hover:bg-slate-100"><X size={14} /></button>
            </div>
            {docState.comments.length === 0 && (
              <p className="text-xs text-slate-500">No comments yet. Select text and use Review → Add Comment.</p>
            )}
            {docState.comments.map((c) => (
              <div key={c.id} className="mb-2 rounded border border-yellow-200 bg-yellow-50 p-2">
                <button
                  type="button"
                  className="block w-full text-left text-xs"
                  onClick={() => {
                    editor
                      .chain()
                      .focus()
                      .setTextSelection({
                        from: Math.min(c.from, editor.state.doc.content.size),
                        to: Math.min(c.to, editor.state.doc.content.size),
                      })
                      .run()
                  }}
                  title="Select the commented text"
                >
                  <span className="italic text-slate-500">“{c.quote}”</span>
                  <span className="mt-1 block text-slate-800">{c.text}</span>
                </button>
                <button type="button" className="mt-1 text-[11px] text-red-600 hover:underline" onClick={() => deleteComment(c.id)}>
                  Delete
                </button>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* ---------- Status bar ---------- */}
      <div className="no-print flex h-7 shrink-0 items-center gap-4 border-t border-slate-300 bg-[#eef1f5] px-3 text-[11px] text-slate-600">
        <span>Page {estPages} of {estPages}</span>
        <span>{counts.words} words</span>
        <span>{counts.chars} characters</span>
        <span className="flex-1" />
        <span>{statusText}</span>
        <span>{zoom}%</span>
      </div>

      {/* ---------- Hidden file inputs ---------- */}
      <input
        ref={importInputRef}
        type="file"
        accept=".docx,.html,.htm"
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0]
          if (f) props.onImportFile(f)
          e.target.value = ''
        }}
      />
      <input
        ref={pictureInputRef}
        type="file"
        accept="image/*"
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0]
          if (!f) return
          const reader = new FileReader()
          reader.onload = () => {
            const src = reader.result as string
            editor.chain().focus().setImage({ src }).run()
          }
          reader.readAsDataURL(f)
          e.target.value = ''
        }}
      />

      {/* ---------- Dialogs ---------- */}
      {dialog === 'docInfo' && (
        <Dialog title="Document Information" onClose={() => setDialog(null)}>
          <dl className="space-y-1.5 text-sm">
            <div className="flex justify-between"><dt className="text-slate-500">Title</dt><dd>{docState.title}</dd></div>
            <div className="flex justify-between"><dt className="text-slate-500">Words</dt><dd>{counts.words}</dd></div>
            <div className="flex justify-between"><dt className="text-slate-500">Characters</dt><dd>{counts.chars}</dd></div>
            <div className="flex justify-between"><dt className="text-slate-500">Estimated pages</dt><dd>{estPages}</dd></div>
            <div className="flex justify-between"><dt className="text-slate-500">Created</dt><dd>{new Date(docState.createdAt).toLocaleString()}</dd></div>
            <div className="flex justify-between"><dt className="text-slate-500">Modified</dt><dd>{new Date(docState.updatedAt).toLocaleString()}</dd></div>
            <div className="flex justify-between"><dt className="text-slate-500">Headings</dt><dd>{docCounts.headings}</dd></div>
            <div className="flex justify-between"><dt className="text-slate-500">Images</dt><dd>{docCounts.images}</dd></div>
            <div className="flex justify-between"><dt className="text-slate-500">Tables</dt><dd>{docCounts.tables}</dd></div>
          </dl>
        </Dialog>
      )}

      {dialog === 'wordCount' && (
        <Dialog title="Word Count" onClose={() => setDialog(null)}>
          <dl className="space-y-1.5 text-sm">
            <div className="flex justify-between"><dt className="text-slate-500">Words</dt><dd>{counts.words}</dd></div>
            <div className="flex justify-between"><dt className="text-slate-500">Characters</dt><dd>{counts.chars}</dd></div>
            <div className="flex justify-between"><dt className="text-slate-500">Estimated pages</dt><dd>{estPages}</dd></div>
          </dl>
        </Dialog>
      )}

      {dialog === 'review' && (
        <Dialog title="Document Review" onClose={() => setDialog(null)}>
          <dl className="space-y-1.5 text-sm">
            <div className="flex justify-between"><dt className="text-slate-500">Headings</dt><dd>{docCounts.headings}</dd></div>
            <div className="flex justify-between"><dt className="text-slate-500">Images</dt><dd>{docCounts.images}</dd></div>
            <div className="flex justify-between"><dt className="text-slate-500">Tables</dt><dd>{docCounts.tables}</dd></div>
            <div className="flex justify-between"><dt className="text-slate-500">Comments</dt><dd>{docState.comments.length}</dd></div>
            <div className="flex justify-between"><dt className="text-slate-500">Footnotes</dt><dd>{docState.footnotes.length}</dd></div>
            <div className="flex justify-between"><dt className="text-slate-500">Endnotes</dt><dd>{docState.endnotes.length}</dd></div>
            <div className="flex justify-between"><dt className="text-slate-500">Citations</dt><dd>{docState.citations.length}</dd></div>
            <div className="flex justify-between"><dt className="text-slate-500">Words</dt><dd>{counts.words}</dd></div>
          </dl>
        </Dialog>
      )}

      {dialog === 'toc' && (
        <Dialog title="Table of Contents" onClose={() => setDialog(null)}>
          {headings.length === 0 && <p className="text-sm text-slate-500">No headings in this document yet.</p>}
          <ul className="space-y-1">
            {headings.map((h, i) => (
              <li key={`${h.pos}-${i}`}>
                <button
                  type="button"
                  className="block w-full rounded px-2 py-1 text-left text-sm hover:bg-blue-50"
                  style={{ paddingLeft: `${(h.level - 1) * 16 + 8}px` }}
                  onClick={() => {
                    setDialog(null)
                    jumpToPos(h.pos + 1)
                  }}
                >
                  {h.text}
                </button>
              </li>
            ))}
          </ul>
          <div className="mt-4 flex justify-end">
            <button type="button" className={btnCls} onClick={insertTocBlock}>Insert TOC at cursor</button>
          </div>
        </Dialog>
      )}

      {dialog === 'table' && (
        <Dialog title="Insert Table" onClose={() => setDialog(null)}>
          <div className="space-y-3 text-sm">
            <label className="block">
              Rows
              <input type="number" min={1} max={20} value={tableRows} onChange={(e) => setTableRows(Math.max(1, Math.min(20, Number(e.target.value) || 1)))} className={inputCls} />
            </label>
            <label className="block">
              Columns
              <input type="number" min={1} max={10} value={tableCols} onChange={(e) => setTableCols(Math.max(1, Math.min(10, Number(e.target.value) || 1)))} className={inputCls} />
            </label>
            <div className="flex justify-end gap-2">
              <button type="button" className={btnGhostCls} onClick={() => setDialog(null)}>Cancel</button>
              <button
                type="button"
                className={btnCls}
                onClick={() => {
                  chain().insertTable({ rows: tableRows, cols: tableCols, withHeaderRow: true }).run()
                  setDialog(null)
                }}
              >
                Insert
              </button>
            </div>
          </div>
        </Dialog>
      )}

      {dialog === 'symbols' && (
        <Dialog title="Special Symbols" onClose={() => setDialog(null)}>
          <div className="grid grid-cols-8 gap-1">
            {SYMBOLS.map((s) => (
              <button
                key={s}
                type="button"
                className="rounded border border-slate-200 py-2 text-lg hover:bg-blue-50"
                onClick={() => {
                  chain().insertContent(s).run()
                  setDialog(null)
                }}
              >
                {s}
              </button>
            ))}
          </div>
        </Dialog>
      )}

      {dialog === 'citation' && (
        <Dialog title="Add Citation" onClose={() => setDialog(null)}>
          <div className="space-y-3">
            <label className="block text-sm">
              Author
              <input value={citAuthor} onChange={(e) => setCitAuthor(e.target.value)} className={inputCls} placeholder="e.g. Khan" />
            </label>
            <label className="block text-sm">
              Year
              <input value={citYear} onChange={(e) => setCitYear(e.target.value)} className={inputCls} placeholder="e.g. 2024" />
            </label>
            <label className="block text-sm">
              Title
              <input value={citTitle} onChange={(e) => setCitTitle(e.target.value)} className={inputCls} placeholder="e.g. Biology of the Cell" />
            </label>
            <div className="flex justify-end gap-2">
              <button type="button" className={btnGhostCls} onClick={() => setDialog(null)}>Cancel</button>
              <button type="button" className={btnCls} onClick={addCitation}>Insert “(Author, Year)”</button>
            </div>
          </div>
        </Dialog>
      )}

      {dialog === 'find' && (
        <Dialog title="Find and Replace" onClose={() => setDialog(null)}>
          <div className="space-y-3">
            <label className="block text-sm">
              Find
              <input
                value={findText}
                onChange={(e) => {
                  setFindText(e.target.value)
                  setMatches([])
                  setMatchIdx(-1)
                }}
                className={inputCls}
                placeholder="Text to find"
                autoFocus
              />
            </label>
            <label className="block text-sm">
              Replace with
              <input value={replaceText} onChange={(e) => setReplaceText(e.target.value)} className={inputCls} placeholder="Replacement text" />
            </label>
            <p className="text-xs text-slate-500">
              {matches.length > 0 ? `${matchIdx + 1} of ${matches.length} matches` : findText ? 'No matches found yet — press Find Next.' : 'Type text to search for.'}
            </p>
            <div className="flex flex-wrap justify-end gap-2">
              <button type="button" className={btnGhostCls} onClick={() => runFind(-1)}>Find Previous</button>
              <button type="button" className={btnGhostCls} onClick={() => runFind(1)}>Find Next</button>
              <button type="button" className={btnGhostCls} onClick={replaceCurrent} disabled={matchIdx < 0}>Replace</button>
              <button type="button" className={btnCls} onClick={replaceAll} disabled={!findText}>Replace All</button>
            </div>
            <p className="text-[11px] text-slate-400">Tip: Ctrl+F opens this dialog, Ctrl+H too.</p>
          </div>
        </Dialog>
      )}
    </div>
  )
}
