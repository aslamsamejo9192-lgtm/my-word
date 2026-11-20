import { useMemo, useRef, useState } from 'react'
import {
  Copy,
  FileText,
  FilePlus2,
  FileUp,
  FileType2,
  Pencil,
  Search,
  Trash2,
} from 'lucide-react'
import type { MyDocument } from '../lib/types'

interface DashboardProps {
  docs: MyDocument[]
  onOpen: (doc: MyDocument) => void
  onCreate: () => void
  onImportDocx: (file: File) => void
  onOpenPdf: (file?: File) => void
  onRename: (id: string, title: string) => void
  onDelete: (id: string) => void
  onDuplicate: (id: string) => void
}

function fmtDate(ts: number): string {
  return new Date(ts).toLocaleDateString(undefined, {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
  })
}

export default function Dashboard({
  docs,
  onOpen,
  onCreate,
  onImportDocx,
  onOpenPdf,
  onRename,
  onDelete,
  onDuplicate,
}: DashboardProps) {
  const [query, setQuery] = useState('')
  const docxInputRef = useRef<HTMLInputElement>(null)
  const pdfInputRef = useRef<HTMLInputElement>(null)

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase()
    if (!q) return docs
    return docs.filter((d) => d.title.toLowerCase().includes(q))
  }, [docs, query])

  const recent = docs.slice(0, 5)

  const handleRename = (doc: MyDocument) => {
    const name = window.prompt('Rename document', doc.title)
    if (name && name.trim() && name.trim() !== doc.title) {
      onRename(doc.id, name.trim())
    }
  }

  const handleDelete = (doc: MyDocument) => {
    if (window.confirm(`Delete "${doc.title}"? This cannot be undone.`)) {
      onDelete(doc.id)
    }
  }

  return (
    <div className="h-full overflow-y-auto bg-[#dbe4f0]">
      {/* Top bar */}
      <div className="titlebar px-4 py-3 text-white shadow">
        <div className="mx-auto flex max-w-6xl items-center gap-3">
          <FileText className="h-7 w-7" />
          <div>
            <h1 className="text-xl font-bold leading-tight">My Word</h1>
            <p className="text-xs text-blue-100">Online Word Processor</p>
          </div>
          <div className="ml-auto flex gap-2">
            <button
              onClick={onCreate}
              className="flex items-center gap-2 rounded bg-white px-3 py-2 text-sm font-semibold text-blue-800 shadow hover:bg-blue-50"
            >
              <FilePlus2 className="h-4 w-4" /> New Document
            </button>
            <button
              onClick={() => docxInputRef.current?.click()}
              className="flex items-center gap-2 rounded bg-blue-700 px-3 py-2 text-sm font-semibold text-white shadow hover:bg-blue-600"
            >
              <FileUp className="h-4 w-4" /> Open DOCX
            </button>
            <button
              onClick={() => pdfInputRef.current?.click()}
              className="flex items-center gap-2 rounded bg-blue-700 px-3 py-2 text-sm font-semibold text-white shadow hover:bg-blue-600"
            >
              <FileType2 className="h-4 w-4" /> Open PDF
            </button>
          </div>
        </div>
      </div>

      <input
        ref={docxInputRef}
        type="file"
        accept=".docx,.html,.htm"
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0]
          if (f) onImportDocx(f)
          e.target.value = ''
        }}
      />
      <input
        ref={pdfInputRef}
        type="file"
        accept="application/pdf,.pdf"
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0]
          if (f) onOpenPdf(f)
          e.target.value = ''
        }}
      />

      <div className="mx-auto max-w-6xl px-4 py-6">
        {/* Search */}
        <div className="relative mb-6">
          <Search className="absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Search documents…"
            className="w-full rounded-lg border border-slate-300 bg-white py-2.5 pl-10 pr-4 text-sm shadow-sm outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-200"
          />
        </div>

        {/* Create cards */}
        <div className="mb-8 grid grid-cols-1 gap-4 sm:grid-cols-3">
          <button
            onClick={onCreate}
            className="rounded-xl border border-slate-200 bg-white p-5 text-left shadow-sm transition hover:border-blue-400 hover:shadow"
          >
            <FilePlus2 className="mb-2 h-8 w-8 text-blue-700" />
            <div className="font-semibold text-slate-800">Blank Document</div>
            <div className="text-xs text-slate-500">Start writing from scratch</div>
          </button>
          <button
            onClick={() => docxInputRef.current?.click()}
            className="rounded-xl border border-slate-200 bg-white p-5 text-left shadow-sm transition hover:border-blue-400 hover:shadow"
          >
            <FileUp className="mb-2 h-8 w-8 text-blue-700" />
            <div className="font-semibold text-slate-800">Open DOCX</div>
            <div className="text-xs text-slate-500">Import a Word document to edit</div>
          </button>
          <button
            onClick={() => onOpenPdf()}
            className="rounded-xl border border-slate-200 bg-white p-5 text-left shadow-sm transition hover:border-blue-400 hover:shadow"
          >
            <FileType2 className="mb-2 h-8 w-8 text-blue-700" />
            <div className="font-semibold text-slate-800">Open PDF</div>
            <div className="text-xs text-slate-500">View, annotate and edit PDF files</div>
          </button>
        </div>

        {/* Recent */}
        {recent.length > 0 && (
          <div className="mb-8">
            <h2 className="mb-3 text-sm font-bold uppercase tracking-wide text-slate-600">
              Recent Files
            </h2>
            <div className="flex flex-wrap gap-2">
              {recent.map((d) => (
                <button
                  key={d.id}
                  onClick={() => onOpen(d)}
                  className="rounded-full border border-blue-200 bg-white px-3 py-1.5 text-xs font-medium text-blue-800 shadow-sm hover:bg-blue-50"
                >
                  {d.title}
                </button>
              ))}
            </div>
          </div>
        )}

        {/* My Documents */}
        <h2 className="mb-3 text-sm font-bold uppercase tracking-wide text-slate-600">
          My Documents ({filtered.length})
        </h2>
        {filtered.length === 0 ? (
          <div className="rounded-xl border border-dashed border-slate-300 bg-white/60 p-10 text-center text-sm text-slate-500">
            {docs.length === 0
              ? 'No documents yet. Create a blank document or open a DOCX file to get started.'
              : 'No documents match your search.'}
          </div>
        ) : (
          <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3">
            {filtered.map((d) => (
              <div
                key={d.id}
                className="rounded-xl border border-slate-200 bg-white p-4 shadow-sm transition hover:shadow"
              >
                <button onClick={() => onOpen(d)} className="block w-full text-left">
                  <FileText className="mb-2 h-8 w-8 text-blue-700" />
                  <div className="truncate font-semibold text-slate-800">{d.title}</div>
                  <div className="mt-1 text-xs text-slate-500">
                    {d.origin === 'docx' ? 'Word document (.docx)' : 'My Word document'} ·{' '}
                    {fmtDate(d.updatedAt)}
                  </div>
                  {d.originalFileName && (
                    <div className="truncate text-xs text-slate-400">
                      Imported from {d.originalFileName}
                    </div>
                  )}
                </button>
                <div className="mt-3 flex gap-1 border-t border-slate-100 pt-3">
                  <button
                    onClick={() => handleRename(d)}
                    className="flex items-center gap-1 rounded px-2 py-1 text-xs font-medium text-slate-600 hover:bg-slate-100"
                  >
                    <Pencil className="h-3.5 w-3.5" /> Rename
                  </button>
                  <button
                    onClick={() => onDuplicate(d.id)}
                    className="flex items-center gap-1 rounded px-2 py-1 text-xs font-medium text-slate-600 hover:bg-slate-100"
                  >
                    <Copy className="h-3.5 w-3.5" /> Duplicate
                  </button>
                  <button
                    onClick={() => handleDelete(d)}
                    className="ml-auto flex items-center gap-1 rounded px-2 py-1 text-xs font-medium text-red-600 hover:bg-red-50"
                  >
                    <Trash2 className="h-3.5 w-3.5" /> Delete
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}

        <p className="mt-8 text-center text-xs text-slate-500">
          Documents are stored locally in your browser (IndexedDB) and work offline. Nothing is
          uploaded anywhere unless you choose to use cloud sync.
        </p>
      </div>
    </div>
  )
}
