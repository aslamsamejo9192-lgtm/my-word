import { lazy, Suspense, useCallback, useEffect, useState } from 'react'
import Dashboard from './components/Dashboard'
import EditorView from './components/EditorView'

const PdfWorkspace = lazy(() => import('./components/PdfWorkspace'))
import type { MyDocument } from './lib/types'
import { newDocument, uid } from './lib/types'
import * as store from './lib/store'
import { importDocxFile } from './lib/docx'

type View =
  | { kind: 'dashboard' }
  | { kind: 'editor'; doc: MyDocument }
  | { kind: 'pdf'; file: File | null }

export default function App() {
  const [view, setView] = useState<View>({ kind: 'dashboard' })
  const [docs, setDocs] = useState<MyDocument[]>([])

  const loadDocs = useCallback(async () => {
    setDocs(await store.listDocuments())
  }, [])

  useEffect(() => {
    loadDocs()
  }, [loadDocs])

  // Recovery: after a refresh, reopen the document that was being edited.
  useEffect(() => {
    const lastId = sessionStorage.getItem('myword-last-doc')
    if (lastId) {
      store.getDocument(lastId).then((d) => {
        if (d) setView({ kind: 'editor', doc: d })
      })
    }
  }, [])

  const saveDoc = useCallback(
    async (doc: MyDocument) => {
      await store.putDocument(doc)
      sessionStorage.setItem('myword-last-doc', doc.id)
      setDocs((prev) => {
        const exists = prev.some((d) => d.id === doc.id)
        const next = exists
          ? prev.map((d) => (d.id === doc.id ? doc : d))
          : [doc, ...prev]
        return next.sort((a, b) => b.updatedAt - a.updatedAt)
      })
    },
    [],
  )

  const openDoc = useCallback((doc: MyDocument) => {
    sessionStorage.setItem('myword-last-doc', doc.id)
    setView({ kind: 'editor', doc })
  }, [])

  const createBlank = useCallback(async () => {
    const doc = newDocument()
    await store.putDocument(doc)
    await loadDocs()
    openDoc(doc)
  }, [loadDocs, openDoc])

  const importFile = useCallback(
    async (file: File) => {
      const isDocx = /\.docx$/i.test(file.name)
      if (isDocx) {
        const { html, title } = await importDocxFile(file)
        const doc = newDocument(title)
        doc.origin = 'docx'
        doc.originalFileName = file.name
        doc.originalFileBlob = file
        doc.html = html
        // content stays blank JSON; the editor converts the HTML on open.
        await store.putDocument(doc)
        await loadDocs()
        openDoc(doc)
      } else if (/\.html?$/i.test(file.name)) {
        const html = await file.text()
        const doc = newDocument(file.name.replace(/\.html?$/i, ''))
        doc.html = html
        await store.putDocument(doc)
        await loadDocs()
        openDoc(doc)
      } else {
        alert('Only .docx and .html files can be opened as documents.')
      }
    },
    [loadDocs, openDoc],
  )

  const saveAs = useCallback(
    async (doc: MyDocument, newTitle: string) => {
      const copy: MyDocument = {
        ...doc,
        id: uid(),
        title: newTitle,
        createdAt: Date.now(),
        updatedAt: Date.now(),
      }
      await store.putDocument(copy)
      await loadDocs()
      openDoc(copy)
    },
    [loadDocs, openDoc],
  )

  const renameDoc = useCallback(
    async (id: string, title: string) => {
      const doc = await store.getDocument(id)
      if (!doc) return
      await store.putDocument({ ...doc, title })
      await loadDocs()
    },
    [loadDocs],
  )

  const deleteDoc = useCallback(
    async (id: string) => {
      await store.deleteDocument(id)
      if (view.kind === 'editor' && view.doc.id === id) setView({ kind: 'dashboard' })
      await loadDocs()
    },
    [loadDocs, view],
  )

  const duplicateDoc = useCallback(async (id: string) => {
    await store.duplicateDocument(id)
    await loadDocs()
  }, [loadDocs])

  if (view.kind === 'editor') {
    return (
      <EditorView
        doc={view.doc}
        recentDocs={docs.slice(0, 8).map((d) => ({ id: d.id, title: d.title, updatedAt: d.updatedAt }))}
        onSave={saveDoc}
        onSaveAs={saveAs}
        onClose={() => setView({ kind: 'dashboard' })}
        onNew={createBlank}
        onOpenRecent={async (id) => {
          const d = await store.getDocument(id)
          if (d) openDoc(d)
        }}
        onImportFile={importFile}
        onOpenPdf={() => setView({ kind: 'pdf', file: null })}
      />
    )
  }

  if (view.kind === 'pdf') {
    return (
      <Suspense
        fallback={
          <div className="flex h-full items-center justify-center bg-[#dbe4f0] text-sm text-slate-600">
            Loading PDF Tools…
          </div>
        }
      >
        <PdfWorkspace initialFile={view.file} onClose={() => setView({ kind: 'dashboard' })} />
      </Suspense>
    )
  }

  return (
    <Dashboard
      docs={docs}
      onOpen={openDoc}
      onCreate={createBlank}
      onImportDocx={importFile}
      onOpenPdf={(file) => setView({ kind: 'pdf', file: file ?? null })}
      onRename={renameDoc}
      onDelete={deleteDoc}
      onDuplicate={duplicateDoc}
    />
  )
}
