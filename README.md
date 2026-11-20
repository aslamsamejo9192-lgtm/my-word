# My Word — Online Word Processor

A complete, working web-based word processor inspired by Microsoft Word 2010.
Built with **React + TypeScript + Vite**, **Tiptap** (rich text editing),
**Tailwind CSS v4**, **Mammoth.js** (DOCX import), the **`docx`** library
(DOCX export), **PDF.js** (PDF viewing) and **pdf-lib** (PDF modification).

Documents are stored locally in **IndexedDB** — they work offline and survive
browser refreshes. Nothing is uploaded anywhere unless you configure cloud sync.

## What it does

- **Dashboard** — My Documents, Create Blank Document, Open DOCX, Open PDF,
  Recent Files, search, document cards with last-modified date and file type,
  Rename / Duplicate / Delete on every card.
- **Word 2010-style editor** — blue title bar with Quick Access Toolbar
  (Undo / Redo / Save), ribbon tabs FILE / HOME / INSERT / PAGE LAYOUT /
  REFERENCES / REVIEW / VIEW, grey workspace, white A4 pages, rulers,
  page margins, page numbering, zoom, print/web layout, navigation pane,
  full screen.
- **Editing** — fonts, sizes, bold/italic/underline/strikethrough,
  subscript/superscript, highlight, font colour, clear formatting, bullets,
  numbering, alignment, line spacing, indent, paragraph borders, styles
  (Normal, Heading 1/2, Title, Subtitle), tables, images, hyperlinks,
  page breaks, headers/footers, date/time, symbols, footnotes, endnotes,
  citations + bibliography, table of contents, comments, find & replace,
  track changes (approximation — see limitations), word/character count.
- **Document management** — create / open / save / Save As / rename /
  duplicate / delete, recent documents, search, autosave (~2 s after edits),
  unsaved-changes warning, recovery of the last document after refresh.
- **DOCX** — import real `.docx` files (formatting, tables, headings, lists
  and images preserved as far as the libraries allow), edit them, export a
  real valid `.docx` (verified: valid zip with `word/document.xml`). The
  original imported file is kept in IndexedDB as a backup.
- **PDF Tools** — open PDFs, multi-page viewer with thumbnails, page jump,
  zoom, fit-to-width/page, text search, highlight / underline / strike-through
  annotations, text comments, text boxes, pen drawing, shapes, colour picker,
  undo/redo for annotations, rotate / delete / reorder pages, merge PDFs,
  split, extract pages, page numbers, watermark, download the modified PDF.
- **Keyboard shortcuts** — Ctrl+S save, Ctrl+B/I/U, Ctrl+Z/Y, Ctrl+F find,
  Ctrl+H replace, Ctrl+P print, Ctrl+A select all.
- **Cloud (optional)** — Supabase client is created only when
  `VITE_SUPABASE_URL` and `VITE_SUPABASE_ANON_KEY` are set; otherwise the app
  is fully local. Never put a service-role key in the frontend.

## Run it locally

```bash
npm install
npm run dev      # development server
npm run build    # production build (outputs to dist/)
npm run preview  # serve the production build
```

## Honest limitations

- **Track Changes is a visual approximation**, not Microsoft Word's real
  track-changes: text typed while it is on gets a blue underline mark. There
  is no accept/reject workflow.
- **Existing PDF text cannot be edited.** Annotations are drawn on top of the
  pages and baked in on download. PDF comments are visible yellow note boxes,
  not native pop-up text annotations (pdf-lib has no public TextAnnotation
  helper). Text entry for PDF text boxes uses the browser prompt.
- **DOCX fidelity is not perfect.** Complex Word layouts, macros, SmartArt
  and advanced page formatting may not be preserved. On export: highlight
  colours are mapped to the nearest named Word highlight colour; nested
  tables inside table cells, blockquotes or lists are dropped; page-number
  fields export as nothing.
- **Comment / footnote jump positions** are stored when created and can drift
  after heavy editing.
- **Paste** uses the clipboard API with an execCommand fallback; browsers may
  block programmatic paste, in which case an alert explains it.
- **Two-page view** is a double-width sheet with 2 text columns — an
  approximation of Word's two-page view.
- The production JS bundle is large (~1.3 MB main chunk + a lazy PDF chunk)
  because it ships the full PDF.js worker; a chunk-size warning during build
  is expected.

## Project structure

```
src/
  App.tsx                 App shell: dashboard / editor / PDF routing, recovery
  main.tsx                Entry point
  index.css               Tailwind + Word 2010 styling + print CSS
  components/
    Dashboard.tsx         My Documents, search, cards, rename/delete
    EditorView.tsx        Tiptap editor + full ribbon + dialogs + print
    PdfWorkspace.tsx      PDF viewer, annotations, page management
  lib/
    types.ts              Document + page-settings types
    store.ts              IndexedDB (idb) document storage
    docx.ts               DOCX import (mammoth) + export (docx lib)
    supabase.ts           Optional Supabase client (env-gated)
e2e/
  sample-import.docx      Test fixture for DOCX import
  sample.pdf              Test fixture for the PDF workspace (3 pages)
  smoke.mjs               Playwright smoke test (run from ~/workspace/ov-tools)
```

---

## Setup guide (Roman Urdu) — AI Studio se Vercel tak

### 1. Google AI Studio mein project banana

1. Apne browser mein **Google AI Studio** kholein (aistudio.google.com).
2. **"Create" / "New App"** par click karein.
3. Project ka naam likhein: **My Word**.
4. Jo code diya gaya hai (yeh poora project), use AI Studio mein paste/upload
   kar dein. AI Studio project generate kar dega.

### 2. Source code download karna

1. AI Studio mein project khulne ke baad **"Export" / "Download"** ka button
   dhoondein (usually upar right side hota hai).
2. **ZIP file download** karein aur apne phone/computer par extract karein.
3. Check karein ke andar `package.json`, `src/` folder aur `README.md`
   mojood hon — yeh zaroori files hain.

### 3. GitHub par upload karna

1. **github.com** par jayein aur **New repository** banayein, naam: `my-word`.
   Repository ko **Public** rakhein (Vercel free deploy ke liye asaan hai).
2. **"uploading an existing file"** par click karein.
3. Apne extract kiye hue project ke **saare files aur folders** drag-and-drop
   karein (khayal rahe: `node_modules` aur `dist` folders upload NA karein —
   woh build ke waqt khud ban jate hain).
4. Neeche **Commit changes** par click karein.

### 4. Vercel se connect aur deploy karna

1. **vercel.com** par jayein aur **GitHub se sign in** karein.
2. **"Add New…" → "Project"** par click karein.
3. Apni `my-word` repository select karein aur **Import** karein.
4. Settings mein kuch change karne ki zaroorat nahi — Framework **Vite** khud
   detect ho jayega, Build Command `npm run build` aur Output Directory
   `dist` honi chahiye.
5. **Deploy** par click karein. 1–2 minute mein aapko website ka link mil
   jayega, jaise `https://my-word.vercel.app`.
6. Deploy fail ho to Vercel ke **Deployments → Build Logs** mein error dekhein
   — aksar koi file miss hoti hai; dubara upload karke **Redeploy** karein.

### 5. Supabase (optional — zaroori nahi)

Cloud sync **optional** hai; bina is ke app poori tarah local chalta hai.

1. **supabase.com** par free project banayein.
2. Project Settings → API se **Project URL** aur **anon public key** copy karein.
   **Service-role key kabhi bhi frontend mein NA dalain.**
3. Vercel mein apne project ki **Settings → Environment Variables** mein
   yeh do variables add karein:
   - `VITE_SUPABASE_URL` = aapka Project URL
   - `VITE_SUPABASE_ANON_KEY` = aapki anon public key
4. Supabase mein **Row Level Security (RLS)** ON karein taake har user sirf
   apne documents dekh sake.
5. Variables add karne ke baad **Redeploy** karein.

### 6. Update karne ke baad dobara deploy

1. Code mein koi change karein (AI Studio mein ya apne editor mein).
2. Badli hui files ko GitHub repository mein dobara upload karein (same
   tareeqa: upload files → Commit).
3. **Bas!** Vercel GitHub se connect hai, to woh khud-ba-khud naya deploy
   shuru kar dega. Kuch minute mein naya version live ho jayega.
4. Agar auto-deploy na ho to Vercel → **Deployments → ⋯ → Redeploy** karein.
