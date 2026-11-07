# Office One

Word, Excel and PowerPoint in your browser. Everything runs on your machine: documents are stored in your
browser's IndexedDB, nothing is uploaded, and no account or Microsoft Office installation is needed.

Original design and implementation, inspired by the usability of the Office web apps. Not affiliated with Microsoft.

## Quick start

Requires Node.js 22+ (developed on 24; see `.nvmrc`).

```bash
npm install
npm run dev          # http://localhost:5173
```

| Command | What it does |
| --- | --- |
| `npm run build` | Type-check and produce the production build in `dist/` |
| `npm run preview` | Serve the production build locally |
| `npm run check` | Type-check, lint and run the unit tests |
| `npm run e2e` | Playwright journeys against the dev server (uses your installed Google Chrome) |
| `npm run e2e:prod` | Builds, then runs smoke tests against the production bundle with its CSP and service worker |
| `npm run verify` | All of the above |
| `npm run fonts:update` | Regenerate the Google Fonts catalogue from Google's public metadata |

Works in current Chrome, Edge, Safari and Firefox.

## Mac app (.dmg)

```bash
npm run desktop:build            # Apple silicon: release/Office One-1.0.0-arm64.dmg
npm run desktop:build -- x64     # Intel Macs
```

This builds the web app, wraps it in an Electron shell (`desktop/`), signs it ad hoc and produces a drag-to-Applications
disk image. The app serves its bundled files from a private `app://` origin, so storage and fonts behave as in the
browser. It is **not notarised** (that needs a paid Apple Developer ID), so the first launch needs right-click → Open.
Documents live in the app's own storage and are separate from the browser version, so move files across by
downloading from Drive and dropping them into the app. `npm run desktop:icon` regenerates the icon from `public/icon.svg`.

## What's in it

**Home and Drive** — start tiles drawn as the document they create, real previews of recent files (the words, the
cells, the first slide), starter templates, and a shared file manager with grid and list layouts (folders, rename,
duplicate, move, favourites, trash with undo, sort, upload, download). Drop a file anywhere to import it. Press
`⌘K` / `Ctrl K` for a command palette that finds files and runs actions. Light and dark themes, keyboard shortcuts
(press `?`), and offline support once loaded.

**Word** — Tiptap rich-text editor on a page canvas with a six-tab ribbon modelled on Word for the web
(Home, Insert, Layout, References, Review, View). Fonts and sizes, bold/italic/underline/strike, sub/superscript,
colours, highlight, case change, lists and checklists, alignment, line spacing, styles, tables, pictures, links,
symbols and emoji, find and replace (case / whole word), margins, paper size, orientation and columns, navigation
pane, table of contents, word count, format painter, zoom, print. Import DOCX/TXT; export DOCX and PDF.

**Excel** — virtualised grid (tested with 20,000 rows) with formula bar, name box, multiple sheets, freeze panes,
sort, conditional formatting (rules, top/bottom, duplicates, colour scales), charts (column, bar, line, area, pie),
find and replace, per-cell fonts, number formats and decimals, fill, copy/paste with reference shifting, row and
column insert/delete that rewrites formulas, and function autocomplete. Import/export CSV and XLSX.
**About 200 functions**: math, statistics, logical, text, date/time, lookup (VLOOKUP, HLOOKUP, XLOOKUP, INDEX, MATCH),
financial (PMT, NPV, IRR…), information and engineering, plus dynamic arrays (SORT, FILTER, UNIQUE, SEQUENCE) that
spill into neighbouring cells. Incremental recalculation with dependency tracking.

**PowerPoint** — slide editor with rich text inside boxes, shapes, images, snapping, alignment and distribution,
multi-select (shift-click, marquee, select all), grouping, layer order, undo/redo, speaker notes, a library of 22 themes
(light, dark and colourful; separate heading and body fonts, second accent, gradient and pattern backgrounds,
per-layout decoration; live-preview gallery in the Design tab), layouts, and ten real slide transitions (fade, push, wipe, split, reveal, cover, uncover, zoom, dissolve) with
direction and duration. Slideshow with presenter view (notes, next slide, timer). Export PPTX (with notes) and PDF.

**Fonts** — the whole Google Fonts library (1,959 families) plus fonts installed on your device, in one searchable,
category-filtered picker. Google families load only when previewed or used, and a font is only usable once it has
been verified to render.

## Design

Ink on mist: blue-tinted ink text on a cool desk, with white reserved for things that behave like paper (pages,
slides, thumbnails, menus, dialogs). One accent, cobalt, marks focus and selection; Word, Excel and PowerPoint keep
small identity marks in cobalt, pine and ember. Geist for the interface, Newsreader for titles and headings, Geist
Mono only for formulas and cell addresses, all self-hosted so they work offline. The ribbon drops group captions
(every control keeps a tooltip and an accessible name), which halves its height.

## Architecture

```
src/
  app/            Shell, routing (editors are lazy-loaded), error boundary and storage guard
  components/     Editor header, top bar with search, dialogs
  design-system/  Buttons, ribbon, portal menus, modals, colour and font pickers
  applications/   Route-level screens: home, drive, word, excel, powerpoint
  editors/        The editing engines
    word/           Tiptap editor, find/replace, page setup, DOCX/PDF export
    excel/          model, formula evaluator + function library, grid, charts, sheet operations
    powerpoint/     model, canvas, rich text, transitions, presenter, PPTX/PDF export
  storage/        Dexie (IndexedDB) database
  file-system/    File service shared by every app, starter templates
  services/       Importers, downloads, font catalogue and loader
  hooks/          useAutosave, useHistory (undo/redo), useFiles, useFileLoader
public/           Service worker, web manifest, icon
e2e/, e2e-prod/   Playwright journeys (dev server / production build)
docs/             Reference notes
```

Design notes:
- One `files` table holds every document. An app loads a file by id, owns its in-memory state while open, and writes
  back through `useAutosave`: debounced, retried with backoff, flushed on `pagehide`, and guarded by a "leave site?"
  prompt if a save is pending or has failed.
- The formula engine (`editors/excel/formula.ts`) is a tokenizer, recursive-descent parser and lazy evaluator.
  Dependencies are recorded as formulas read cells, so an edit invalidates only that cell's dependents. Long
  dependency chains are resolved iteratively, so they cannot overflow the call stack. The function library
  (`functions.ts`) is a registry, and a test fails if a registered function has no test case.
- Production builds ship a strict Content-Security-Policy (`script-src 'self'`; Google Fonts is the only third-party
  origin) and a service worker for the offline app shell. Documents never pass through the service worker.

## Privacy and data

Documents, settings and the recent-fonts list live in your browser (IndexedDB and localStorage). The only network
requests are to Google Fonts, for font previews you open or fonts you use. Clearing site data erases your documents;
export anything you want to keep from Drive. Offline, Google fonts are hidden and installed fonts still work.

## Known limitations

- Single user, single browser profile. No sync or sharing.
- **Word**: no comments, track changes, footnotes or headers/footers. DOCX export covers headings, text formatting,
  lists, tables, images, links and page setup, not every Word feature. PDF export is rendered by the browser and is
  not pixel-identical to the editor. Page count is approximate (a guide line marks each page height).
- **Excel**: no pivot tables, merged cells, data validation or macros. XLSX export writes values and formulas but
  not cell styling, charts or conditional formats (SheetJS community edition). The `xlsx` package is the
  unmaintained 0.18.5 release, so only import files you trust. Sheets are 100 × 26 by default (imports grow to fit).
- **PowerPoint**: transitions are kept in Office One and are not written to PPTX (the generator has no transition
  API). PPTX import is not supported. Fonts are referenced by name (and set as the file's heading and body theme
  fonts), so PowerPoint substitutes any Google Font not installed on the viewing machine. Theme colours are written
  as explicit colours on every object, not as a PowerPoint colour scheme, so "Change theme" in PowerPoint will not
  recolour the deck. Gradient and pattern backgrounds are exported as a background picture (not editable, and
  fixed to the 16:9 slide), and decoration such as the translucent circles is exported as ordinary shapes. PDF
  export rasterises each slide, so it matches the screen exactly.
- Google Fonts needs a network connection the first time a family is used.

## Roadmap

Cloud sync and accounts, real-time collaboration, comments and version history, PPTX import, richer chart options,
pivot tables, track changes, and an accessibility audit with screen-reader testing.
