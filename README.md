# Office One

Word, Excel and PowerPoint in your browser. It is free, needs no account, and nothing is uploaded:
your files stay on your own computer.

Not affiliated with Microsoft.

## What you can do

- **Word**: write and format documents. Import DOCX, export DOCX or PDF.
- **Excel**: spreadsheets with about 200 formulas, charts and sorting. Import and export CSV or XLSX.
- **PowerPoint**: make slides with 22 ready-made themes (light, dark and colourful), shapes, pictures and
  transitions. Present full screen. Export PPTX or PDF.
- **Drive**: one place for all your files, with folders, search, favourites and trash.

## Run it on your computer

You need [Node.js](https://nodejs.org) version 22 or newer.

**Step 1.** Download the project and open a terminal in its folder.

```bash
git clone https://github.com/arnavvarshney350-lab/office-one.git
cd office-one
```

**Step 2.** Install what it needs (only the first time).

```bash
npm install
```

**Step 3.** Start it.

```bash
npm run dev
```

**Step 4.** Open <http://localhost:5173> in your browser. Done!

Works in current Chrome, Edge, Safari and Firefox.

## Use it as a Mac app (optional)

**Step 1.** Do steps 1 and 2 above.

**Step 2.** Build the app.

```bash
npm run desktop:build
```

**Step 3.** Open the `release` folder and double-click `Office One-1.0.0-arm64.dmg`.

**Step 4.** Drag **Office One** into **Applications**.

**Step 5.** The first time, right-click the app and choose **Open**. (macOS asks because the app is not
signed with a paid Apple account.)

For an Intel Mac, use `npm run desktop:build -- x64` in step 2.

## Make a presentation with a theme

1. Open **PowerPoint** and click any theme to start a new deck.
2. To change the look later, open the **Design** tab and click another theme.
3. Use the **Light / Dark / Colourful** buttons to filter. Your text and pictures are kept.

## Save your work

- Files are saved automatically in your browser.
- Clearing your browser's site data **deletes** them. To keep a file, export it from Drive.
- The Mac app has its own storage, separate from the browser version.

## Useful commands

| Command | What it does |
| --- | --- |
| `npm run dev` | Start the app while you work on it |
| `npm run build` | Make the final version in `dist/` |
| `npm run check` | Check types, code style and unit tests |
| `npm run e2e` | Run the browser tests (needs Google Chrome installed) |
| `npm run verify` | Run every check and test |

## Good to know

- One person, one browser. There is no sync or sharing between devices.
- Fonts come from Google Fonts, so you need internet the first time you use a font. Offline, only fonts
  installed on your computer are offered.
- **Word**: no comments, track changes, footnotes, or headers and footers.
- **Excel**: no pivot tables, merged cells or macros. XLSX export keeps values and formulas but not cell
  colours or charts.
- **PowerPoint**:
  - Slide transitions are not saved in PPTX files.
  - Opening a PPTX file in Office One is not supported.
  - Theme fonts are saved by name, so PowerPoint swaps in another font if it isn't installed.
  - Theme colours are saved as plain colours, so PowerPoint's own "Change theme" will not recolour the deck.
  - Gradient and pattern backgrounds are saved as a picture in PPTX files.

## For developers

Built with React, TypeScript, Vite and Tailwind. The code lives in `src/`:

| Folder | What is inside |
| --- | --- |
| `src/applications/` | The screens: home, drive, word, excel, powerpoint |
| `src/editors/` | The editing engines (formulas, slides, rich text, exports) |
| `src/design-system/` | Buttons, menus, ribbon, pickers |
| `src/services/` | Importers, downloads, fonts |
| `src/storage/` | The browser database |
| `e2e/`, `e2e-prod/` | Browser tests |

Slide themes are listed in `src/editors/powerpoint/themes.ts`. To add one, copy an entry and change its
colours and fonts. The unit tests check that its text is readable.

More notes are in `docs/`.
