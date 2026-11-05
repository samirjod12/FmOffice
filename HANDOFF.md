# FmOffice — HANDOFF (read this first, then continue the work)

Last updated: 2026-11-03 (session 1, started fresh — an earlier session's files were lost and could not be recovered)

## What the owner wants
- **FmOffice** = a fan-made, modern, friendly clone of Microsoft Office. Blue logo = `assets/fmoffice-logo.png`.
- **FmOffice Sheets** = the Microsoft Excel clone. Green logo = `assets/fmoffice-sheets-logo.png`.
- Same controls, commands, shortcuts, features and layout idea as Excel, but a modern, friendly UI.
- Install flow: the user runs a small **FmOffice Setup** (target: **max 50 MB**) which downloads the apps (Sheets, later Docs/Slides) from **GitHub Releases**. The owner compiles everything with **GitHub Actions** (no local build machine).
- The owner writes casual English; keep answers short and practical.

## Decisions made
- UI = plain HTML/CSS/JS, **no npm dependencies** (the build sandbox has no npm/CDN access). Works in a browser and inside Tauri.
- Desktop shell = **Tauri** (small binaries, ~10–20 MB), not Electron (80–100 MB).
- `.xlsx` read/write is hand-written in `app/js/xlsx.js` (ZIP store + `DecompressionStream` for reading).
- Formula engine is in `app/js/engine.js` (UMD, runs in Node for tests).

## Status
See the checklist at the bottom. Update it every time you finish something.

## File map
```
app/index.html          entry page of FmOffice Sheets
app/css/style.css       theme (light/dark), ribbon, grid chrome
app/js/engine.js        workbook model, parser, 150+ functions, number formats, ref shifting, row/col insert/delete
app/js/xlsx.js          .xlsx import/export (own ZIP writer; reads with DecompressionStream), CSV
app/js/core.js          state, canvas grid renderer, selection, editing, autocomplete, clipboard, fill, undo (JSON snapshots)
app/js/charts.js        SVG charts (col/stack/bar/line/area/pie/scatter)
app/js/ribbon.js        ribbon, all commands (App.CMD), dialogs, menus, sheet tabs, file I/O, command search, sample data
tests/engine.test.js    node tests:  node tests/engine.test.js   (58 pass)
src-tauri/              Tauri v2 desktop wrapper for Sheets (dialog+fs plugins for native Open/Save)
setup/                  FmOffice Setup launcher (Tauri v2): downloads FmOffice-Sheets-windows-x64.zip from GitHub Releases
.github/workflows/      build.yml: tests, Windows (Sheets zip + Setup exe), Linux/macOS bundles, release on tag v*
assets/                 logos + 1024px transparent icons
```

## Verified vs NOT verified
- VERIFIED (headless Chromium + node): engine tests, typing/formulas, undo/redo, paste shifting, fill down, series fill, sort, chart, .xlsx round trip, dark mode, no console errors.
- NOT VERIFIED: the Rust/Tauri code (`src-tauri`, `setup/src-tauri`) has never been compiled — the build sandbox had no crates.io access. First GitHub Actions run may need small fixes (read the log). The installer was never run on Windows.
- Excel files from real Excel: imported with styles/merges/widths, but charts, images, data validation, defined names, conditional formats and table styles inside the file are dropped on import.

## Known gaps / backlog (priority order)
1. Run the GitHub workflow, fix any Rust/Tauri compile errors; test Setup on Windows.
2. Named ranges, data validation (dropdown lists), cell comments, hyperlinks, images.
3. Import/export charts + conditional formats in .xlsx; shared formulas on import.
4. Pivot tables, array/spill formulas (FILTER, SORT, SEQUENCE), LET/LAMBDA, more financial/stat functions (IRR, NPER, RATE…).
5. Copy as HTML table to system clipboard (formats to Excel), paste HTML from Excel.
6. Auto row height on wrap when editing; drag-to-move selection; multi-range (Ctrl+click) selection.
7. File associations (.xlsx double-click opens in app) in Tauri; code signing for Windows SmartScreen.
8. FmOffice Docs and Slides (setup UI already lists them as "Coming soon").
9. Undo uses whole-workbook JSON snapshots: fine for normal sheets, slow for >100k cells — switch to diff-based.

## Checklist
- [x] Project folder + logos
- [x] Formula engine + tests
- [x] Grid UI, selection, editing, formula bar
- [x] Ribbon (Home / Insert / Formulas / Data / View), formatting, number formats
- [x] Undo/redo, copy/paste, fill handle, sort, find/replace
- [x] Sheet tabs (add/rename/delete/duplicate/color), freeze panes, row/col insert/delete
- [x] .xlsx / .csv open + save
- [x] Charts, conditional formatting, filters
- [x] Tauri wrapper, GitHub Actions, Setup launcher (uncompiled)
- [ ] First real build on GitHub + Windows test
- [ ] Data validation, named ranges, pivot tables
- [ ] FmOffice Docs / Slides


## Update (latest session)
- **Classic skin**: `app/css/classic.css` makes Sheets look like the Excel-2013 layout the user supplied as reference (FILE tab, HOME/INSERT/PAGE LAYOUT/FORMULAS/DATA/REVIEW/VIEW, grouped ribbon with launcher corners, green READY status bar, sheet-tab row). Original icons/branding only; never copy Microsoft logos/template art.
- **Backstage**: `app/js/backstage.js` + `app/css/backstage.css` — FILE opens Home/New/Open/Save/Print/Account/Feedback/Options; 6 original templates; Recent list (names only, localStorage `fmsheets.recent`). Shown automatically on first launch.
- New commands: Format Painter, font grow/shrink, outdent, Page Layout (margins/orientation/size used by print), Review tab (find/replace/notes/undo).
- Docs (`docs.html`) loads and renders; docx round-trip not yet tested in browser.
- Tauri `open_url` command added (http/https/mailto only).
- Engine tests: 82 pass.
- **Still TODO**: FmOffice Slides; single `fmoffice` exe + Setup updates for Docs/Slides; text orientation button; Page Break/Layout view buttons are placeholders; Rust code has never been compiled; Setup never run on Windows; pivot is static.
