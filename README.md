# FmOffice

Free, fan-made, modern take on an office suite. **FmOffice Sheets** is a spreadsheet app with the familiar Excel-style layout, shortcuts and formulas.
*Not affiliated with or endorsed by Microsoft.*

| App | Logo | Status |
|---|---|---|
| FmOffice Sheets | green (`assets/fmoffice-sheets-logo.png`) | working |
| FmOffice Setup | blue (`assets/fmoffice-logo.png`) | working (needs first GitHub build) |

## Try it now (no install)
Open `app/index.html` in Chrome/Edge, or run `npx serve app`.

## Build & release with GitHub (no local tools needed)
1. Create a GitHub repo and upload this whole folder (keep the folder structure).
2. In **Actions → Build FmOffice → Run workflow** (checks everything compiles), or publish a release by tagging:
   `git tag v0.1.0 && git push origin v0.1.0`
3. The workflow attaches to the Release:
   - `FmOffice-Setup.exe` – the small installer (a few MB). Users run this.
   - `FmOffice-Sheets-windows-x64.zip` – the app it downloads from GitHub.
4. Setup downloads `releases/latest/download/FmOffice-Sheets-windows-x64.zip` from the same repo, installs to `%LOCALAPPDATA%\FmOffice`, and adds Start-menu + desktop shortcuts.

Windows needs the WebView2 runtime (already included in Windows 10/11).

## Features (Sheets)
Grid up to 1,048,576 × 16,384 · 150+ functions (SUM, IF, VLOOKUP, XLOOKUP, SUMIFS, TEXT, DATE…) · formula bar, name box, autocomplete, colored references, F4 · undo/redo · copy/cut/paste (values, formats, transpose) · fill handle & series · sort/filter/remove duplicates · find & replace · number formats · fonts, colors, borders, merge, wrap · conditional formatting (rules, data bars, color scales) · charts (column, stacked, bar, line, area, pie, scatter) · multiple sheets · freeze panes · row/column insert/delete/hide/resize · `.xlsx` and `.csv` open/save · dark mode · autosave recovery · “Tell me what you want to do” (Alt+Q).

## Developer
```
node tests/engine.test.js     # formula-engine tests
```
See `HANDOFF.md` for architecture, status and backlog.
