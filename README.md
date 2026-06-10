# BoringSheets

[![License: MIT](https://img.shields.io/badge/License-MIT-0ae448.svg)](LICENSE)
[![Latest release](https://img.shields.io/github/v/release/Bedduworkable/boringsheets?include_prereleases&color=0ae448)](https://github.com/Bedduworkable/boringsheets/releases/latest)
[![Downloads](https://img.shields.io/github/downloads/Bedduworkable/boringsheets/total?color=0ae448)](https://github.com/Bedduworkable/boringsheets/releases)
![Platform: macOS](https://img.shields.io/badge/platform-macOS-lightgrey.svg)

**A fast, native spreadsheet for macOS — free for life, and it never connects to the internet.**

BoringSheets opens and saves real Excel `.xlsx` files, with 175+ functions,
charts, conditional formatting and filters. The spreadsheet engine — grid
renderer, formula parser, calculation/dependency engine, and `.xlsx`
reader/writer — is **written from scratch in TypeScript**, not wrapped around an
existing library. Built on Electron. Mac-first; Windows planned.

- 🆓 **Free for life** — no subscription, no account, no telemetry.
- 🔌 **100% offline** — zero network requests; your data never leaves your device.
- 📄 **Real `.xlsx`** — formulas, formatting, charts, merges, conditional formatting, named ranges.
- ⚙️ **From-scratch engine** — tokenizer → parser → AST → evaluator → dependency graph with incremental recalc and cycle detection.

> **Status:** early beta, Mac-first. The build is currently unsigned — on first
> launch, right-click the app and choose **Open**.

License: [MIT](LICENSE).

## Download

### ⬇ [Download for macOS (Apple Silicon)](https://github.com/Bedduworkable/boringsheets/releases/latest)

> Apple Silicon (M-series) only for now — an Intel / universal build is coming.
> The beta is **unsigned**, so on first launch **right-click the app → Open**, then click **Open** in the dialog (a one-time step). See [all releases ›](https://github.com/Bedduworkable/boringsheets/releases).

## Run it

```bash
npm install        # one time
npm run dev        # launch the app with live reload (Mac/Windows)
```

For a production-style run:

```bash
npm run build      # bundle renderer + Electron shell
npm start          # launch the built app
```

## Features

- **Grid**: canvas-rendered with virtual scrolling, keyboard navigation
  (arrows, Tab, Enter, Page Up/Down, Home), drag-select ranges, click column/row
  headers to select, in-cell editing, copy / cut / paste (TSV — interops with
  real Excel), and undo / redo (full command history).
- **Structure**: insert / delete rows & columns **with automatic formula
  reference rewriting** (references shift, deleted targets become `#REF!`),
  drag-to-resize columns/rows (double-click a column border to auto-fit),
  **merge / unmerge** cells, and a right-click **context menu**.
- **Fill handle**: drag the selection corner to copy values, extend numeric
  series, or offset formulas (relative refs move, `$`-anchored refs stay).
- **Formulas**: full expression parser with Excel operator precedence and
  **175+ built-in functions** — math/trig, statistical (MEDIAN, STDEV,
  PERCENTILE…), the IFS family (SUMIFS/COUNTIFS/AVERAGEIFS/MAXIFS/MINIFS),
  text, logical (IF, SWITCH, IFERROR…), lookup (VLOOKUP, XLOOKUP, INDEX/MATCH,
  CHOOSE…), date/time, information (ISNUMBER, ISBLANK…), and financial
  (PMT, FV, PV, NPV, IPMT, PPMT…). Incremental recalculation via a dependency
  graph, with circular-reference detection.
- **Cross-sheet references**: `Sheet2!A1` and `'My Sheet'!A1:B2`, with a
  workbook-wide dependency graph (edits on one sheet recalc dependents on
  others) and automatic reference rewriting when a sheet is renamed.
- **Freeze panes**: freeze rows above / columns left of the selection so they
  stay put while the rest scrolls.
- **Sort & filter**: sort a selection ascending/descending by the active column
  (relative formulas re-anchor as rows move), and a Google-Sheets-style filter —
  a funnel in each column header with sort and a searchable value checklist;
  multiple columns combine with AND.
- **Formatting**: font family & size, bold / italic / underline / strikethrough,
  text & fill color, alignment, text wrap (rows auto-fit), borders, and
  Excel-style number formats (currency, percent, comma, dates).
- **Conditional formatting**: cell-value rules (>, <, between, equal, text
  contains), duplicate/top/bottom, color scales, and data bars.
- **Charts**: column, bar, line, area, pie, and scatter — floating, draggable,
  resizable, and live-updating as their source data changes.
- **Data validation**: list (dropdown) and numeric/text-length rules, enforced
  on entry, with an in-cell dropdown picker.
- **Find & Replace** (Cmd/Ctrl+F) with match-case and replace-all.
- **Files**: open and save real `.xlsx` workbooks (values, formulas, shared
  strings, and styles) through native file dialogs; **CSV import/export** and
  **Export to PDF**. Multiple sheets with tabs.

## Architecture

```
electron/        Desktop shell (main + preload, native file dialogs, menus)
src/
  model/         Cell, Sheet, Workbook data model
  engine/        tokenizer → parser → evaluator → functions → calc (recalc) + format
  grid/          Canvas grid renderer + input handling
  ui/            Toolbar
  io/            .xlsx read/write (SpreadsheetML over zip)
  app.ts         Controller wiring engine + grid + UI + files
test/            Headless engine smoke tests
```

## Testing the engine headlessly

```bash
npx esbuild test/engine.test.ts --bundle --platform=node --format=cjs --outfile=test/.engine.test.cjs && node test/.engine.test.cjs
```

## Roadmap toward fuller Excel parity

Remaining: pivot tables, the rest of Excel's ~480 functions, persisting **charts**
into the `.xlsx` file (charts are in-app only for now — formulas, styles, fonts,
merges, conditional formatting, data validation, freeze panes, named ranges and
notes already round-trip), and richer chart styling.
