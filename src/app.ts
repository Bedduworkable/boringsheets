// Application controller: owns the workbook, calc engine and grid, and wires up
// the formula bar, toolbar, sheet tabs, undo/redo, clipboard, structural edits
// (insert/delete rows & columns), merges, fill, and file I/O.

import { Workbook } from "./model/workbook.js";
import { Sheet, DataValidation } from "./model/sheet.js";
import { CalcEngine } from "./engine/calc.js";
import { Grid, GridHost, SelRange } from "./grid/grid.js";
import { CellFormat, CellValue, CellError, BorderSet } from "./model/types.js";
import { ConditionalEngine, ConditionalRule, CondVisual } from "./engine/conditional.js";
import { parseCsv, toCsv } from "./io/csv.js";
import { drawChart, ChartSpec, ChartType } from "./charts/render.js";
import { SheetChart, DEFAULT_ROW_HEIGHT, DEFAULT_COL_WIDTH } from "./model/sheet.js";
import { wrapLines } from "./engine/textwrap.js";
import { showConditionalDialog } from "./ui/conditional.js";
import { showValidationDialog } from "./ui/datavalidation.js";
import { showChartDialog } from "./ui/chartdialog.js";
import { showPrompt } from "./ui/prompt.js";
import { FormulaAutocomplete } from "./ui/autocomplete.js";
import { formatA1, parseA1, parseKey } from "./engine/references.js";
import { rewriteFormulaRefs, offsetFormula, renameSheetRefs, RefTransform } from "./engine/printer.js";
import { formatValue } from "./engine/format.js";
import { readXlsx, writeXlsx, sheetToTsv } from "./io/xlsx.js";
import { buildToolbar } from "./ui/toolbar.js";
import { showContextMenu } from "./ui/contextmenu.js";
import { showFindReplace } from "./ui/findreplace.js";
import { showColumnFilter } from "./ui/filter.js";

interface Command {
  undo(): void;
  redo(): void;
}

interface CellDelta {
  sheetIndex: number;
  row: number;
  col: number;
  oldRaw: string;
  newRaw: string;
  oldFormat?: CellFormat;
  newFormat?: CellFormat;
}

export class App implements GridHost {
  wb = new Workbook();
  engine: CalcEngine;
  grid: Grid;

  private nameBox: HTMLInputElement;
  private formulaInput: HTMLInputElement;
  private tabsEl: HTMLElement;
  private toolbarEl: HTMLElement;
  private measureCtx: CanvasRenderingContext2D;
  private chartLayer: HTMLElement;
  private statusBar: HTMLElement;
  private chartSeq = 0;

  // Internal clipboard for rich (formula + format) copy/paste within the app.
  // `tsv` is what we put on the system clipboard, used to detect our own paste.
  private internalClip: { rows: { raw: string; format?: CellFormat }[][]; r0: number; c0: number; cut: boolean } | null = null;
  private clipTsv = "";

  private undoStack: Command[] = [];
  private redoStack: Command[] = [];
  private pendingDeltas: CellDelta[] | null = null;

  // Cached conditional-formatting engine, rebuilt when rules or sheet change.
  private condEngine: ConditionalEngine | null = null;
  private condEngineKey = "";
  private condVersion = 0;

  constructor(els: {
    canvas: HTMLCanvasElement;
    container: HTMLElement;
    editor: HTMLInputElement;
    nameBox: HTMLInputElement;
    formulaInput: HTMLInputElement;
    tabs: HTMLElement;
    toolbar: HTMLElement;
  }) {
    this.nameBox = els.nameBox;
    this.formulaInput = els.formulaInput;
    this.tabsEl = els.tabs;
    this.toolbarEl = els.toolbar;
    this.measureCtx = document.createElement("canvas").getContext("2d")!;
    this.chartLayer = document.getElementById("chart-layer") || document.createElement("div");
    this.statusBar = document.getElementById("status-bar") || document.createElement("div");

    this.engine = new CalcEngine(this.wb);
    this.grid = new Grid(els.canvas, els.container, els.editor, this.wb.active, this);

    buildToolbar(this.toolbarEl, this);
    this.wireFormulaBar();
    this.wireMenus();
    this.wireClipboard();
    this.renderTabs();
    this.onSelectionChange();
    document.title = this.titleText();
  }

  // ===== GridHost: edits =====

  // Interactive edit: enforces data validation, then applies.
  commit(row: number, col: number, raw: string) {
    if (raw !== "" && !this.checkValidation(row, col, raw)) return;
    this.commitRaw(row, col, raw);
  }

  // Apply an edit without validation (used by fill/paste/programmatic paths).
  private commitRaw(row: number, col: number, raw: string) {
    const sheet = this.wb.active;
    const cur = sheet.getCell(row, col);
    const oldRaw = cur?.raw ?? "";
    if (oldRaw === raw) return;
    const fmt = cur?.format ? { ...cur.format } : undefined;
    this.pushDelta({
      sheetIndex: this.wb.activeIndex,
      row,
      col,
      oldRaw,
      newRaw: raw,
      oldFormat: fmt,
      newFormat: fmt,
    });
    this.engine.setCellRaw(row, col, raw);
    if (fmt) sheet.ensureCell(row, col).format = fmt;
    this.recomputeRowHeight(sheet, row);
    this.markDirty();
    this.grid.render();
    if (sheet.charts.length) this.renderCharts();
  }

  // ===== Data validation (GridHost + enforcement) =====

  listValidation(row: number, col: number): string[] | null {
    const v = this.wb.active.validationAt(row, col);
    return v && v.type === "list" && v.source && v.source.length ? v.source : null;
  }

  onDropdown(row: number, col: number, clientX: number, clientY: number) {
    const vals = this.listValidation(row, col);
    if (!vals) return;
    showContextMenu(
      clientX,
      clientY,
      vals.map((val) => ({
        label: val === "" ? "(blank)" : val,
        action: () => {
          this.grid.setActive(row, col);
          this.commitRaw(row, col, val);
        },
      }))
    );
  }

  // Returns true if `raw` satisfies the cell's validation (alerts on failure).
  private checkValidation(row: number, col: number, raw: string): boolean {
    const v = this.wb.active.validationAt(row, col);
    if (!v) return true;
    const fail = (msg: string) => {
      alert(v.errorMessage || msg);
      return false;
    };
    if (v.type === "list") {
      if (v.source && v.source.includes(raw)) return true;
      return fail(`Value must be one of: ${(v.source || []).join(", ")}`);
    }
    if (v.type === "number") {
      const n = Number(raw);
      if (Number.isNaN(n)) return fail("Value must be a number");
      return this.checkNumberOp(v, n) ? true : fail("Number is out of the allowed range");
    }
    if (v.type === "textLength") {
      return this.checkNumberOp(v, raw.length) ? true : fail("Text length is not allowed");
    }
    return true;
  }

  private checkNumberOp(v: DataValidation, n: number): boolean {
    const { operator, min, max } = v;
    switch (operator) {
      case "between": return min !== undefined && max !== undefined && n >= min && n <= max;
      case "notBetween": return !(min !== undefined && max !== undefined && n >= min && n <= max);
      case "gt": return min !== undefined && n > min;
      case "lt": return min !== undefined && n < min;
      case "gte": return min !== undefined && n >= min;
      case "lte": return min !== undefined && n <= min;
      case "eq": return min !== undefined && n === min;
      case "ne": return min !== undefined && n !== min;
      default: return true;
    }
  }

  // Add a list validation over the current selection.
  addListValidation(values: string[]) {
    const sel = this.grid.selRange();
    this.structural((sheet) => {
      sheet.validations.push({ range: { ...sel }, type: "list", source: values, allowBlank: true });
    });
  }

  addNumberValidation(operator: DataValidation["operator"], min: number, max?: number) {
    const sel = this.grid.selRange();
    this.structural((sheet) => {
      sheet.validations.push({ range: { ...sel }, type: "number", operator, min, max });
    });
  }

  addTextLengthValidation(operator: DataValidation["operator"], min: number, max?: number) {
    const sel = this.grid.selRange();
    this.structural((sheet) => {
      sheet.validations.push({ range: { ...sel }, type: "textLength", operator, min, max });
    });
  }

  clearValidation() {
    const sel = this.grid.selRange();
    this.structural((sheet) => {
      sheet.validations = sheet.validations.filter(
        (v) => !(v.range.r0 >= sel.r0 && v.range.r1 <= sel.r1 && v.range.c0 >= sel.c0 && v.range.c1 <= sel.c1)
      );
    });
  }

  // ===== Conditional formatting (GridHost + management) =====

  conditional(row: number, col: number): CondVisual | null {
    const sheet = this.wb.active;
    if (sheet.conditionalRules.length === 0) return null;
    const key = `${sheet.id}:${this.condVersion}`;
    if (key !== this.condEngineKey) {
      this.condEngine = new ConditionalEngine(sheet.conditionalRules);
      this.condEngineKey = key;
    }
    const value = sheet.getCell(row, col)?.value ?? null;
    return this.condEngine!.resolve(row, col, value, (r, c) => sheet.getCell(r, c)?.value ?? null);
  }

  addConditionalRule(rule: ConditionalRule) {
    this.structural((sheet) => sheet.conditionalRules.push(rule));
    this.condVersion++;
  }

  clearConditional() {
    const sel = this.grid.selRange();
    this.structural((sheet) => {
      sheet.conditionalRules = sheet.conditionalRules.filter(
        (r) => !(r.range.r0 >= sel.r0 && r.range.r1 <= sel.r1 && r.range.c0 >= sel.c0 && r.range.c1 <= sel.c1)
      );
    });
    this.condVersion++;
  }

  selectionRect() {
    return this.grid.selRange();
  }

  // ===== Borders =====

  applyBorder(type: "all" | "outer" | "top" | "bottom" | "left" | "right" | "none", color = "#000000") {
    const sel = this.grid.selRange();
    for (let r = sel.r0; r <= sel.r1; r++) {
      for (let c = sel.c0; c <= sel.c1; c++) {
        const cell = this.wb.active.ensureCell(r, c);
        const oldFormat = cell.format ? { ...cell.format } : undefined;
        let border: BorderSet | undefined;
        if (type === "none") {
          border = undefined;
        } else {
          const b: BorderSet = { ...(cell.format?.border || {}), color };
          if (type === "all") { b.top = b.bottom = b.left = b.right = true; }
          else if (type === "outer") {
            if (r === sel.r0) b.top = true;
            if (r === sel.r1) b.bottom = true;
            if (c === sel.c0) b.left = true;
            if (c === sel.c1) b.right = true;
          } else if (type === "top" && r === sel.r0) b.top = true;
          else if (type === "bottom" && r === sel.r1) b.bottom = true;
          else if (type === "left" && c === sel.c0) b.left = true;
          else if (type === "right" && c === sel.c1) b.right = true;
          border = b.top || b.bottom || b.left || b.right ? b : undefined;
        }
        const next: CellFormat = { ...(cell.format || {}) };
        if (border) next.border = border;
        else delete next.border;
        cell.format = Object.keys(next).length ? next : undefined;
        this.wb.active.deleteCellIfEmpty(r, c);
        this.pushDelta({
          sheetIndex: this.wb.activeIndex,
          row: r,
          col: c,
          oldRaw: cell.raw,
          newRaw: cell.raw,
          oldFormat,
          newFormat: cell.format ? { ...cell.format } : undefined,
        });
      }
    }
    this.markDirty();
    this.grid.render();
  }

  // ===== Notes =====

  async editNote(row = this.grid.active.row, col = this.grid.active.col) {
    const cell = this.wb.active.getCell(row, col);
    const note = await showPrompt("Cell note:", cell?.note ?? "");
    if (note === null) return;
    const c = this.wb.active.ensureCell(row, col);
    c.note = note.trim() === "" ? undefined : note;
    this.wb.active.deleteCellIfEmpty(row, col);
    this.markDirty();
    this.grid.render();
  }

  // ===== Hide / unhide =====

  hideRows() {
    const sel = this.grid.selRange();
    this.structural((sheet) => {
      for (let r = sel.r0; r <= sel.r1; r++) sheet.hiddenRows.add(r);
    });
  }
  hideCols() {
    const sel = this.grid.selRange();
    this.structural((sheet) => {
      for (let c = sel.c0; c <= sel.c1; c++) sheet.hiddenCols.add(c);
    });
  }
  unhide() {
    const sel = this.grid.selRange();
    this.structural((sheet) => {
      for (let r = sel.r0; r <= sel.r1; r++) sheet.hiddenRows.delete(r);
      for (let c = sel.c0; c <= sel.c1; c++) sheet.hiddenCols.delete(c);
    });
  }

  // ===== Format painter =====

  private painterFormat: CellFormat | null = null;

  startFormatPainter() {
    const f = this.activeFormat();
    this.painterFormat = f ? { ...f } : {};
    document.body.style.cursor = "copy";
  }

  private applyPainter() {
    if (!this.painterFormat) return;
    const fmt = this.painterFormat;
    this.painterFormat = null;
    document.body.style.cursor = "";
    const sel = this.grid.selRange();
    for (let r = sel.r0; r <= sel.r1; r++) {
      for (let c = sel.c0; c <= sel.c1; c++) {
        const cell = this.wb.active.ensureCell(r, c);
        const oldFormat = cell.format ? { ...cell.format } : undefined;
        cell.format = Object.keys(fmt).length ? { ...fmt } : undefined;
        this.wb.active.deleteCellIfEmpty(r, c);
        this.pushDelta({
          sheetIndex: this.wb.activeIndex,
          row: r,
          col: c,
          oldRaw: cell.raw,
          newRaw: cell.raw,
          oldFormat,
          newFormat: cell.format ? { ...cell.format } : undefined,
        });
      }
    }
    this.markDirty();
    this.grid.render();
  }

  // ===== Named ranges =====

  defineName(name: string) {
    const sel = this.grid.selRange();
    const sheet = this.wb.active;
    const existing = this.wb.names.find((n) => n.name.toLowerCase() === name.toLowerCase());
    const nr = { name, sheetId: sheet.id, r0: sel.r0, c0: sel.c0, r1: sel.r1, c1: sel.c1 };
    if (existing) Object.assign(existing, nr);
    else this.wb.names.push(nr);
    this.engine.rebuild();
    this.grid.render();
    this.markDirty();
  }

  // Navigate to a named range if it exists; returns true if it did.
  private gotoName(name: string): boolean {
    const nr = this.wb.names.find((n) => n.name.toLowerCase() === name.toLowerCase());
    if (!nr) return false;
    const idx = this.wb.sheets.findIndex((s) => s.id === nr.sheetId);
    if (idx < 0) return false;
    this.activateSheet(idx);
    this.grid.setActive(nr.r1, nr.c1);
    this.grid.setActive(nr.r0, nr.c0, true);
    return true;
  }

  // ===== Zoom =====

  onZoomChange(_zoom: number) {
    this.refreshToolbarState();
  }
  getZoomPct(): number {
    return Math.round(this.grid.getZoom() * 100);
  }
  zoomIn() {
    this.grid.zoomBy(0.1);
  }
  zoomOut() {
    this.grid.zoomBy(-0.1);
  }
  zoomReset() {
    this.grid.setZoom(1);
  }

  openConditionalDialog() {
    showConditionalDialog(this);
  }
  openValidationDialog() {
    showValidationDialog(this);
  }
  openChartDialog() {
    showChartDialog(this);
  }

  // ===== Charts =====

  addChart(type: ChartType) {
    const sel = this.grid.selRange();
    const chart: SheetChart = {
      id: `chart-${++this.chartSeq}`,
      spec: { type, categories: [], series: [] },
      dataRange: { ...sel },
      byRows: false,
      x: 140,
      y: 90,
      w: 440,
      h: 300,
    };
    this.structural((sheet) => sheet.charts.push(chart));
    this.renderCharts();
  }

  removeChart(id: string) {
    this.structural((sheet) => {
      sheet.charts = sheet.charts.filter((c) => c.id !== id);
    });
    this.renderCharts();
  }

  // Extract a ChartSpec from a chart's data range, using the convention that the
  // first row holds series names and the first column holds category labels.
  private buildChartSpec(chart: SheetChart): ChartSpec {
    const sheet = this.wb.active;
    const { r0, c0, r1, c1 } = chart.dataRange;
    const numCols = c1 - c0 + 1;
    const numRows = r1 - r0 + 1;
    const text = (r: number, c: number) => {
      const cell = sheet.getCell(r, c);
      return cell ? formatValue(cell.value, cell.format?.numFmt) : "";
    };
    const num = (r: number, c: number) => {
      const v = sheet.getCell(r, c)?.value;
      return typeof v === "number" ? v : NaN;
    };
    const hasHeaderRow = numRows > 1;
    const hasCategoryCol = numCols > 1;
    const dataR0 = hasHeaderRow ? r0 + 1 : r0;
    const dataC0 = hasCategoryCol ? c0 + 1 : c0;
    const categories: string[] = [];
    for (let r = dataR0; r <= r1; r++) categories.push(hasCategoryCol ? text(r, c0) : String(r - dataR0 + 1));
    const series = [];
    for (let c = dataC0; c <= c1; c++) {
      const name = hasHeaderRow ? text(r0, c) : `Series ${c - dataC0 + 1}`;
      const values: number[] = [];
      for (let r = dataR0; r <= r1; r++) values.push(num(r, c));
      series.push({ name, values });
    }
    return { type: chart.spec.type, title: chart.spec.title, categories, series };
  }

  // Rebuild the floating chart overlay for the active sheet.
  renderCharts() {
    this.chartLayer.innerHTML = "";
    const dpr = window.devicePixelRatio || 1;
    for (const chart of this.wb.active.charts) {
      const box = document.createElement("div");
      box.className = "chart-box";
      box.style.left = `${chart.x}px`;
      box.style.top = `${chart.y}px`;
      box.style.width = `${chart.w}px`;
      box.style.height = `${chart.h}px`;

      const header = document.createElement("div");
      header.className = "chart-header";
      const titleSel = document.createElement("select");
      for (const t of ["column", "bar", "line", "area", "pie", "scatter"]) {
        const o = document.createElement("option");
        o.value = t;
        o.textContent = t;
        if (t === chart.spec.type) o.selected = true;
        titleSel.appendChild(o);
      }
      titleSel.className = "chart-type";
      titleSel.addEventListener("change", () => {
        chart.spec.type = titleSel.value as ChartType;
        this.markDirty();
        this.renderCharts();
      });
      header.appendChild(titleSel);
      const close = document.createElement("span");
      close.className = "chart-close";
      close.textContent = "✕";
      close.addEventListener("click", () => this.removeChart(chart.id));
      header.appendChild(close);
      box.appendChild(header);

      const canvas = document.createElement("canvas");
      canvas.className = "chart-canvas";
      const cw = chart.w;
      const ch = chart.h - 26;
      canvas.style.width = `${cw}px`;
      canvas.style.height = `${ch}px`;
      canvas.width = Math.floor(cw * dpr);
      canvas.height = Math.floor(ch * dpr);
      box.appendChild(canvas);

      const resize = document.createElement("div");
      resize.className = "chart-resize";
      box.appendChild(resize);

      this.chartLayer.appendChild(box);

      const ctx = canvas.getContext("2d")!;
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
      try {
        drawChart(ctx, this.buildChartSpec(chart), cw, ch);
      } catch {
        /* never let a chart break the app */
      }

      this.wireChartDrag(box, header, resize, chart);
    }
  }

  private wireChartDrag(box: HTMLElement, header: HTMLElement, resize: HTMLElement, chart: SheetChart) {
    header.addEventListener("mousedown", (e) => {
      if ((e.target as HTMLElement).tagName === "SELECT" || (e.target as HTMLElement).className === "chart-close") return;
      e.preventDefault();
      const sx = e.clientX;
      const sy = e.clientY;
      const ox = chart.x;
      const oy = chart.y;
      const move = (ev: MouseEvent) => {
        chart.x = Math.max(0, ox + ev.clientX - sx);
        chart.y = Math.max(0, oy + ev.clientY - sy);
        box.style.left = `${chart.x}px`;
        box.style.top = `${chart.y}px`;
      };
      const up = () => {
        window.removeEventListener("mousemove", move);
        window.removeEventListener("mouseup", up);
        this.markDirty();
      };
      window.addEventListener("mousemove", move);
      window.addEventListener("mouseup", up);
    });

    resize.addEventListener("mousedown", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const sx = e.clientX;
      const sy = e.clientY;
      const ow = chart.w;
      const oh = chart.h;
      const move = (ev: MouseEvent) => {
        chart.w = Math.max(220, ow + ev.clientX - sx);
        chart.h = Math.max(160, oh + ev.clientY - sy);
        box.style.width = `${chart.w}px`;
        box.style.height = `${chart.h}px`;
      };
      const up = () => {
        window.removeEventListener("mousemove", move);
        window.removeEventListener("mouseup", up);
        this.markDirty();
        this.renderCharts();
      };
      window.addEventListener("mousemove", move);
      window.addEventListener("mouseup", up);
    });
  }

  onSelectionChange() {
    const { row, col } = this.grid.active;
    this.nameBox.value = formatA1(row, col);
    this.formulaInput.value = this.wb.active.getRaw(row, col);
    this.refreshToolbarState();
    this.updateStatusBar();
    // Format painter applies to the selection once it's finalized (not mid-drag).
    if (this.painterFormat && !this.grid.isSelecting()) this.applyPainter();
  }

  // Live aggregates of the current selection (like Excel's status bar).
  private updateStatusBar() {
    const sel = this.grid.selRange();
    const nums: number[] = [];
    let count = 0;
    for (let r = sel.r0; r <= sel.r1; r++) {
      for (let c = sel.c0; c <= sel.c1; c++) {
        const v = this.wb.active.getCell(r, c)?.value;
        if (v === null || v === undefined || v === "") continue;
        count++;
        if (typeof v === "number") nums.push(v);
      }
    }
    const fmt = (n: number) => {
      const r = Math.round(n * 1e9) / 1e9;
      return Number.isInteger(r) ? String(r) : r.toLocaleString(undefined, { maximumFractionDigits: 6 });
    };
    const parts: string[] = [];
    if (nums.length >= 2) {
      const sum = nums.reduce((a, b) => a + b, 0);
      parts.push(`Average: <b>${fmt(sum / nums.length)}</b>`);
      parts.push(`Count: <b>${count}</b>`);
      parts.push(`Min: <b>${fmt(Math.min(...nums))}</b>`);
      parts.push(`Max: <b>${fmt(Math.max(...nums))}</b>`);
      parts.push(`Sum: <b>${fmt(sum)}</b>`);
    } else if (count > 1) {
      parts.push(`Count: <b>${count}</b>`);
    }
    this.statusBar.innerHTML = parts.map((p) => `<span class="sb-stat">${p}</span>`).join("");
  }

  onContextMenu(row: number, col: number, clientX: number, clientY: number) {
    const sel = this.grid.selRange();
    const merged = !!this.wb.active.mergeAt(row, col);
    showContextMenu(clientX, clientY, [
      { label: "Cut", action: () => document.execCommand("cut") },
      { label: "Copy", action: () => document.execCommand("copy") },
      { label: "Paste", action: () => document.execCommand("paste") },
      { separator: true },
      { label: "Insert row above", action: () => this.insertRows(sel.r0, 1) },
      { label: "Insert column left", action: () => this.insertCols(sel.c0, 1) },
      { label: "Delete row(s)", action: () => this.deleteRows(sel.r0, sel.r1 - sel.r0 + 1) },
      { label: "Delete column(s)", action: () => this.deleteCols(sel.c0, sel.c1 - sel.c0 + 1) },
      { separator: true },
      merged
        ? { label: "Unmerge cells", action: () => this.unmergeCells() }
        : { label: "Merge cells", action: () => this.mergeCells() },
      { separator: true },
      { label: "Hide rows", action: () => this.hideRows() },
      { label: "Hide columns", action: () => this.hideCols() },
      { label: "Unhide", action: () => this.unhide() },
      { separator: true },
      { label: this.wb.active.getCell(row, col)?.note ? "Edit note…" : "Insert note…", action: () => this.editNote(row, col) },
      { separator: true },
      { label: "Sort A → Z", action: () => this.sortRange(true) },
      { label: "Sort Z → A", action: () => this.sortRange(false) },
      { label: this.wb.active.filter ? "Remove filter" : "Create filter", action: () => this.toggleFilter() },
      { label: "Freeze panes here", action: () => this.toggleFreeze() },
      { separator: true },
      { label: "Clear contents", action: () => this.grid.clearSelection() },
    ]);
  }

  // `prevWidth` is the width BEFORE the change (the grid passes the pre-drag
  // size; auto-fit omits it so we read the current width). We apply the new
  // width here so auto-fit actually takes effect, and undo restores prevWidth.
  resizeCol(col: number, width: number, prevWidth?: number) {
    const sheet = this.wb.active;
    const idx = this.wb.activeIndex;
    const prev = prevWidth !== undefined ? prevWidth : sheet.colWidths.get(col) ?? DEFAULT_COL_WIDTH;
    width = Math.max(20, Math.round(width));
    if (width === prev) {
      sheet.colWidths.set(col, width);
      this.recomputeRowHeightsForColumn(sheet, col);
      this.grid.render();
      return;
    }
    const apply = (w: number) => {
      const s = this.onSheet(idx);
      s.colWidths.set(col, w);
      this.recomputeRowHeightsForColumn(s, col);
      this.grid.render();
    };
    apply(width);
    this.pushCommand({ redo: () => apply(width), undo: () => apply(prev) });
    this.markDirty();
  }

  // `manual` = the user dragged this row (so auto-fit should leave it alone).
  resizeRow(row: number, height: number, prevHeight?: number, manual = true) {
    const sheet = this.wb.active;
    const idx = this.wb.activeIndex;
    const prev = prevHeight !== undefined ? prevHeight : sheet.rowHeights.get(row) ?? DEFAULT_ROW_HEIGHT;
    height = Math.max(12, Math.round(height));
    const setManual = (s: Sheet) => (manual ? s.manualRows.add(row) : s.manualRows.delete(row));
    if (height === prev) {
      sheet.rowHeights.set(row, height);
      setManual(sheet);
      this.grid.render();
      return;
    }
    const apply = (h: number) => {
      const s = this.onSheet(idx);
      s.rowHeights.set(row, h);
      setManual(s);
      this.grid.render();
    };
    apply(height);
    this.pushCommand({ redo: () => apply(height), undo: () => apply(prev) });
    this.markDirty();
  }

  // Auto-fit one column to its widest (non-wrapped) cell. Capped generously so
  // long content like reference numbers fully expands.
  autofitCol(col: number) {
    const sheet = this.wb.active;
    let max = 30;
    for (const [key, cell] of sheet.cells) {
      const { col: c } = parseKey(key);
      if (c !== col || cell.value === null || cell.format?.wrap) continue;
      const fmt = cell.format;
      this.measureCtx.font = `${fmt?.italic ? "italic " : ""}${fmt?.bold ? "700 " : ""}${fmt?.fontSize || 13}px -apple-system, Segoe UI, sans-serif`;
      const text = formatValue(cell.value, fmt?.numFmt);
      max = Math.max(max, this.measureCtx.measureText(text).width + 12);
    }
    this.resizeCol(col, Math.min(1000, Math.ceil(max)));
  }

  // Auto-fit one row's height to its content (wrapped cells span multiple lines).
  autofitRow(row: number) {
    const sheet = this.wb.active;
    let max = DEFAULT_ROW_HEIGHT;
    for (let c = 0; c <= sheet.maxCol; c++) {
      const cell = sheet.getCell(row, c);
      if (!cell || cell.value === null) continue;
      const fmt = cell.format;
      const fontSize = fmt?.fontSize || 13;
      if (fmt?.wrap) {
        this.measureCtx.font = `${fmt.italic ? "italic " : ""}${fmt.bold ? "700 " : ""}${fontSize}px -apple-system, Segoe UI, sans-serif`;
        const lines = wrapLines((s) => this.measureCtx.measureText(s).width, formatValue(cell.value, fmt.numFmt), sheet.colWidth(c) - 8);
        max = Math.max(max, lines.length * fontSize * 1.35 + 6);
      } else {
        max = Math.max(max, fontSize * 1.6);
      }
    }
    // double-click auto-fit clears any manual height for this row
    this.resizeRow(row, Math.ceil(max), undefined, false);
  }

  // ===== Quick actions =====

  // Alt+= : insert SUM of the numeric run above (or to the left) of the active cell.
  autoSum() {
    const { row, col } = this.grid.active;
    const sheet = this.wb.active;
    let r = row - 1;
    while (r >= 0 && typeof sheet.getCell(r, col)?.value === "number") r--;
    if (r + 1 < row) {
      this.commit(row, col, `=SUM(${formatA1(r + 1, col)}:${formatA1(row - 1, col)})`);
      this.grid.setActive(row + 1, col);
      return;
    }
    let c = col - 1;
    while (c >= 0 && typeof sheet.getCell(row, c)?.value === "number") c--;
    if (c + 1 < col) {
      this.commit(row, col, `=SUM(${formatA1(row, c + 1)}:${formatA1(row, col - 1)})`);
      this.grid.setActive(row, col + 1);
      return;
    }
    this.grid.beginEdit("=SUM(");
  }

  // Ctrl+D : fill the top row of the selection down (or the cell above into a
  // single active cell).
  fillDown() {
    const sel = this.grid.selRange();
    const sheet = this.wb.active;
    if (sel.r0 === sel.r1 && sel.c0 === sel.c1) {
      if (sel.r0 === 0) return;
      this.copyCellInto(sel.r0 - 1, sel.c0, sel.r0, sel.c0, 1, 0);
    } else {
      for (let c = sel.c0; c <= sel.c1; c++)
        for (let r = sel.r0 + 1; r <= sel.r1; r++) this.copyCellInto(sel.r0, c, r, c, r - sel.r0, 0);
    }
    void sheet;
    this.grid.render();
  }

  // Ctrl+R : fill the left column of the selection right.
  fillRight() {
    const sel = this.grid.selRange();
    if (sel.r0 === sel.r1 && sel.c0 === sel.c1) {
      if (sel.c0 === 0) return;
      this.copyCellInto(sel.r0, sel.c0 - 1, sel.r0, sel.c0, 0, 1);
    } else {
      for (let r = sel.r0; r <= sel.r1; r++)
        for (let c = sel.c0 + 1; c <= sel.c1; c++) this.copyCellInto(r, sel.c0, r, c, 0, c - sel.c0);
    }
    this.grid.render();
  }

  // Copy raw (offsetting relative formulas) + format from src to dest.
  private copyCellInto(sr: number, sc: number, dr: number, dc: number, dRow: number, dCol: number) {
    const sheet = this.wb.active;
    const srcRaw = sheet.getRaw(sr, sc);
    const raw = srcRaw.startsWith("=") ? offsetFormula(srcRaw, dRow, dCol) : srcRaw;
    this.commitRaw(dr, dc, raw);
    const srcFmt = sheet.getCell(sr, sc)?.format;
    if (srcFmt) sheet.ensureCell(dr, dc).format = { ...srcFmt };
  }

  insertDate() {
    const d = new Date();
    const { row, col } = this.grid.active;
    this.commit(row, col, `=DATE(${d.getFullYear()},${d.getMonth() + 1},${d.getDate()})`);
    this.applyFormat({ numFmt: "yyyy-mm-dd" });
  }

  insertTime() {
    const d = new Date();
    const frac = (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400;
    const { row, col } = this.grid.active;
    this.commit(row, col, String(Math.round(frac * 1e9) / 1e9));
    this.applyFormat({ numFmt: "h:mm:ss" });
  }

  changeDecimals(delta: number) {
    const { row, col } = this.grid.active;
    const cur = this.wb.active.getFormat(row, col)?.numFmt;
    this.applyFormat({ numFmt: adjustDecimals(cur, delta) });
  }

  fillRange(src: SelRange, dest: SelRange) {
    const sheet = this.wb.active;
    const vertical = dest.r1 > src.r1;
    if (vertical) {
      for (let c = src.c0; c <= src.c1; c++) {
        const colCells: { raw: string }[] = [];
        for (let r = src.r0; r <= src.r1; r++) colCells.push({ raw: sheet.getRaw(r, c) });
        const nums = colCells.map((x) => Number(x.raw)).filter((n) => !Number.isNaN(n));
        const allNum = nums.length === colCells.length && colCells.every((x) => x.raw !== "");
        const step = nums.length >= 2 ? nums[nums.length - 1] - nums[nums.length - 2] : nums.length === 1 ? 0 : 1;
        for (let r = src.r1 + 1; r <= dest.r1; r++) {
          const offset = r - src.r1;
          const tileIdx = (r - src.r0) % colCells.length;
          const srcRaw = colCells[tileIdx].raw;
          if (srcRaw.startsWith("=")) this.commitRaw(r, c, offsetFormula(srcRaw, r - (src.r0 + tileIdx), 0));
          else if (allNum && nums.length >= 1) this.commitRaw(r, c, String(nums[nums.length - 1] + step * offset));
          else this.commitRaw(r, c, srcRaw);
        }
      }
    } else {
      for (let r = src.r0; r <= src.r1; r++) {
        const rowCells: { raw: string }[] = [];
        for (let c = src.c0; c <= src.c1; c++) rowCells.push({ raw: sheet.getRaw(r, c) });
        const nums = rowCells.map((x) => Number(x.raw)).filter((n) => !Number.isNaN(n));
        const allNum = nums.length === rowCells.length && rowCells.every((x) => x.raw !== "");
        const step = nums.length >= 2 ? nums[nums.length - 1] - nums[nums.length - 2] : nums.length === 1 ? 0 : 1;
        for (let c = src.c1 + 1; c <= dest.c1; c++) {
          const offset = c - src.c1;
          const tileIdx = (c - src.c0) % rowCells.length;
          const srcRaw = rowCells[tileIdx].raw;
          if (srcRaw.startsWith("=")) this.commitRaw(r, c, offsetFormula(srcRaw, 0, c - (src.c0 + tileIdx)));
          else if (allNum && nums.length >= 1) this.commitRaw(r, c, String(nums[nums.length - 1] + step * offset));
          else this.commitRaw(r, c, srcRaw);
        }
      }
    }
    this.grid.render();
  }

  // ===== Structural edits (snapshot-based undo) =====

  private structural(mutate: (sheet: Sheet) => void) {
    const idx = this.wb.activeIndex;
    // Snapshot every sheet: a structural edit can rewrite cross-sheet formulas
    // anywhere in the workbook, so undo must be able to restore all of them.
    const before = this.wb.sheets.map((s) => s.snapshot());
    mutate(this.wb.active);
    this.engine.rebuild();
    const after = this.wb.sheets.map((s) => s.snapshot());
    this.pushCommand({
      redo: () => {
        after.forEach((snap, i) => this.wb.sheets[i].restore(snap));
        this.refreshAfterStructural(idx);
      },
      undo: () => {
        before.forEach((snap, i) => this.wb.sheets[i].restore(snap));
        this.refreshAfterStructural(idx);
      },
    });
    this.markDirty();
    this.grid.render();
    this.onSelectionChange();
  }

  insertRows(at: number, count: number) {
    this.structural((sheet) => {
      sheet.insertRows(at, count);
      this.rewriteAll(sheet, (r, c) => (r >= at ? { row: r + count, col: c } : { row: r, col: c }));
    });
  }
  deleteRows(at: number, count: number) {
    this.structural((sheet) => {
      sheet.deleteRows(at, count);
      this.rewriteAll(sheet, (r, c) => {
        if (r >= at && r < at + count) return null;
        return { row: r >= at + count ? r - count : r, col: c };
      });
    });
  }
  insertCols(at: number, count: number) {
    this.structural((sheet) => {
      sheet.insertCols(at, count);
      this.rewriteAll(sheet, (r, c) => (c >= at ? { row: r, col: c + count } : { row: r, col: c }));
    });
  }
  deleteCols(at: number, count: number) {
    this.structural((sheet) => {
      sheet.deleteCols(at, count);
      this.rewriteAll(sheet, (r, c) => {
        if (c >= at && c < at + count) return null;
        return { row: r, col: c >= at + count ? c - count : c };
      });
    });
  }

  // Rewrite every formula in the WORKBOOK using a coordinate transform, but only
  // for references that actually target the edited sheet `structSheet`. A
  // reference qualified with another sheet (or unqualified but living on another
  // sheet) is left untouched.
  private rewriteAll(
    structSheet: Sheet,
    base: (row: number, col: number) => { row: number; col: number } | null
  ) {
    const structName = structSheet.name.toLowerCase();
    for (const home of this.wb.sheets) {
      for (const [, cell] of home.cells) {
        if (!cell.raw.startsWith("=")) continue;
        const transform: RefTransform = (r, c, refSheet) => {
          const targetsStruct =
            refSheet !== undefined ? refSheet.toLowerCase() === structName : home === structSheet;
          if (!targetsStruct) return undefined;
          return base(r, c);
        };
        const nextRaw = rewriteFormulaRefs(cell.raw, transform);
        if (nextRaw !== cell.raw) cell.raw = nextRaw;
      }
    }
  }

  // ===== Merge =====

  mergeCells() {
    const sel = this.grid.selRange();
    if (sel.r0 === sel.r1 && sel.c0 === sel.c1) return;
    this.structural((sheet) => {
      // keep only the top-left value; clear the rest
      for (let r = sel.r0; r <= sel.r1; r++) {
        for (let c = sel.c0; c <= sel.c1; c++) {
          if (r === sel.r0 && c === sel.c0) continue;
          if (sheet.getRaw(r, c) !== "") this.engine.setCellRaw(r, c, "");
        }
      }
      sheet.addMerge({ r0: sel.r0, c0: sel.c0, r1: sel.r1, c1: sel.c1 });
    });
  }

  unmergeCells() {
    const { row, col } = this.grid.active;
    this.structural((sheet) => sheet.removeMergeAt(row, col));
  }

  // ===== Freeze panes =====

  // Toggle freezing rows above / columns left of the active cell (Excel-style).
  toggleFreeze() {
    const s = this.wb.active;
    const { row, col } = this.grid.active;
    if (s.frozenRows || s.frozenCols) {
      s.frozenRows = 0;
      s.frozenCols = 0;
    } else {
      s.frozenRows = row;
      s.frozenCols = col;
    }
    this.grid.resetScroll();
    this.grid.render();
    this.markDirty();
  }

  isFrozen(): boolean {
    return this.wb.active.frozenRows > 0 || this.wb.active.frozenCols > 0;
  }

  // ===== Sort & filter =====

  // Sort the rows of the selected range by the active cell's column. Relative
  // formula references are offset by how far each row moved (as Excel does).
  sortRange(ascending: boolean) {
    const sel = this.grid.selRange();
    if (sel.r0 === sel.r1) return;
    const keyCol = Math.min(Math.max(this.grid.active.col, sel.c0), sel.c1);
    this.sortBlock(sel, keyCol, ascending);
  }

  // Sort the data rows of `range` by `keyCol`; relative formulas re-anchor.
  private sortBlock(range: { r0: number; c0: number; r1: number; c1: number }, keyCol: number, ascending: boolean) {
    if (range.r1 <= range.r0) return;
    this.structural((sheet) => {
      interface RowData { origRow: number; key: CellValue; cells: { raw: string; format?: CellFormat }[] }
      const rows: RowData[] = [];
      for (let r = range.r0; r <= range.r1; r++) {
        const cells: { raw: string; format?: CellFormat }[] = [];
        for (let c = range.c0; c <= range.c1; c++) {
          const cell = sheet.getCell(r, c);
          cells.push({ raw: cell?.raw ?? "", format: cell?.format ? { ...cell.format } : undefined });
        }
        rows.push({ origRow: r, key: sheet.getCell(r, keyCol)?.value ?? null, cells });
      }
      const dir = ascending ? 1 : -1;
      const sorted = rows
        .map((r, i) => ({ r, i }))
        .sort((a, b) => compareForSort(a.r.key, b.r.key) * dir || a.i - b.i)
        .map((x) => x.r);
      sorted.forEach((rowData, idx) => {
        const newRow = range.r0 + idx;
        const dRow = newRow - rowData.origRow;
        rowData.cells.forEach((cd, ci) => {
          const c = range.c0 + ci;
          const raw = cd.raw.startsWith("=") && dRow !== 0 ? offsetFormula(cd.raw, dRow, 0) : cd.raw;
          this.engine.setCellRaw(newRow, c, raw);
          if (cd.format) sheet.ensureCell(newRow, c).format = { ...cd.format };
          else {
            const cc = sheet.getCell(newRow, c);
            if (cc) cc.format = undefined;
          }
        });
      });
    });
  }

  // ===== Filter (Google Sheets style) =====

  toggleFilter() {
    if (this.wb.active.filter) this.removeFilter();
    else this.createFilter();
  }

  createFilter() {
    const sheet = this.wb.active;
    const sel = this.grid.selRange();
    let range: { r0: number; c0: number; r1: number; c1: number };
    if (sel.r0 === sel.r1 && sel.c0 === sel.c1) {
      // single cell → cover the whole used data block
      let r0 = Infinity, c0 = Infinity, r1 = 0, c1 = 0, any = false;
      for (const [k, cell] of sheet.cells) {
        if (cell.value === null || cell.value === "") continue;
        const { row, col } = parseKey(k);
        any = true;
        r0 = Math.min(r0, row); c0 = Math.min(c0, col); r1 = Math.max(r1, row); c1 = Math.max(c1, col);
      }
      range = any ? { r0, c0, r1, c1 } : { r0: sel.r0, c0: sel.c0, r1: sel.r0, c1: sel.c0 };
    } else {
      range = { r0: sel.r0, c0: sel.c0, r1: sel.r1, c1: sel.c1 };
    }
    sheet.filter = { range, cols: {} };
    this.markDirty();
    this.grid.render();
  }

  removeFilter() {
    const sheet = this.wb.active;
    if (!sheet.filter) return;
    for (let r = sheet.filter.range.r0; r <= sheet.filter.range.r1; r++) sheet.hiddenRows.delete(r);
    sheet.filter = null;
    this.markDirty();
    this.grid.render();
  }

  onFilterDropdown(col: number, clientX: number, clientY: number) {
    showColumnFilter(this, col, clientX, clientY);
  }

  // distinct display values in a filtered column's data rows
  filterColumnValues(col: number): string[] {
    const sheet = this.wb.active;
    const f = sheet.filter;
    if (!f) return [];
    const seen = new Set<string>();
    for (let r = f.range.r0 + 1; r <= f.range.r1; r++) {
      const cell = sheet.getCell(r, col);
      seen.add(formatValue(cell?.value ?? null, cell?.format?.numFmt));
    }
    return [...seen].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  }

  filterColumnHidden(col: number): string[] {
    return this.wb.active.filter?.cols[col] ?? [];
  }

  applyColumnFilter(col: number, hiddenValues: string[]) {
    const f = this.wb.active.filter;
    if (!f) return;
    if (hiddenValues.length) f.cols[col] = hiddenValues;
    else delete f.cols[col];
    this.recomputeFilter();
    this.markDirty();
  }

  sortFilterColumn(col: number, ascending: boolean) {
    const f = this.wb.active.filter;
    if (!f) return;
    this.sortBlock({ r0: f.range.r0 + 1, c0: f.range.c0, r1: f.range.r1, c1: f.range.c1 }, col, ascending);
    this.recomputeFilter();
  }

  // Hide a row if ANY column's filter excludes its value (AND across columns).
  private recomputeFilter() {
    const sheet = this.wb.active;
    const f = sheet.filter;
    if (!f) return;
    const cols = Object.keys(f.cols).map(Number).filter((c) => f.cols[c].length);
    for (let r = f.range.r0 + 1; r <= f.range.r1; r++) {
      let hide = false;
      for (const c of cols) {
        const cell = sheet.getCell(r, c);
        const disp = formatValue(cell?.value ?? null, cell?.format?.numFmt);
        if (f.cols[c].includes(disp)) { hide = true; break; }
      }
      if (hide) sheet.hiddenRows.add(r);
      else sheet.hiddenRows.delete(r);
    }
    this.grid.render();
  }

  // ===== Undo / redo =====

  private pushDelta(delta: CellDelta) {
    if (!this.pendingDeltas) {
      this.pendingDeltas = [];
      queueMicrotask(() => {
        const deltas = this.pendingDeltas;
        this.pendingDeltas = null;
        if (deltas && deltas.length) this.pushCommand(this.cellEditCommand(deltas));
      });
    }
    this.pendingDeltas.push(delta);
  }

  private cellEditCommand(deltas: CellDelta[]): Command {
    const apply = (dir: "undo" | "redo") => {
      const idx = deltas[0].sheetIndex;
      this.activateSheet(idx);
      const list = dir === "undo" ? [...deltas].reverse() : deltas;
      const rows = new Set<number>();
      for (const d of list) {
        const raw = dir === "undo" ? d.oldRaw : d.newRaw;
        const fmt = dir === "undo" ? d.oldFormat : d.newFormat;
        this.engine.setCellRaw(d.row, d.col, raw);
        if (fmt) this.wb.active.ensureCell(d.row, d.col).format = { ...fmt };
        else {
          const c = this.wb.active.getCell(d.row, d.col);
          if (c) {
            c.format = undefined;
            this.wb.active.deleteCellIfEmpty(d.row, d.col);
          }
        }
        rows.add(d.row);
      }
      // Re-fit row heights so undo/redo of a font/content change resizes rows
      // back too (the auto-fit isn't itself recorded in the delta).
      for (const r of rows) this.recomputeRowHeight(this.wb.active, r);
      this.grid.render();
      this.onSelectionChange();
    };
    return { undo: () => apply("undo"), redo: () => apply("redo") };
  }

  private pushCommand(cmd: Command) {
    this.undoStack.push(cmd);
    this.redoStack = [];
  }

  undo() {
    // flush any pending cell-edit group first
    this.flushPending();
    const cmd = this.undoStack.pop();
    if (!cmd) return;
    cmd.undo();
    this.redoStack.push(cmd);
  }

  redo() {
    const cmd = this.redoStack.pop();
    if (!cmd) return;
    cmd.redo();
    this.undoStack.push(cmd);
  }

  private flushPending() {
    if (this.pendingDeltas && this.pendingDeltas.length) {
      const deltas = this.pendingDeltas;
      this.pendingDeltas = null;
      this.pushCommand(this.cellEditCommand(deltas));
    }
  }

  // ===== Formatting =====

  applyFormat(patch: Partial<CellFormat>) {
    const sel = this.grid.selRange();
    for (let r = sel.r0; r <= sel.r1; r++) {
      for (let c = sel.c0; c <= sel.c1; c++) {
        const cell = this.wb.active.ensureCell(r, c);
        const oldFormat = cell.format ? { ...cell.format } : undefined;
        const next: CellFormat = { ...(cell.format || {}), ...patch };
        for (const k of Object.keys(next) as (keyof CellFormat)[]) {
          if (next[k] === false || next[k] === undefined) delete next[k];
        }
        cell.format = Object.keys(next).length ? next : undefined;
        this.wb.active.deleteCellIfEmpty(r, c);
        this.pushDelta({
          sheetIndex: this.wb.activeIndex,
          row: r,
          col: c,
          oldRaw: cell.raw,
          newRaw: cell.raw,
          oldFormat,
          newFormat: cell.format ? { ...cell.format } : undefined,
        });
      }
    }
    this.markDirty();
    this.grid.render();
    this.refreshToolbarState();
    // if a cell is open for editing, live-preview the new font in the editor
    if (this.grid.isEditing()) this.grid.refreshEditorStyle();
  }

  toggleFormat(key: "bold" | "italic" | "underline" | "strike") {
    const { row, col } = this.grid.active;
    const current = this.wb.active.getFormat(row, col)?.[key];
    this.applyFormat({ [key]: !current } as Partial<CellFormat>);
  }

  // ===== Font family + size (control speaks points; model stores px) =====

  setFontFamily(family: string) {
    this.applyFormat({ fontFamily: family || undefined });
  }

  setFontSize(pt: number) {
    if (!pt || pt < 1) return;
    this.applyFormat({ fontSize: Math.round((pt * 4) / 3) });
    const sel = this.grid.selRange();
    for (let r = sel.r0; r <= sel.r1; r++) this.recomputeRowHeight(this.wb.active, r);
    this.grid.render();
  }

  activeFontPt(): number {
    return Math.round((this.activeFormat()?.fontSize ?? 13) * 0.75);
  }
  activeFontFamily(): string {
    return this.activeFormat()?.fontFamily ?? "";
  }

  // Wrap text in the selection. Toggling on recomputes affected row heights so
  // the wrapped text fits; toggling off resets those rows to the default height.
  toggleWrap() {
    const { row, col } = this.grid.active;
    const turnOn = !this.wb.active.getFormat(row, col)?.wrap;
    const sel = this.grid.selRange();
    this.structural((sheet) => {
      for (let r = sel.r0; r <= sel.r1; r++) {
        for (let c = sel.c0; c <= sel.c1; c++) {
          const cell = sheet.ensureCell(r, c);
          const f = { ...(cell.format || {}) };
          if (turnOn) f.wrap = true;
          else delete f.wrap;
          cell.format = Object.keys(f).length ? f : undefined;
          if (!cell.raw && !cell.format) sheet.deleteCellIfEmpty(r, c);
        }
      }
      for (let r = sel.r0; r <= sel.r1; r++) this.recomputeRowHeight(sheet, r);
    });
  }

  // Recompute one row's height from its wrapped cells (tallest wins).
  // Auto-fit a row's height to its tallest cell — by font size, or by wrapped
  // line count. Skips rows the user sized by hand. (Excel/Sheets behavior.)
  private recomputeRowHeight(sheet: Sheet, row: number) {
    if (sheet.manualRows.has(row)) return;
    let needed = DEFAULT_ROW_HEIGHT;
    for (let c = 0; c <= sheet.maxCol; c++) {
      const cell = sheet.getCell(row, c);
      if (!cell || cell.value === null || cell.value === "") continue;
      const fmt = cell.format;
      const fontSize = fmt?.fontSize || 13;
      if (fmt?.wrap) {
        this.measureCtx.font = `${fmt.italic ? "italic " : ""}${fmt.bold ? "700 " : ""}${fontSize}px -apple-system, Segoe UI, sans-serif`;
        const lines = wrapLines((s) => this.measureCtx.measureText(s).width, formatValue(cell.value, fmt.numFmt), sheet.colWidth(c) - 8);
        needed = Math.max(needed, lines.length * fontSize * 1.35 + 6);
      } else {
        needed = Math.max(needed, fontSize * 1.3 + 5);
      }
    }
    needed = Math.ceil(needed);
    if (needed > DEFAULT_ROW_HEIGHT) sheet.rowHeights.set(row, needed);
    else sheet.rowHeights.delete(row);
  }

  // A column resize can change wrapped-row heights in that column.
  private recomputeRowHeightsForColumn(sheet: Sheet, col: number) {
    for (let r = 0; r <= sheet.maxRow; r++) {
      if (sheet.getCell(r, col)?.format?.wrap) this.recomputeRowHeight(sheet, r);
    }
  }

  // Widen every used column to fit its widest cell.
  autofitAllColumns() {
    this.structural((sheet) => {
      const widths = new Map<number, number>();
      for (const [key, cell] of sheet.cells) {
        const { row, col } = parseKey(key);
        if (cell.value === null) continue;
        const fmt = cell.format;
        if (fmt?.wrap) continue; // wrapped cells don't drive column width
        this.measureCtx.font = `${fmt?.bold ? "700 " : ""}${fmt?.fontSize || 13}px -apple-system, Segoe UI, sans-serif`;
        const t = formatValue(cell.value, fmt?.numFmt);
        const wpx = this.measureCtx.measureText(t).width + 12;
        widths.set(col, Math.max(widths.get(col) ?? 0, wpx));
        void row;
      }
      for (const [col, wpx] of widths) sheet.colWidths.set(col, Math.min(420, Math.max(40, Math.ceil(wpx))));
    });
  }

  activeFormat(): CellFormat | undefined {
    const { row, col } = this.grid.active;
    return this.wb.active.getFormat(row, col);
  }

  private refreshToolbarState() {
    this.toolbarEl.dispatchEvent(new CustomEvent("refresh"));
  }

  // ===== Sheet switching helpers =====

  private onSheet(idx: number): Sheet {
    return this.wb.sheets[idx];
  }

  private activateSheet(idx: number) {
    if (this.wb.activeIndex !== idx) {
      this.wb.setActive(idx);
      this.engine.setSheet(this.wb.active);
      this.grid.setSheet(this.wb.active);
      this.renderTabs();
    } else {
      this.engine.setSheet(this.wb.active);
    }
    this.renderCharts();
  }

  // Force a full switch to a sheet + UI refresh. Used by add/duplicate/delete,
  // where activeIndex may already equal idx (so activateSheet would no-op the
  // grid/tab refresh).
  private showSheet(idx: number) {
    this.wb.setActive(idx);
    this.engine.setSheet(this.wb.active);
    this.grid.setSheet(this.wb.active);
    this.renderTabs();
    this.renderCharts();
  }

  private refreshAfterStructural(idx: number) {
    this.activateSheet(idx);
    this.engine.rebuild();
    this.condVersion++; // rules/charts may have changed; invalidate caches
    this.grid.render();
    this.renderCharts();
    this.onSelectionChange();
  }

  // ===== Formula bar =====

  private wireFormulaBar() {
    this.nameBox.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      const val = this.nameBox.value.trim();
      const ref = parseA1(val);
      if (ref) {
        this.grid.setActive(ref.row, ref.col);
      } else if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(val)) {
        // navigate to an existing name, or define a new one over the selection
        if (!this.gotoName(val)) this.defineName(val);
      }
      (e.target as HTMLInputElement).blur();
    });
    const fbAuto = new FormulaAutocomplete(this.formulaInput);
    this.formulaInput.addEventListener("focus", () => {
      if (!this.grid.isEditing()) this.grid.beginEdit(this.formulaInput.value);
    });
    this.formulaInput.addEventListener("input", () => {
      fbAuto.onInput();
      this.grid.syncEditorFromFormulaBar(this.formulaInput.value);
    });
    this.formulaInput.addEventListener("keydown", (e) => {
      if (fbAuto.handleKeydown(e)) return;
      if (e.key === "Enter") {
        e.preventDefault();
        fbAuto.close();
        this.grid.commitEdit(1, 0);
      } else if (e.key === "Escape") {
        fbAuto.close();
        this.grid.cancelEdit();
        this.onSelectionChange();
      }
    });
  }

  // ===== Sheet tabs =====

  renderTabs() {
    this.tabsEl.innerHTML = "";
    this.wb.sheets.forEach((s, i) => {
      const tab = document.createElement("div");
      tab.className = "sheet-tab" + (i === this.wb.activeIndex ? " active" : "");
      tab.textContent = s.name;
      tab.addEventListener("click", () => this.switchSheet(i));
      tab.addEventListener("dblclick", () => this.renameSheet(i));
      tab.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        showContextMenu(e.clientX, e.clientY, [
          { label: "Rename…", action: () => this.renameSheet(i) },
          { label: "Duplicate", action: () => this.duplicateSheet(i) },
          { label: "Delete", disabled: this.wb.sheets.length <= 1, action: () => this.deleteSheet(i) },
          { separator: true },
          { label: "Insert sheet", action: () => { this.wb.addSheet(); this.showSheet(this.wb.activeIndex); this.markDirty(); } },
        ]);
      });
      this.tabsEl.appendChild(tab);
    });
    const add = document.createElement("div");
    add.className = "sheet-add";
    add.textContent = "+";
    add.title = "Add sheet";
    add.addEventListener("click", () => {
      this.wb.addSheet();
      this.showSheet(this.wb.activeIndex);
      this.markDirty();
    });
    this.tabsEl.appendChild(add);
  }

  private switchSheet(i: number) {
    if (i === this.wb.activeIndex) return;
    if (this.grid.isEditing()) this.grid.commitEdit(0, 0);
    this.activateSheet(i);
  }

  private async renameSheet(i: number) {
    const oldName = this.wb.sheets[i].name;
    const input = await showPrompt("Sheet name:", oldName);
    const newName = input?.trim();
    if (!newName || newName === oldName) return;
    if (this.wb.sheets.some((s, j) => j !== i && s.name.toLowerCase() === newName.toLowerCase())) {
      alert(`A sheet named "${newName}" already exists.`);
      return;
    }
    this.wb.sheets[i].name = newName;
    // update every formula that referenced the sheet by its old name
    for (const home of this.wb.sheets) {
      for (const [, cell] of home.cells) {
        if (cell.raw.startsWith("=")) {
          const next = renameSheetRefs(cell.raw, oldName, newName);
          if (next !== cell.raw) cell.raw = next;
        }
      }
    }
    this.engine.rebuild();
    this.renderTabs();
    this.grid.render();
    this.markDirty();
  }

  private uniqueSheetName(base: string): string {
    const names = new Set(this.wb.sheets.map((s) => s.name.toLowerCase()));
    if (!names.has(base.toLowerCase())) return base;
    let n = 2;
    while (names.has(`${base} ${n}`.toLowerCase())) n++;
    return `${base} ${n}`;
  }

  private duplicateSheet(i: number) {
    const src = this.wb.sheets[i];
    const copy = new Sheet(this.uniqueSheetName(`${src.name} copy`));
    copy.restore(src.snapshot()); // deep-copies cells, formats, merges, etc.
    this.wb.sheets.splice(i + 1, 0, copy);
    this.engine.rebuild();
    this.showSheet(i + 1);
    this.markDirty();
  }

  private deleteSheet(i: number) {
    if (this.wb.sheets.length <= 1) return;
    if (!confirm(`Delete sheet "${this.wb.sheets[i].name}"? This can't be undone.`)) return;
    this.wb.sheets.splice(i, 1);
    this.engine.rebuild();
    this.showSheet(Math.min(this.wb.activeIndex, this.wb.sheets.length - 1));
    this.markDirty();
  }

  // ===== Clipboard =====

  // Snapshot the selection (raw + format) into the internal clipboard, and put
  // a TSV of the displayed values on the system clipboard for external apps.
  private captureClip(e: ClipboardEvent, cut: boolean) {
    const sheet = this.wb.active;
    const sel = this.grid.selRange();
    const rows: { raw: string; format?: CellFormat }[][] = [];
    for (let r = sel.r0; r <= sel.r1; r++) {
      const row: { raw: string; format?: CellFormat }[] = [];
      for (let c = sel.c0; c <= sel.c1; c++) {
        const cell = sheet.getCell(r, c);
        row.push({ raw: cell?.raw ?? "", format: cell?.format ? { ...cell.format } : undefined });
      }
      rows.push(row);
    }
    const tsv = sheetToTsv(sheet, sel.r0, sel.c0, sel.r1, sel.c1);
    this.internalClip = { rows, r0: sel.r0, c0: sel.c0, cut };
    this.clipTsv = tsv;
    e.clipboardData?.setData("text/plain", tsv);
    e.preventDefault();
  }

  private wireClipboard() {
    document.addEventListener("copy", (e) => {
      if (this.grid.isEditing()) return;
      this.captureClip(e, false);
    });
    document.addEventListener("cut", (e) => {
      if (this.grid.isEditing()) return;
      this.captureClip(e, true);
    });
    document.addEventListener("paste", (e) => {
      if (this.grid.isEditing()) return;
      const text = e.clipboardData?.getData("text/plain") ?? "";
      e.preventDefault();
      const { row, col } = this.grid.active;

      // Our own copy → paste with formulas + formatting (relative refs adjusted).
      if (this.internalClip && text === this.clipTsv) {
        const { rows, r0, c0, cut } = this.internalClip;
        const dRow = row - r0;
        const dCol = col - c0;
        rows.forEach((rowArr, dr) => {
          rowArr.forEach((cellData, dc) => {
            const tr = row + dr;
            const tc = col + dc;
            const raw = cellData.raw.startsWith("=") ? offsetFormula(cellData.raw, dRow, dCol) : cellData.raw;
            this.commitRaw(tr, tc, raw);
            const cell = this.wb.active.ensureCell(tr, tc);
            cell.format = cellData.format ? { ...cellData.format } : undefined;
            this.wb.active.deleteCellIfEmpty(tr, tc);
          });
        });
        if (cut) {
          // a cut is a move: clear the source cells, then disarm
          for (let r = r0; r < r0 + rows.length; r++)
            for (let c = c0; c < c0 + rows[0].length; c++)
              if (r < row || r >= row + rows.length || c < col || c >= col + rows[0].length) this.commitRaw(r, c, "");
          this.internalClip = null;
        }
        this.grid.render();
        return;
      }

      // External text → paste displayed values as a TSV block.
      if (!text) return;
      const lines = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").replace(/\n$/, "").split("\n");
      lines.forEach((line, dr) => {
        line.split("\t").forEach((val, dc) => this.commitRaw(row + dr, col + dc, val));
      });
      this.grid.render();
    });
  }

  // ===== Find & replace =====

  openFindReplace() {
    showFindReplace(this);
  }

  // Search the active sheet's raw text; returns matching cell coords in order.
  findMatches(query: string, matchCase: boolean): { row: number; col: number }[] {
    const out: { row: number; col: number }[] = [];
    const q = matchCase ? query : query.toLowerCase();
    for (const [key, cell] of this.wb.active.cells) {
      const text = cell.raw;
      const hay = matchCase ? text : text.toLowerCase();
      if (q !== "" && hay.includes(q)) {
        out.push(parseKey(key));
      }
    }
    out.sort((a, b) => (a.row - b.row) || (a.col - b.col));
    return out;
  }

  gotoCell(row: number, col: number) {
    this.grid.setActive(row, col);
  }

  replaceAll(query: string, replacement: string, matchCase: boolean): number {
    let count = 0;
    const matches = this.findMatches(query, matchCase);
    for (const { row, col } of matches) {
      const raw = this.wb.active.getRaw(row, col);
      const next = matchCase
        ? raw.split(query).join(replacement)
        : replaceCaseInsensitive(raw, query, replacement);
      if (next !== raw) {
        this.commitRaw(row, col, next);
        count++;
      }
    }
    this.grid.render();
    return count;
  }

  replaceOne(row: number, col: number, query: string, replacement: string, matchCase: boolean) {
    const raw = this.wb.active.getRaw(row, col);
    const next = matchCase
      ? raw.replace(query, replacement)
      : replaceCaseInsensitive(raw, query, replacement, true);
    if (next !== raw) this.commitRaw(row, col, next);
    this.grid.render();
  }

  // ===== Menu actions =====

  private wireMenus() {
    const native = (window as any).native;
    if (!native) return;
    native.onMenu("menu:new", () => this.newWorkbook());
    native.onMenu("menu:open", () => this.openFile());
    native.onMenu("menu:save", () => this.save());
    native.onMenu("menu:saveAs", () => this.saveAs());
    native.onMenu("menu:undo", () => this.undo());
    native.onMenu("menu:redo", () => this.redo());
    native.onMenu("menu:importCsv", () => this.importCsv());
    native.onMenu("menu:exportCsv", () => this.exportCsv());
    native.onMenu("menu:exportPdf", () => this.exportPdf());
    native.onMenu("menu:zoomIn", () => this.zoomIn());
    native.onMenu("menu:zoomOut", () => this.zoomOut());
    native.onMenu("menu:zoomReset", () => this.zoomReset());
  }

  newWorkbook() {
    this.wb = new Workbook();
    this.engine = new CalcEngine(this.wb);
    this.grid.setSheet(this.wb.active);
    this.undoStack = [];
    this.redoStack = [];
    this.renderTabs();
    this.renderCharts();
    document.title = this.titleText();
  }

  async openFile() {
    const native = (window as any).native;
    if (!native) return;
    const res = await native.openXlsx();
    if (!res) return;
    try {
      this.wb = readXlsx(res.data);
      this.wb.filePath = res.path;
      this.engine = new CalcEngine(this.wb);
      this.engine.rebuild();
      this.grid.setSheet(this.wb.active);
      this.undoStack = [];
      this.redoStack = [];
      this.renderTabs();
      this.renderCharts();
      this.wb.dirty = false;
      document.title = this.titleText();
    } catch (err) {
      alert("Could not open file: " + err);
    }
  }

  async save() {
    if (!this.wb.filePath) return this.saveAs();
    const native = (window as any).native;
    const data = writeXlsx(this.wb);
    const res = await native.saveFile(this.wb.filePath, data.buffer);
    if (res) {
      this.wb.dirty = false;
      document.title = this.titleText();
    }
  }

  async saveAs() {
    const native = (window as any).native;
    if (!native) return;
    const data = writeXlsx(this.wb);
    const res = await native.saveXlsxAs(data.buffer, this.wb.filePath || `${this.wb.sheets[0].name}.xlsx`);
    if (res) {
      this.wb.filePath = res.path;
      this.wb.dirty = false;
      document.title = this.titleText();
    }
  }

  // Used by the main process's close handler.
  isDirty(): boolean {
    return this.wb.dirty;
  }

  // Save for the close-prompt; returns true if persisted (false if the user
  // cancelled the Save As dialog).
  async saveForClose(): Promise<boolean> {
    const native = (window as any).native;
    if (!native) return true;
    const data = writeXlsx(this.wb);
    if (this.wb.filePath) {
      const res = await native.saveFile(this.wb.filePath, data.buffer);
      if (res) this.wb.dirty = false;
      return !!res;
    }
    const res = await native.saveXlsxAs(data.buffer, `${this.wb.sheets[0].name}.xlsx`);
    if (res) {
      this.wb.filePath = res.path;
      this.wb.dirty = false;
      document.title = this.titleText();
    }
    return !!res;
  }

  // ===== CSV / PDF =====

  async importCsv() {
    const native = (window as any).native;
    if (!native?.openCsv) return;
    const res = await native.openCsv();
    if (!res) return;
    const rows = parseCsv(res.text);
    this.wb = new Workbook();
    this.engine = new CalcEngine(this.wb);
    rows.forEach((cols, r) => cols.forEach((val, c) => { if (val !== "") this.engine.setCellRaw(r, c, val); }));
    this.engine.rebuild();
    this.grid.setSheet(this.wb.active);
    this.undoStack = [];
    this.redoStack = [];
    this.renderTabs();
    this.renderCharts();
    this.wb.dirty = false;
    document.title = this.titleText();
  }

  async exportCsv() {
    const native = (window as any).native;
    if (!native?.saveCsv) return;
    const sheet = this.wb.active;
    let maxR = 0;
    let maxC = 0;
    for (const k of sheet.cells.keys()) {
      const { row, col } = parseKey(k);
      if (row > maxR) maxR = row;
      if (col > maxC) maxC = col;
    }
    const rows: string[][] = [];
    for (let r = 0; r <= maxR; r++) {
      const line: string[] = [];
      for (let c = 0; c <= maxC; c++) {
        const cell = sheet.getCell(r, c);
        line.push(cell ? formatValue(cell.value, cell.format?.numFmt) : "");
      }
      rows.push(line);
    }
    await native.saveCsv(toCsv(rows), `${sheet.name}.csv`);
  }

  async exportPdf() {
    const native = (window as any).native;
    if (!native?.printPdf) return;
    await native.printPdf(`${this.wb.sheets[0].name}.pdf`);
  }

  private markDirty() {
    if (!this.wb.dirty) {
      this.wb.dirty = true;
      document.title = this.titleText();
    }
  }

  private titleText(): string {
    const name = this.wb.filePath ? this.wb.filePath.split("/").pop() : "Untitled";
    return `${this.wb.dirty ? "• " : ""}${name} — BoringSheets`;
  }
}

// Sort comparison: numbers numerically, blanks last, otherwise case-insensitive
// text. Errors sort after everything.
function compareForSort(a: CellValue, b: CellValue): number {
  const rank = (v: CellValue) => (v === null || v === "" ? 3 : v instanceof CellError ? 2 : typeof v === "number" || typeof v === "boolean" ? 0 : 1);
  const ra = rank(a);
  const rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (ra === 0) {
    const na = typeof a === "boolean" ? (a ? 1 : 0) : (a as number);
    const nb = typeof b === "boolean" ? (b ? 1 : 0) : (b as number);
    return na - nb;
  }
  const sa = String(a).toLowerCase();
  const sb = String(b).toLowerCase();
  return sa < sb ? -1 : sa > sb ? 1 : 0;
}

// Increase/decrease the decimal places of a number format, preserving any
// currency prefix, thousands separator and percent sign.
function adjustDecimals(fmt: string | undefined, delta: number): string | undefined {
  const f = fmt && fmt !== "General" ? fmt : "0";
  const isPct = f.includes("%");
  const hasThousands = f.includes("#,##0") || /[#0],[#0]/.test(f);
  const curDecimals = fmt && fmt !== "General" ? (f.split(".")[1]?.match(/0/g) || []).length : 0;
  const n = Math.max(0, Math.min(9, curDecimals + delta));
  const prefixMatch = /^("[^"]*"|[$£€])/.exec(f.trim());
  const prefix = prefixMatch ? prefixMatch[0] : "";
  const intPart = hasThousands ? "#,##0" : "0";
  const out = prefix + intPart + (n > 0 ? "." + "0".repeat(n) : "") + (isPct ? "%" : "");
  return out;
}

function replaceCaseInsensitive(haystack: string, find: string, repl: string, onlyFirst = false): string {
  if (find === "") return haystack;
  let out = "";
  let i = 0;
  const lowHay = haystack.toLowerCase();
  const lowFind = find.toLowerCase();
  while (i < haystack.length) {
    if (lowHay.startsWith(lowFind, i)) {
      out += repl;
      i += find.length;
      if (onlyFirst) {
        out += haystack.slice(i);
        return out;
      }
    } else {
      out += haystack[i++];
    }
  }
  return out;
}
