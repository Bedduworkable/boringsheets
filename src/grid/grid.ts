// Canvas-rendered spreadsheet grid with virtual scrolling, frozen panes,
// selection, keyboard navigation, in-cell editing, merged cells, drag-resize,
// a fill handle, and a right-click context menu. Rendering reads from the
// Sheet; structural changes go back through the host callbacks.

import { Sheet, HEADER_WIDTH, HEADER_HEIGHT, MergeRegion } from "../model/sheet.js";
import { colToLetter, formatA1, parseA1 } from "../engine/references.js";
import { formatValue } from "../engine/format.js";
import { wrapLines } from "../engine/textwrap.js";
import { CellError, CellFormat } from "../model/types.js";
import type { CondVisual } from "../engine/conditional.js";
import { FormulaAutocomplete } from "../ui/autocomplete.js";

export interface SelRange {
  r0: number;
  r1: number;
  c0: number;
  c1: number;
}

export interface GridHost {
  commit(row: number, col: number, raw: string): void;
  onSelectionChange(): void;
  onContextMenu(row: number, col: number, clientX: number, clientY: number): void;
  resizeCol(col: number, width: number, prevWidth?: number): void;
  resizeRow(row: number, height: number, prevHeight?: number): void;
  autofitCol(col: number): void;
  autofitRow(row: number): void;
  fillRange(src: SelRange, dest: SelRange): void;
  // Returns the allowed values if the cell has a list (dropdown) validation.
  listValidation(row: number, col: number): string[] | null;
  // Open the dropdown picker for a list-validated cell.
  onDropdown(row: number, col: number, clientX: number, clientY: number): void;
  // Conditional-formatting visual for a cell, or null.
  conditional(row: number, col: number): CondVisual | null;
  // The grid's zoom level changed (so the toolbar can show the %).
  onZoomChange(zoom: number): void;
  // A filter funnel in the header row was clicked.
  onFilterDropdown(col: number, clientX: number, clientY: number): void;
}

export interface CellPos {
  row: number;
  col: number;
}

interface ColSlot { col: number; x: number; w: number; frozen: boolean }
interface RowSlot { row: number; y: number; h: number; frozen: boolean }

const RESIZE_ZONE = 4;
const HANDLE_SIZE = 6;

// On commit, auto-append any missing closing parens (Excel does this) so
// "=SUM(B2:B5" becomes "=SUM(B2:B5)". Parens inside string literals are ignored.
function autoCloseParens(formula: string): string {
  if (!formula.startsWith("=")) return formula;
  let depth = 0;
  let inStr = false;
  for (let i = 0; i < formula.length; i++) {
    const ch = formula[i];
    if (ch === '"') {
      if (inStr && formula[i + 1] === '"') { i++; continue; } // escaped ""
      inStr = !inStr;
      continue;
    }
    if (inStr) continue;
    if (ch === "(") depth++;
    else if (ch === ")") depth = Math.max(0, depth - 1);
  }
  return depth > 0 ? formula + ")".repeat(depth) : formula;
}
// Range Finder colors (cycled per reference in a formula being edited).
const REF_COLORS = ["#1a73e8", "#d83b01", "#188038", "#a142f4", "#e37400", "#0b8043"];

export class Grid {
  private ctx: CanvasRenderingContext2D;
  private dpr = window.devicePixelRatio || 1;
  private viewW = 0;
  private viewH = 0;

  scrollX = 0;
  scrollY = 0;
  zoom = 1;

  active: CellPos = { row: 0, col: 0 };
  anchor: CellPos = { row: 0, col: 0 };

  private editing = false;
  private autocomplete: FormulaAutocomplete;
  // Point mode (Excel Range Finder): inserting cell references by clicking while
  // editing a formula. `pointing` is active during a click/drag; `refBoxes` are
  // the color-coded highlights for every reference in the formula.
  private pointing: { insStart: number; insEnd: number; startCell: CellPos; prefix: string } | null = null;
  // Range Finder: every reference in the formula being edited, color-coded.
  private refBoxes: { r0: number; c0: number; r1: number; c1: number; color: string }[] = [];
  private dragging = false;
  private resizing: { type: "col" | "row"; index: number; start: number; startSize: number } | null = null;
  private filling: { src: SelRange; dest: SelRange } | null = null;

  constructor(
    private canvas: HTMLCanvasElement,
    private container: HTMLElement,
    private editor: HTMLInputElement,
    private sheet: Sheet,
    private host: GridHost
  ) {
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("no 2d context");
    this.ctx = ctx;
    this.autocomplete = new FormulaAutocomplete(editor);
    this.attach();
    this.resize();
  }

  setSheet(sheet: Sheet) {
    this.sheet = sheet;
    this.active = { row: 0, col: 0 };
    this.anchor = { row: 0, col: 0 };
    this.scrollX = 0;
    this.scrollY = 0;
    this.render();
    this.host.onSelectionChange();
  }

  getSheet() {
    return this.sheet;
  }

  getZoom() {
    return this.zoom;
  }
  setZoom(z: number) {
    this.zoom = Math.max(0.5, Math.min(2.5, Math.round(z * 100) / 100));
    this.resize();
    this.host.onZoomChange(this.zoom);
  }
  zoomBy(delta: number) {
    this.setZoom(this.zoom + delta);
  }

  resetScroll() {
    this.scrollX = 0;
    this.scrollY = 0;
    this.render();
  }

  // --- freeze geometry ---
  private get fCols() {
    return this.sheet.frozenCols;
  }
  private get fRows() {
    return this.sheet.frozenRows;
  }
  private frozenW() {
    let w = 0;
    for (let c = 0; c < this.fCols; c++) w += this.sheet.colWidth(c);
    return w;
  }
  private frozenH() {
    let h = 0;
    for (let r = 0; r < this.fRows; r++) h += this.sheet.rowHeight(r);
    return h;
  }
  private originX() {
    return HEADER_WIDTH + this.frozenW();
  }
  private originY() {
    return HEADER_HEIGHT + this.frozenH();
  }
  // scroll-space x of a scrolling column (col >= frozenCols), measured from the
  // first scrolling column.
  private scrollColOffset(col: number) {
    return this.sheet.colX(col) - this.sheet.colX(this.fCols);
  }
  private scrollRowOffset(row: number) {
    return this.sheet.rowY(row) - this.sheet.rowY(this.fRows);
  }

  // --- selection ---
  selRange(): SelRange {
    return this.expandToMerges({
      r0: Math.min(this.active.row, this.anchor.row),
      r1: Math.max(this.active.row, this.anchor.row),
      c0: Math.min(this.active.col, this.anchor.col),
      c1: Math.max(this.active.col, this.anchor.col),
    });
  }
  private expandToMerges(sel: SelRange): SelRange {
    let changed = true;
    while (changed) {
      changed = false;
      for (const m of this.sheet.merges) {
        if (sel.r0 <= m.r1 && sel.r1 >= m.r0 && sel.c0 <= m.c1 && sel.c1 >= m.c0) {
          if (m.r0 < sel.r0) { sel.r0 = m.r0; changed = true; }
          if (m.r1 > sel.r1) { sel.r1 = m.r1; changed = true; }
          if (m.c0 < sel.c0) { sel.c0 = m.c0; changed = true; }
          if (m.c1 > sel.c1) { sel.c1 = m.c1; changed = true; }
        }
      }
    }
    return sel;
  }
  private snapToAnchor(pos: CellPos): CellPos {
    const m = this.sheet.mergeAt(pos.row, pos.col);
    return m ? { row: m.r0, col: m.c0 } : pos;
  }
  setActive(row: number, col: number, extend = false) {
    this.active = this.snapToAnchor({ row: Math.max(0, row), col: Math.max(0, col) });
    if (!extend) this.anchor = { ...this.active };
    this.scrollIntoView();
    this.render();
    this.host.onSelectionChange();
  }

  // --- layout ---
  resize() {
    const rect = this.container.getBoundingClientRect();
    // Everything is laid out in logical (unzoomed) pixels; one transform scales
    // the whole canvas by dpr*zoom, so the visible logical area shrinks/grows
    // with zoom.
    this.viewW = rect.width / this.zoom;
    this.viewH = rect.height / this.zoom;
    this.canvas.width = Math.floor(rect.width * this.dpr);
    this.canvas.height = Math.floor(rect.height * this.dpr);
    this.canvas.style.width = `${rect.width}px`;
    this.canvas.style.height = `${rect.height}px`;
    this.ctx.setTransform(this.dpr * this.zoom, 0, 0, this.dpr * this.zoom, 0, 0);
    this.render();
  }

  // pixel x of a column's left edge (freeze-aware)
  private colLeft(col: number): number {
    if (col < this.fCols) return HEADER_WIDTH + this.sheet.colX(col);
    return this.originX() + this.scrollColOffset(col) - this.scrollX;
  }
  private rowTop(row: number): number {
    if (row < this.fRows) return HEADER_HEIGHT + this.sheet.rowY(row);
    return this.originY() + this.scrollRowOffset(row) - this.scrollY;
  }

  private firstScrollCol(): number {
    let c = this.fCols;
    while (this.scrollColOffset(c) + this.sheet.colWidth(c) <= this.scrollX) c++;
    return c;
  }
  private firstScrollRow(): number {
    let r = this.fRows;
    while (this.scrollRowOffset(r) + this.sheet.rowHeight(r) <= this.scrollY) r++;
    return r;
  }

  private colSlots(): ColSlot[] {
    const out: ColSlot[] = [];
    for (let c = 0; c < this.fCols; c++) out.push({ col: c, x: this.colLeft(c), w: this.sheet.colWidth(c), frozen: true });
    const start = this.firstScrollCol();
    let x = this.colLeft(start);
    for (let c = start; x < this.viewW && c <= this.sheet.maxCol + 50; c++) {
      const w = this.sheet.colWidth(c);
      out.push({ col: c, x, w, frozen: false });
      x += w;
    }
    return out;
  }
  private rowSlots(): RowSlot[] {
    const out: RowSlot[] = [];
    for (let r = 0; r < this.fRows; r++) out.push({ row: r, y: this.rowTop(r), h: this.sheet.rowHeight(r), frozen: true });
    const start = this.firstScrollRow();
    let y = this.rowTop(start);
    for (let r = start; y < this.viewH && r <= this.sheet.maxRow + 50; r++) {
      const h = this.sheet.rowHeight(r);
      out.push({ row: r, y, h, frozen: false });
      y += h;
    }
    return out;
  }

  // --- rendering ---
  render() {
    const ctx = this.ctx;
    const css = getComputedStyle(document.documentElement);
    const border = css.getPropertyValue("--border").trim() || "#d7d7db";
    const headerBg = css.getPropertyValue("--header-bg").trim() || "#f3f3f4";
    const headerActive = css.getPropertyValue("--header-active").trim() || "#d4e6f7";
    const selFill = css.getPropertyValue("--selection").trim() || "rgba(33,115,70,0.12)";
    const selBorder = css.getPropertyValue("--selection-border").trim() || "#217346";

    ctx.clearRect(0, 0, this.viewW, this.viewH);
    ctx.fillStyle = "#fff";
    ctx.fillRect(0, 0, this.viewW, this.viewH);

    const cols = this.colSlots();
    const rows = this.rowSlots();
    const sel = this.selRange();
    const ox = this.originX();
    const oy = this.originY();

    const fCols = cols.filter((c) => c.frozen);
    const sCols = cols.filter((c) => !c.frozen);
    const fRows = rows.filter((r) => r.frozen);
    const sRows = rows.filter((r) => !r.frozen);

    // Draw cell bodies + gridlines in four panes (main, top, left, corner),
    // each clipped so scrolling content can't bleed under the frozen bands.
    const drawPane = (cs: ColSlot[], rs: RowSlot[], cx: number, cy: number, cw: number, ch: number) => {
      if (cw <= 0 || ch <= 0 || !cs.length || !rs.length) return;
      ctx.save();
      ctx.beginPath();
      ctx.rect(cx, cy, cw, ch);
      ctx.clip();
      ctx.textBaseline = "middle";
      // Pass 1: backgrounds + selection tint.
      for (const rr of rs) {
        for (const cc of cs) {
          const merge = this.sheet.mergeAt(rr.row, cc.col);
          if (merge && !(merge.r0 === rr.row && merge.c0 === cc.col)) continue;
          const w = merge ? this.mergeWidth(merge) : cc.w;
          const h = merge ? this.mergeHeight(merge) : rr.h;
          const inSel = rr.row >= sel.r0 && rr.row <= sel.r1 && cc.col >= sel.c0 && cc.col <= sel.c1;
          this.drawCellBg(rr.row, cc.col, cc.x, rr.y, w, h, inSel, selFill);
        }
      }
      // Pass 2: gridlines (above backgrounds, below text).
      ctx.strokeStyle = border;
      ctx.lineWidth = 1;
      ctx.beginPath();
      for (const cc of cs) {
        ctx.moveTo(Math.round(cc.x) + 0.5, cy);
        ctx.lineTo(Math.round(cc.x) + 0.5, cy + ch);
        ctx.moveTo(Math.round(cc.x + cc.w) + 0.5, cy);
        ctx.lineTo(Math.round(cc.x + cc.w) + 0.5, cy + ch);
      }
      for (const rr of rs) {
        ctx.moveTo(cx, Math.round(rr.y) + 0.5);
        ctx.lineTo(cx + cw, Math.round(rr.y) + 0.5);
        ctx.moveTo(cx, Math.round(rr.y + rr.h) + 0.5);
        ctx.lineTo(cx + cw, Math.round(rr.y + rr.h) + 0.5);
      }
      ctx.stroke();
      // Pass 3: borders, note markers, then text — all above gridlines.
      for (const rr of rs) {
        for (const cc of cs) {
          const merge = this.sheet.mergeAt(rr.row, cc.col);
          if (merge && !(merge.r0 === rr.row && merge.c0 === cc.col)) continue;
          const w = merge ? this.mergeWidth(merge) : cc.w;
          const h = merge ? this.mergeHeight(merge) : rr.h;
          this.drawCellDecor(rr.row, cc.col, cc.x, rr.y, w, h);
          this.drawCellText(rr.row, cc.col, cc.x, rr.y, w, h, !merge);
        }
      }
      ctx.restore();
    };

    drawPane(sCols, sRows, ox, oy, this.viewW - ox, this.viewH - oy); // main
    drawPane(sCols, fRows, ox, HEADER_HEIGHT, this.viewW - ox, oy - HEADER_HEIGHT); // top
    drawPane(fCols, sRows, HEADER_WIDTH, oy, ox - HEADER_WIDTH, this.viewH - oy); // left
    drawPane(fCols, fRows, HEADER_WIDTH, HEADER_HEIGHT, ox - HEADER_WIDTH, oy - HEADER_HEIGHT); // corner

    // merged-region repaint (skip when it straddles a frozen boundary)
    for (const m of this.sheet.merges) {
      if ((this.fRows && m.r0 < this.fRows && m.r1 >= this.fRows) ||
          (this.fCols && m.c0 < this.fCols && m.c1 >= this.fCols)) continue;
      const x = this.colLeft(m.c0);
      const y = this.rowTop(m.r0);
      const w = this.mergeWidth(m);
      const h = this.mergeHeight(m);
      if (x > this.viewW || y > this.viewH || x + w < HEADER_WIDTH || y + h < HEADER_HEIGHT) continue;
      const inSel = m.r0 >= sel.r0 && m.r1 <= sel.r1 && m.c0 >= sel.c0 && m.c1 <= sel.c1;
      ctx.save();
      ctx.beginPath();
      ctx.rect(HEADER_WIDTH, HEADER_HEIGHT, this.viewW - HEADER_WIDTH, this.viewH - HEADER_HEIGHT);
      ctx.clip();
      ctx.fillStyle = "#fff";
      ctx.fillRect(x, y, w, h);
      this.drawCell(m.r0, m.c0, x, y, w, h, inSel, selFill);
      ctx.strokeStyle = border;
      ctx.lineWidth = 1;
      ctx.strokeRect(Math.round(x) + 0.5, Math.round(y) + 0.5, w, h);
      ctx.restore();
    }

    // selection outline
    const selBox = this.pixelBox(sel);
    if (selBox && selBox.w > 0 && selBox.h > 0) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(HEADER_WIDTH, HEADER_HEIGHT, this.viewW - HEADER_WIDTH, this.viewH - HEADER_HEIGHT);
      ctx.clip();
      ctx.strokeStyle = selBorder;
      ctx.lineWidth = 2;
      ctx.strokeRect(selBox.x + 1, selBox.y + 1, selBox.w - 2, selBox.h - 2);
      if (!this.filling) {
        ctx.fillStyle = selBorder;
        ctx.fillRect(selBox.x + selBox.w - HANDLE_SIZE / 2 - 1, selBox.y + selBox.h - HANDLE_SIZE / 2 - 1, HANDLE_SIZE, HANDLE_SIZE);
      }
      if (this.filling) {
        const box = this.pixelBox(this.filling.dest);
        if (box) {
          ctx.strokeStyle = selBorder;
          ctx.setLineDash([4, 2]);
          ctx.lineWidth = 1.5;
          ctx.strokeRect(box.x + 1, box.y + 1, box.w - 2, box.h - 2);
          ctx.setLineDash([]);
        }
      }
      ctx.restore();
    }

    // Range Finder: highlight every reference in the formula being edited.
    if (this.refBoxes.length) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(HEADER_WIDTH, HEADER_HEIGHT, this.viewW - HEADER_WIDTH, this.viewH - HEADER_HEIGHT);
      ctx.clip();
      for (const b of this.refBoxes) {
        const box = this.pixelBox(b);
        if (!box || box.w <= 0 || box.h <= 0) continue;
        if (box.x > this.viewW || box.y > this.viewH || box.x + box.w < HEADER_WIDTH || box.y + box.h < HEADER_HEIGHT) continue;
        ctx.strokeStyle = b.color;
        ctx.setLineDash([4, 3]);
        ctx.lineWidth = 1.5;
        ctx.strokeRect(box.x + 1, box.y + 1, box.w - 2, box.h - 2);
        ctx.setLineDash([]);
        ctx.fillStyle = b.color + "1f"; // ~12% alpha
        ctx.fillRect(box.x + 1, box.y + 1, box.w - 2, box.h - 2);
      }
      ctx.restore();
    }

    // headers
    ctx.fillStyle = headerBg;
    ctx.fillRect(0, 0, this.viewW, HEADER_HEIGHT);
    ctx.fillRect(0, 0, HEADER_WIDTH, this.viewH);
    ctx.font = "12px -apple-system, Segoe UI, sans-serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle"; // vertically center header text (was left at default)
    for (const cc of cols) {
      if (cc.col >= sel.c0 && cc.col <= sel.c1) {
        ctx.fillStyle = headerActive;
        ctx.fillRect(cc.x, 0, cc.w, HEADER_HEIGHT);
      }
      ctx.fillStyle = "#444";
      ctx.fillText(colToLetter(cc.col), cc.x + cc.w / 2, HEADER_HEIGHT / 2 + 1);
    }
    for (const rr of rows) {
      if (rr.row >= sel.r0 && rr.row <= sel.r1) {
        ctx.fillStyle = headerActive;
        ctx.fillRect(0, rr.y, HEADER_WIDTH, rr.h);
      }
      ctx.fillStyle = "#444";
      ctx.fillText(String(rr.row + 1), HEADER_WIDTH / 2, rr.y + rr.h / 2);
    }
    // header separators between each column letter / row number
    ctx.strokeStyle = "#c4c4c8";
    ctx.lineWidth = 1;
    ctx.beginPath();
    for (const cc of cols) {
      ctx.moveTo(Math.round(cc.x + cc.w) + 0.5, 0);
      ctx.lineTo(Math.round(cc.x + cc.w) + 0.5, HEADER_HEIGHT);
    }
    for (const rr of rows) {
      ctx.moveTo(0, Math.round(rr.y + rr.h) + 0.5);
      ctx.lineTo(HEADER_WIDTH, Math.round(rr.y + rr.h) + 0.5);
    }
    ctx.stroke();
    // stronger divider between the headers and the grid
    ctx.strokeStyle = "#9aa0a6";
    ctx.beginPath();
    ctx.moveTo(0, HEADER_HEIGHT + 0.5);
    ctx.lineTo(this.viewW, HEADER_HEIGHT + 0.5);
    ctx.moveTo(HEADER_WIDTH + 0.5, 0);
    ctx.lineTo(HEADER_WIDTH + 0.5, this.viewH);
    ctx.stroke();

    // filter funnel buttons in the header row of an active filter
    this.drawFilterButtons();

    // data-validation dropdown arrow on the active cell
    this.drawDropdownArrow();

    // freeze divider lines
    if (this.fCols || this.fRows) {
      ctx.strokeStyle = "#9aa0a6";
      ctx.lineWidth = 1.5;
      ctx.beginPath();
      if (this.fCols) {
        ctx.moveTo(ox + 0.5, 0);
        ctx.lineTo(ox + 0.5, this.viewH);
      }
      if (this.fRows) {
        ctx.moveTo(0, oy + 0.5);
        ctx.lineTo(this.viewW, oy + 0.5);
      }
      ctx.stroke();
    }
  }

  // Draw a cell's borders and note marker (above gridlines, below text).
  private drawCellDecor(row: number, col: number, x: number, y: number, w: number, h: number) {
    const cell = this.sheet.getCell(row, col);
    if (!cell) return;
    const ctx = this.ctx;
    const b = cell.format?.border;
    if (b && (b.top || b.bottom || b.left || b.right)) {
      ctx.strokeStyle = b.color || "#000000";
      ctx.lineWidth = 1;
      ctx.beginPath();
      if (b.top) { ctx.moveTo(x, Math.round(y) + 0.5); ctx.lineTo(x + w, Math.round(y) + 0.5); }
      if (b.bottom) { ctx.moveTo(x, Math.round(y + h) - 0.5); ctx.lineTo(x + w, Math.round(y + h) - 0.5); }
      if (b.left) { ctx.moveTo(Math.round(x) + 0.5, y); ctx.lineTo(Math.round(x) + 0.5, y + h); }
      if (b.right) { ctx.moveTo(Math.round(x + w) - 0.5, y); ctx.lineTo(Math.round(x + w) - 0.5, y + h); }
      ctx.stroke();
    }
    if (cell.note) {
      ctx.fillStyle = "#d83b01";
      ctx.beginPath();
      ctx.moveTo(x + w - 7, y);
      ctx.lineTo(x + w, y);
      ctx.lineTo(x + w, y + 7);
      ctx.closePath();
      ctx.fill();
    }
  }

  private isEmptyCell(row: number, col: number): boolean {
    const c = this.sheet.getCell(row, col);
    return !c || c.value === null || c.value === "";
  }

  // Composite draw (background then text) — used for merged cells (no spill).
  private drawCell(row: number, col: number, x: number, y: number, w: number, h: number, inSel: boolean, selFill: string) {
    this.drawCellBg(row, col, x, y, w, h, inSel, selFill);
    this.drawCellText(row, col, x, y, w, h, false);
  }

  // Backgrounds: conditional fill, cell fill, data bar, selection tint.
  private drawCellBg(row: number, col: number, x: number, y: number, w: number, h: number, inSel: boolean, selFill: string) {
    const ctx = this.ctx;
    const cell = this.sheet.getCell(row, col);
    const cv = this.host.conditional(row, col);
    const bg = cv?.fillColor ?? cv?.format?.bg ?? cell?.format?.bg;
    if (bg) {
      ctx.fillStyle = bg;
      ctx.fillRect(x, y, w, h);
    }
    if (cv?.dataBar && cv.dataBar.fraction > 0) {
      const barW = Math.max(0, Math.min(1, cv.dataBar.fraction)) * (w - 4);
      ctx.fillStyle = cv.dataBar.color;
      ctx.globalAlpha = 0.55;
      ctx.fillRect(x + 2, y + h - 6, barW, 4);
      ctx.globalAlpha = 1;
    }
    if (inSel) {
      ctx.fillStyle = selFill;
      ctx.fillRect(x, y, w, h);
    }
  }

  // Cell text. When `allowOverflow` and a string is wider than its cell, the
  // clip is extended over adjacent EMPTY cells (Excel-style spill).
  private drawCellText(row: number, col: number, x: number, y: number, w: number, h: number, allowOverflow: boolean) {
    const ctx = this.ctx;
    const cell = this.sheet.getCell(row, col);
    if (!cell || cell.value === null) return;
    const cv = this.host.conditional(row, col);
    const fmt: CellFormat | undefined = cv?.format ? { ...cell.format, ...cv.format } : cell.format;
    const text = formatValue(cell.value, fmt?.numFmt);
    if (text === "") return;
    const isErr = cell.value instanceof CellError;
    const isNum = typeof cell.value === "number";
    ctx.font = `${fmt?.italic ? "italic " : ""}${fmt?.bold ? "700 " : "400 "}${fmt?.fontSize || 13}px ${
      fmt?.fontFamily || "-apple-system, Segoe UI, sans-serif"
    }`;
    ctx.fillStyle = isErr ? "#c0392b" : fmt?.color || "#1b1b1b";
    const align = fmt?.align || (isNum ? "right" : "left");

    // Wrapped text: render multiple lines clipped to the cell (no spill).
    if (fmt?.wrap) {
      const lines = wrapLines((s) => ctx.measureText(s).width, text, w - 8);
      const lineH = (fmt.fontSize || 13) * 1.35;
      ctx.save();
      ctx.beginPath();
      ctx.rect(x, y, w, h);
      ctx.clip();
      let ty = y + 4 + lineH / 2;
      for (const ln of lines) {
        if (align === "right") { ctx.textAlign = "right"; ctx.fillText(ln, x + w - 4, ty); }
        else if (align === "center") { ctx.textAlign = "center"; ctx.fillText(ln, x + w / 2, ty); }
        else { ctx.textAlign = "left"; ctx.fillText(ln, x + 4, ty); }
        ty += lineH;
      }
      ctx.restore();
      return;
    }

    // Excel-style overflow into empty neighbours (text values only).
    let clipX = x;
    let clipW = w;
    if (allowOverflow && typeof cell.value === "string" && !isErr) {
      const need = ctx.measureText(text).width + 8;
      if (need > w) {
        const maxCol = this.sheet.maxCol + 50;
        if (align === "left") {
          let c = col + 1;
          while (clipW < need && c <= maxCol && this.isEmptyCell(row, c)) clipW += this.sheet.colWidth(c++);
        } else if (align === "right") {
          let c = col - 1;
          while (clipW < need && c >= 0 && this.isEmptyCell(row, c)) {
            const cw = this.sheet.colWidth(c--);
            clipW += cw;
            clipX -= cw;
          }
        } else {
          let cr = col + 1;
          let cl = col - 1;
          while (clipW < need && cr <= maxCol && this.isEmptyCell(row, cr)) clipW += this.sheet.colWidth(cr++);
          while (clipW < need && cl >= 0 && this.isEmptyCell(row, cl)) {
            const cw = this.sheet.colWidth(cl--);
            clipW += cw;
            clipX -= cw;
          }
        }
      }
    }

    ctx.save();
    ctx.beginPath();
    ctx.rect(clipX, y, clipW, h);
    ctx.clip();
    const cy = y + h / 2;
    if (align === "right") {
      ctx.textAlign = "right";
      ctx.fillText(text, x + w - 4, cy);
    } else if (align === "center") {
      ctx.textAlign = "center";
      ctx.fillText(text, x + w / 2, cy);
    } else {
      ctx.textAlign = "left";
      ctx.fillText(text, x + 4, cy);
    }
    if (fmt?.underline || fmt?.strike) {
      const m = ctx.measureText(text);
      const startX = align === "right" ? x + w - 4 - m.width : align === "center" ? x + w / 2 - m.width / 2 : x + 4;
      ctx.strokeStyle = ctx.fillStyle as string;
      ctx.lineWidth = 1;
      if (fmt.underline) {
        ctx.beginPath();
        ctx.moveTo(startX, cy + (fmt.fontSize || 13) / 2 - 1);
        ctx.lineTo(startX + m.width, cy + (fmt.fontSize || 13) / 2 - 1);
        ctx.stroke();
      }
      if (fmt.strike) {
        ctx.beginPath();
        ctx.moveTo(startX, cy);
        ctx.lineTo(startX + m.width, cy);
        ctx.stroke();
      }
    }
    ctx.restore();
  }

  private mergeWidth(m: MergeRegion) {
    let w = 0;
    for (let c = m.c0; c <= m.c1; c++) w += this.sheet.colWidth(c);
    return w;
  }
  private mergeHeight(m: MergeRegion) {
    let h = 0;
    for (let r = m.r0; r <= m.r1; r++) h += this.sheet.rowHeight(r);
    return h;
  }

  private pixelBox(sel: SelRange) {
    const x = this.colLeft(sel.c0);
    const y = this.rowTop(sel.r0);
    let w = 0;
    for (let c = sel.c0; c <= sel.c1; c++) w += this.sheet.colWidth(c);
    let h = 0;
    for (let r = sel.r0; r <= sel.r1; r++) h += this.sheet.rowHeight(r);
    return { x, y, w, h };
  }

  // Pixel rect of the validation-dropdown button on the active cell, or null.
  private dropdownRect(): { x: number; y: number; w: number; h: number } | null {
    if (this.editing) return null;
    if (!this.host.listValidation(this.active.row, this.active.col)) return null;
    const m = this.sheet.mergeAt(this.active.row, this.active.col);
    const c0 = m ? m.c0 : this.active.col;
    const r0 = m ? m.r0 : this.active.row;
    const x = this.colLeft(c0);
    const y = this.rowTop(r0);
    const w = m ? this.mergeWidth(m) : this.sheet.colWidth(c0);
    const h = m ? this.mergeHeight(m) : this.sheet.rowHeight(r0);
    if (x + w < HEADER_WIDTH || y + h < HEADER_HEIGHT) return null;
    const bw = 17;
    return { x: x + w - bw, y, w: bw, h };
  }

  // Rect of column c's filter funnel (header-row cell), or null if off-screen.
  private filterButtonRect(c: number): { x: number; y: number; w: number; h: number } | null {
    const f = this.sheet.filter;
    if (!f || c < f.range.c0 || c > f.range.c1) return null;
    const r0 = f.range.r0;
    const cw = this.sheet.colWidth(c);
    const rh = this.sheet.rowHeight(r0);
    if (cw <= 0 || rh <= 0) return null;
    const cx = this.colLeft(c);
    const cy = this.rowTop(r0);
    if (cx + cw < HEADER_WIDTH || cy + rh < HEADER_HEIGHT || cx > this.viewW || cy > this.viewH) return null;
    const bw = 17;
    return { x: cx + cw - bw - 1, y: cy + (rh - 16) / 2, w: bw, h: 16 };
  }

  private drawFilterButtons() {
    const f = this.sheet.filter;
    if (!f) return;
    const ctx = this.ctx;
    ctx.save();
    ctx.beginPath();
    ctx.rect(HEADER_WIDTH, HEADER_HEIGHT, this.viewW - HEADER_WIDTH, this.viewH - HEADER_HEIGHT);
    ctx.clip();
    for (let c = f.range.c0; c <= f.range.c1; c++) {
      const r = this.filterButtonRect(c);
      if (!r) continue;
      const active = (f.cols[c]?.length ?? 0) > 0;
      // button background
      ctx.fillStyle = active ? "#188038" : "#f1f3f4";
      ctx.strokeStyle = active ? "#188038" : "#bdc1c6";
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.rect(r.x + 0.5, r.y + 0.5, r.w - 1, r.h - 1);
      ctx.fill();
      ctx.stroke();
      // funnel glyph
      ctx.fillStyle = active ? "#fff" : "#5f6368";
      const fx = r.x + (r.w - 9) / 2;
      const fy = r.y + 4;
      ctx.beginPath();
      ctx.moveTo(fx, fy);
      ctx.lineTo(fx + 9, fy);
      ctx.lineTo(fx + 5.5, fy + 4);
      ctx.lineTo(fx + 5.5, fy + 8);
      ctx.lineTo(fx + 3.5, fy + 6.5);
      ctx.lineTo(fx + 3.5, fy + 4);
      ctx.closePath();
      ctx.fill();
    }
    ctx.restore();
  }

  private drawDropdownArrow() {
    const rect = this.dropdownRect();
    if (!rect) return;
    const ctx = this.ctx;
    ctx.save();
    ctx.beginPath();
    ctx.rect(HEADER_WIDTH, HEADER_HEIGHT, this.viewW - HEADER_WIDTH, this.viewH - HEADER_HEIGHT);
    ctx.clip();
    ctx.fillStyle = "#e8e8ea";
    ctx.fillRect(rect.x, rect.y, rect.w, rect.h);
    ctx.strokeStyle = "#bcbcc0";
    ctx.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.w - 1, rect.h - 1);
    ctx.fillStyle = "#555";
    ctx.beginPath();
    const cx = rect.x + rect.w / 2;
    const cy = rect.y + rect.h / 2;
    ctx.moveTo(cx - 4, cy - 2);
    ctx.lineTo(cx + 4, cy - 2);
    ctx.lineTo(cx, cy + 3);
    ctx.closePath();
    ctx.fill();
    ctx.restore();
  }

  // --- hit testing (freeze-aware) ---
  private hitTest(px: number, py: number): CellPos | null {
    if (px < HEADER_WIDTH || py < HEADER_HEIGHT) return null;
    const col = this.colAt(px);
    const row = this.rowAt(py);
    return { row, col };
  }
  private colAt(px: number): number {
    const ox = this.originX();
    if (px < ox && this.fCols) {
      let x = HEADER_WIDTH;
      let c = 0;
      while (c < this.fCols - 1 && x + this.sheet.colWidth(c) <= px) { x += this.sheet.colWidth(c); c++; }
      return c;
    }
    let target = this.scrollX + (px - ox) + this.sheet.colX(this.fCols);
    let x = this.sheet.colX(this.fCols);
    let c = this.fCols;
    while (x + this.sheet.colWidth(c) <= target) { x += this.sheet.colWidth(c); c++; }
    void target;
    return c;
  }
  private rowAt(py: number): number {
    const oy = this.originY();
    if (py < oy && this.fRows) {
      let y = HEADER_HEIGHT;
      let r = 0;
      while (r < this.fRows - 1 && y + this.sheet.rowHeight(r) <= py) { y += this.sheet.rowHeight(r); r++; }
      return r;
    }
    const target = this.scrollY + (py - oy) + this.sheet.rowY(this.fRows);
    let y = this.sheet.rowY(this.fRows);
    let r = this.fRows;
    while (y + this.sheet.rowHeight(r) <= target) { y += this.sheet.rowHeight(r); r++; }
    return r;
  }

  private resizeHit(px: number, py: number): { type: "col" | "row"; index: number } | null {
    if (py < HEADER_HEIGHT && px >= HEADER_WIDTH) {
      for (const cc of this.colSlots()) if (Math.abs(px - (cc.x + cc.w)) <= RESIZE_ZONE) return { type: "col", index: cc.col };
    }
    if (px < HEADER_WIDTH && py >= HEADER_HEIGHT) {
      for (const rr of this.rowSlots()) if (Math.abs(py - (rr.y + rr.h)) <= RESIZE_ZONE) return { type: "row", index: rr.row };
    }
    return null;
  }

  private overFillHandle(px: number, py: number): boolean {
    const box = this.pixelBox(this.selRange());
    return Math.abs(px - (box.x + box.w)) <= HANDLE_SIZE && Math.abs(py - (box.y + box.h)) <= HANDLE_SIZE;
  }

  // --- scrolling ---
  private maxScrollX() {
    const total = this.sheet.colX(this.sheet.maxCol + 30) - this.sheet.colX(this.fCols);
    return Math.max(0, total - (this.viewW - this.originX()));
  }
  private maxScrollY() {
    const total = this.sheet.rowY(this.sheet.maxRow + 30) - this.sheet.rowY(this.fRows);
    return Math.max(0, total - (this.viewH - this.originY()));
  }
  scrollBy(dx: number, dy: number) {
    this.scrollX = Math.max(0, Math.min(this.maxScrollX(), this.scrollX + dx));
    this.scrollY = Math.max(0, Math.min(this.maxScrollY(), this.scrollY + dy));
    this.render();
  }
  private scrollIntoView() {
    if (this.active.col >= this.fCols) {
      const off = this.scrollColOffset(this.active.col);
      const w = this.sheet.colWidth(this.active.col);
      const avail = this.viewW - this.originX();
      if (off < this.scrollX) this.scrollX = off;
      else if (off + w > this.scrollX + avail) this.scrollX = off + w - avail;
    }
    if (this.active.row >= this.fRows) {
      const off = this.scrollRowOffset(this.active.row);
      const h = this.sheet.rowHeight(this.active.row);
      const avail = this.viewH - this.originY();
      if (off < this.scrollY) this.scrollY = off;
      else if (off + h > this.scrollY + avail) this.scrollY = off + h - avail;
    }
  }

  // --- editing ---
  isEditing() {
    return this.editing;
  }
  isSelecting() {
    return this.dragging;
  }
  beginEdit(initial?: string) {
    const { row, col } = this.active;
    this.editing = true;
    const cell = this.sheet.getCell(row, col);
    const startText = initial !== undefined ? initial : cell?.raw ?? "";
    this.editor.value = startText;
    this.positionEditor(row, col);
    this.editor.style.display = "block";
    this.editor.focus();
    if (initial === undefined) this.editor.select();
    else this.editor.setSelectionRange(startText.length, startText.length);
    this.updateRefBoxes();
  }
  // --- Point mode (insert references by clicking while editing a formula) ---

  // Try to start a reference insertion at the clicked cell. Returns false if the
  // cursor isn't in a position that accepts a reference (caller commits instead).
  private tryPointStart(hit: CellPos, additive: boolean): boolean {
    const value = this.editor.value;
    if (!value.startsWith("=")) return false;
    // Clicking the cell you're editing would be a self-reference (#CIRCULAR!);
    // ignore it (consume the click but insert nothing).
    if (hit.row === this.active.row && hit.col === this.active.col) return true;
    const cursor = this.editor.selectionStart ?? value.length;
    const before = value.slice(0, cursor);
    // a reference already sits just before the cursor
    const refRe = /(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)$/;
    const m = refRe.exec(before);
    let insStart: number;
    let prefix = "";
    if (additive && before.length > 1 && /[)"\dA-Za-z]$/.test(before)) {
      // ⌘/Ctrl-click: append a NEW comma-separated reference (multi-select)
      insStart = cursor;
      prefix = ",";
    } else if (m) {
      insStart = m.index; // re-aim: replace the reference before the cursor
    } else if (/[=(,+\-*/^&<>%: ]\s*$/.test(before)) {
      insStart = cursor; // insert at a reference-accepting boundary
    } else {
      return false;
    }
    this.pointing = { insStart, insEnd: cursor, startCell: hit, prefix };
    this.applyPointRef(hit, hit);
    return true;
  }

  private applyPointRef(a: CellPos, b: CellPos) {
    if (!this.pointing) return;
    const r0 = Math.min(a.row, b.row);
    const r1 = Math.max(a.row, b.row);
    const c0 = Math.min(a.col, b.col);
    const c1 = Math.max(a.col, b.col);
    const ref = r0 === r1 && c0 === c1 ? formatA1(r0, c0) : `${formatA1(r0, c0)}:${formatA1(r1, c1)}`;
    const text = this.pointing.prefix + ref;
    const v = this.editor.value;
    const next = v.slice(0, this.pointing.insStart) + text + v.slice(this.pointing.insEnd);
    this.editor.value = next;
    this.pointing.insEnd = this.pointing.insStart + text.length;
    this.editor.setSelectionRange(this.pointing.insEnd, this.pointing.insEnd);
    this.updateRefBoxes();
  }

  // Parse the formula being edited and highlight every cell reference (the
  // Excel "Range Finder"), each in its own color.
  updateRefBoxes() {
    const v = this.editor.value;
    if (!this.editing || !v.startsWith("=")) {
      if (this.refBoxes.length) {
        this.refBoxes = [];
        this.render();
      }
      return;
    }
    const re = /(?<![A-Za-z0-9_$:])(\$?[A-Za-z]{1,3}\$?\d+)(?::(\$?[A-Za-z]{1,3}\$?\d+))?(?![A-Za-z0-9_(])/g;
    const boxes: { r0: number; c0: number; r1: number; c1: number; color: string }[] = [];
    let m: RegExpExecArray | null;
    let i = 0;
    while ((m = re.exec(v)) !== null) {
      const a = parseA1(m[1]);
      const b = m[2] ? parseA1(m[2]) : a;
      if (!a || !b) continue;
      boxes.push({
        r0: Math.min(a.row, b.row),
        c0: Math.min(a.col, b.col),
        r1: Math.max(a.row, b.row),
        c1: Math.max(a.col, b.col),
        color: REF_COLORS[i % REF_COLORS.length],
      });
      i++;
    }
    this.refBoxes = boxes;
    this.render();
  }

  private positionEditor(row: number, col: number) {
    const merge = this.sheet.mergeAt(row, col);
    const x = this.colLeft(merge ? merge.c0 : col);
    const y = this.rowTop(merge ? merge.r0 : row);
    const z = this.zoom;
    this.editor.style.left = `${x * z}px`;
    this.editor.style.top = `${y * z}px`;
    // mirror the cell's font so editing previews the real size/family/style
    const fmt = this.sheet.getFormat(merge ? merge.r0 : row, merge ? merge.c0 : col);
    this.editor.style.fontSize = `${(fmt?.fontSize || 13) * z}px`;
    this.editor.style.fontFamily = fmt?.fontFamily || "-apple-system, Segoe UI, sans-serif";
    this.editor.style.fontWeight = fmt?.bold ? "700" : "400";
    this.editor.style.fontStyle = fmt?.italic ? "italic" : "normal";
    this.autosizeEditor();
  }

  // Grow the edit box to fit its font (height) and its text (width), spilling
  // past the column like Google Sheets rather than clipping a big value.
  private autosizeEditor() {
    if (!this.editing) return;
    const z = this.zoom;
    const { row, col } = this.active;
    const merge = this.sheet.mergeAt(row, col);
    const r = merge ? merge.r0 : row;
    const c = merge ? merge.c0 : col;
    const baseW = merge ? this.mergeWidth(merge) : this.sheet.colWidth(c);
    const baseH = merge ? this.mergeHeight(merge) : this.sheet.rowHeight(r);
    const fmt = this.sheet.getFormat(r, c);
    const fontSize = fmt?.fontSize || 13;
    this.ctx.font = `${fmt?.italic ? "italic " : ""}${fmt?.bold ? "700 " : "400 "}${fontSize}px ${fmt?.fontFamily || "-apple-system, Segoe UI, sans-serif"}`;
    const textW = this.ctx.measureText(this.editor.value).width;
    const maxW = this.viewW / z - this.colLeft(c) - 4;
    const w = Math.min(maxW, Math.max(baseW, textW + 16));
    const h = Math.max(baseH, fontSize * 1.5 + 4);
    this.editor.style.width = `${w * z}px`;
    this.editor.style.height = `${h * z}px`;
  }

  // Re-apply the active cell's font to the editor (live preview while editing
  // when the user changes size/family/bold from the toolbar).
  refreshEditorStyle() {
    if (this.editing) this.positionEditor(this.active.row, this.active.col);
  }
  commitEdit(moveRow = 1, moveCol = 0) {
    if (!this.editing) return;
    this.autocomplete.close();
    this.pointing = null;
    this.refBoxes = [];
    const { row, col } = this.active;
    const value = autoCloseParens(this.editor.value);
    this.editing = false;
    this.editor.style.display = "none";
    this.host.commit(row, col, value);
    if (moveRow || moveCol) this.setActive(row + moveRow, col + moveCol);
    else {
      this.render();
      this.host.onSelectionChange();
    }
    this.container.focus();
  }
  cancelEdit() {
    if (!this.editing) return;
    this.autocomplete.close();
    this.pointing = null;
    this.refBoxes = [];
    this.editing = false;
    this.editor.style.display = "none";
    this.container.focus();
    this.render();
  }
  syncEditorFromFormulaBar(text: string) {
    if (!this.editing) this.beginEdit(text);
    else this.editor.value = text;
    this.updateRefBoxes();
  }

  // --- input wiring ---
  private attach() {
    this.container.tabIndex = 0;
    new ResizeObserver(() => this.resize()).observe(this.container);
    this.canvas.addEventListener("mousedown", (e) => this.onMouseDown(e));
    this.canvas.addEventListener("mousemove", (e) => this.onMouseMove(e));
    this.canvas.addEventListener("mousemove", (e) => this.updateNoteTip(e));
    this.canvas.addEventListener("mouseleave", () => this.hideNoteTip());
    window.addEventListener("mouseup", (e) => this.onMouseUp(e));
    this.canvas.addEventListener("dblclick", (e) => this.onDblClick(e));
    this.canvas.addEventListener("contextmenu", (e) => this.onContextMenu(e));
    this.container.addEventListener("wheel", (e) => {
      e.preventDefault();
      if (e.ctrlKey || e.metaKey) {
        this.zoomBy(e.deltaY < 0 ? 0.1 : -0.1);
      } else {
        this.scrollBy(e.deltaX / this.zoom, e.deltaY / this.zoom);
      }
    }, { passive: false });

    this.editor.addEventListener("input", () => {
      this.autocomplete.onInput();
      this.pointing = null; // typing exits the point-drag replace state
      this.updateRefBoxes();
      this.autosizeEditor(); // grow the box as the value gets longer
    });
    this.editor.addEventListener("keydown", (e) => {
      // For keys the editor consumes, stop propagation so the event doesn't ALSO
      // bubble to the grid handler — committing flips editing=false mid-event,
      // which would otherwise let the grid move the cursor a second time
      // (skipping a cell on Tab / a row on Enter).
      if (this.autocomplete.handleKeydown(e)) { e.stopPropagation(); return; }
      if (e.key === "Enter") { e.preventDefault(); e.stopPropagation(); this.commitEdit(e.shiftKey ? -1 : 1, 0); }
      else if (e.key === "Tab") { e.preventDefault(); e.stopPropagation(); this.commitEdit(0, e.shiftKey ? -1 : 1); }
      else if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); this.cancelEdit(); }
      else if (e.key === "F4") { e.preventDefault(); e.stopPropagation(); this.cycleReferenceInEditor(); }
      else if ((e.metaKey || e.ctrlKey) && (e.key === "a" || e.key === "A")) {
        // Select all the editor text (Ctrl+A too, not just Cmd+A on macOS).
        e.preventDefault();
        e.stopPropagation();
        this.editor.select();
      }
    });
    this.container.addEventListener("keydown", (e) => this.onKeyDown(e));
  }

  private localXY(e: MouseEvent) {
    const rect = this.canvas.getBoundingClientRect();
    // Convert screen pixels to logical (unzoomed) coordinates.
    return { px: (e.clientX - rect.left) / this.zoom, py: (e.clientY - rect.top) / this.zoom };
  }

  private noteTip: HTMLElement | null = null;
  private updateNoteTip(e: MouseEvent) {
    if (this.dragging || this.resizing || this.filling) return this.hideNoteTip();
    const { px, py } = this.localXY(e);
    if (px < HEADER_WIDTH || py < HEADER_HEIGHT) return this.hideNoteTip();
    const hit = this.hitTest(px, py);
    const note = hit ? this.sheet.getCell(hit.row, hit.col)?.note : null;
    if (!note) return this.hideNoteTip();
    if (!this.noteTip) {
      this.noteTip = document.createElement("div");
      this.noteTip.className = "note-tip";
      document.body.appendChild(this.noteTip);
    }
    this.noteTip.textContent = note;
    this.noteTip.style.left = `${e.clientX + 14}px`;
    this.noteTip.style.top = `${e.clientY + 14}px`;
    this.noteTip.style.display = "block";
  }
  private hideNoteTip() {
    if (this.noteTip) this.noteTip.style.display = "none";
  }

  private onMouseDown(e: MouseEvent) {
    if (e.button === 2) return;
    const { px, py } = this.localXY(e);

    // filter funnel button in a header row
    const f = this.sheet.filter;
    if (f && py >= HEADER_HEIGHT && px >= HEADER_WIDTH) {
      for (let c = f.range.c0; c <= f.range.c1; c++) {
        const r = this.filterButtonRect(c);
        if (r && px >= r.x && px <= r.x + r.w && py >= r.y && py <= r.y + r.h) {
          this.host.onFilterDropdown(c, e.clientX, e.clientY);
          return;
        }
      }
    }

    // validation dropdown button on the active cell
    const dd = this.dropdownRect();
    if (dd && px >= dd.x && px <= dd.x + dd.w && py >= dd.y && py <= dd.y + dd.h) {
      this.host.onDropdown(this.active.row, this.active.col, e.clientX, e.clientY);
      return;
    }

    const rz = this.resizeHit(px, py);
    if (rz) {
      this.resizing = {
        type: rz.type,
        index: rz.index,
        start: rz.type === "col" ? px : py,
        startSize: rz.type === "col" ? this.sheet.colWidth(rz.index) : this.sheet.rowHeight(rz.index),
      };
      return;
    }
    if (!this.editing && this.overFillHandle(px, py)) {
      const src = this.selRange();
      this.filling = { src, dest: { ...src } };
      return;
    }
    if (py < HEADER_HEIGHT && px >= HEADER_WIDTH) {
      const col = this.colAt(px);
      if (this.editing) this.commitEdit(0, 0);
      this.active = { row: 0, col };
      this.anchor = { row: this.sheet.maxRow, col };
      this.dragging = true;
      this.render();
      this.host.onSelectionChange();
      return;
    }
    if (px < HEADER_WIDTH && py >= HEADER_HEIGHT) {
      const row = this.rowAt(py);
      if (this.editing) this.commitEdit(0, 0);
      this.active = { row, col: 0 };
      this.anchor = { row, col: this.sheet.maxCol };
      this.dragging = true;
      this.render();
      this.host.onSelectionChange();
      return;
    }
    const hit = this.hitTest(px, py);
    if (!hit) return;
    // Point mode: while editing a formula, clicking a cell inserts its reference
    // instead of finishing the edit.
    if (this.editing && this.tryPointStart(hit, e.metaKey || e.ctrlKey)) {
      e.preventDefault(); // keep the editor focused
      return;
    }
    if (this.editing) this.commitEdit(0, 0);
    this.container.focus();
    this.dragging = true;
    this.setActive(hit.row, hit.col, e.shiftKey);
  }

  private onMouseMove(e: MouseEvent) {
    const { px, py } = this.localXY(e);
    if (this.pointing) {
      const hit = this.hitTest(Math.max(px, HEADER_WIDTH + 1), Math.max(py, HEADER_HEIGHT + 1));
      if (hit) this.applyPointRef(this.pointing.startCell, hit);
      return;
    }
    if (this.resizing) {
      if (this.resizing.type === "col") {
        const w = Math.max(20, this.resizing.startSize + (px - this.resizing.start));
        this.sheet.colWidths.set(this.resizing.index, w);
      } else {
        const h = Math.max(12, this.resizing.startSize + (py - this.resizing.start));
        this.sheet.rowHeights.set(this.resizing.index, h);
      }
      this.render();
      return;
    }
    if (this.filling) {
      const hit = this.hitTest(Math.max(px, HEADER_WIDTH + 1), Math.max(py, HEADER_HEIGHT + 1));
      if (hit) {
        const src = this.filling.src;
        const dRows = hit.row - src.r1;
        const dCols = hit.col - src.c1;
        const dest = { ...src };
        if (Math.abs(dRows) >= Math.abs(dCols)) dest.r1 = Math.max(src.r1, hit.row);
        else dest.c1 = Math.max(src.c1, hit.col);
        this.filling.dest = dest;
        this.render();
      }
      return;
    }
    const rz = this.resizeHit(px, py);
    this.canvas.style.cursor = rz ? (rz.type === "col" ? "col-resize" : "row-resize") : this.overFillHandle(px, py) ? "crosshair" : "cell";
    if (this.dragging) {
      const hit = this.hitTest(Math.max(px, HEADER_WIDTH + 1), Math.max(py, HEADER_HEIGHT + 1));
      if (hit && (hit.row !== this.active.row || hit.col !== this.active.col)) {
        this.active = this.snapToAnchor(hit);
        this.render();
        this.host.onSelectionChange();
      }
    }
  }

  private onMouseUp(_e: MouseEvent) {
    if (this.pointing) {
      // end the point drag but keep the marquee + the editor focused
      this.pointing = null;
      this.editor.focus();
      return;
    }
    if (this.resizing) {
      const { type, index, startSize } = this.resizing;
      const size = type === "col" ? this.sheet.colWidth(index) : this.sheet.rowHeight(index);
      this.resizing = null;
      // pass the pre-drag size so the host can build a correct undo step
      if (type === "col") this.host.resizeCol(index, size, startSize);
      else this.host.resizeRow(index, size, startSize);
      return;
    }
    if (this.filling) {
      const { src, dest } = this.filling;
      this.filling = null;
      if (dest.r1 !== src.r1 || dest.c1 !== src.c1) this.host.fillRange(src, dest);
      this.render();
      return;
    }
    const wasDragging = this.dragging;
    this.dragging = false;
    if (wasDragging) this.host.onSelectionChange(); // signal selection finalized
  }

  private onDblClick(e: MouseEvent) {
    const { px, py } = this.localXY(e);
    const rz = this.resizeHit(px, py);
    if (rz) {
      // double-click a header border → auto-fit that column/row to its content
      if (rz.type === "col") this.host.autofitCol(rz.index);
      else this.host.autofitRow(rz.index);
      return;
    }
    // double-click the fill handle → fill down to the adjacent column's extent
    if (!this.editing && this.overFillHandle(px, py)) {
      const src = this.selRange();
      const dest = this.fillExtent(src);
      if (dest.r1 > src.r1) this.host.fillRange(src, dest);
      return;
    }
    if (!this.editing) this.beginEdit();
  }

  private onContextMenu(e: MouseEvent) {
    e.preventDefault();
    const { px, py } = this.localXY(e);
    const hit = this.hitTest(px, py);
    if (!hit) return;
    const sel = this.selRange();
    const inside = hit.row >= sel.r0 && hit.row <= sel.r1 && hit.col >= sel.c0 && hit.col <= sel.c1;
    if (!inside) this.setActive(hit.row, hit.col);
    this.host.onContextMenu(this.active.row, this.active.col, e.clientX, e.clientY);
  }

  private onKeyDown(e: KeyboardEvent) {
    if (this.editing) return;
    const { row, col } = this.active;
    const ext = e.shiftKey;

    // Ctrl/Cmd navigation: jump to data edges, Home → A1, A → select used range.
    if (e.ctrlKey || e.metaKey) {
      switch (e.key) {
        case "ArrowUp": e.preventDefault(); this.jumpEdge(-1, 0, ext); return;
        case "ArrowDown": e.preventDefault(); this.jumpEdge(1, 0, ext); return;
        case "ArrowLeft": e.preventDefault(); this.jumpEdge(0, -1, ext); return;
        case "ArrowRight": e.preventDefault(); this.jumpEdge(0, 1, ext); return;
        case "Home": e.preventDefault(); this.setActive(0, 0, ext); return;
        case "a": case "A": e.preventDefault(); this.selectAll(); return;
      }
      return; // let copy/paste and other Cmd shortcuts bubble up
    }

    const merge = this.sheet.mergeAt(row, col);
    const downStep = merge ? merge.r1 - merge.r0 + 1 : 1;
    const rightStep = merge ? merge.c1 - merge.c0 + 1 : 1;
    switch (e.key) {
      case "ArrowUp": e.preventDefault(); this.setActive(row - 1, col, ext); return;
      case "ArrowDown": e.preventDefault(); this.setActive(row + downStep, col, ext); return;
      case "ArrowLeft": e.preventDefault(); this.setActive(row, col - 1, ext); return;
      case "ArrowRight": e.preventDefault(); this.setActive(row, col + rightStep, ext); return;
      case "Tab": e.preventDefault(); this.setActive(row, col + (ext ? -1 : rightStep)); return;
      case "Enter": e.preventDefault(); this.setActive(row + (ext ? -1 : downStep), col); return;
      case "PageDown": e.preventDefault(); this.setActive(row + 20, col, ext); return;
      case "PageUp": e.preventDefault(); this.setActive(Math.max(0, row - 20), col, ext); return;
      case "Home": e.preventDefault(); this.setActive(row, 0, ext); return;
      case "F2": e.preventDefault(); this.beginEdit(); return;
      case "Backspace":
      case "Delete": e.preventDefault(); this.clearSelection(); return;
      case "Escape": return;
    }
    if (e.key.length === 1 && !e.metaKey && !e.ctrlKey && !e.altKey) {
      this.beginEdit(e.key);
      e.preventDefault();
    }
  }

  clearSelection() {
    const sel = this.selRange();
    for (let r = sel.r0; r <= sel.r1; r++) {
      for (let c = sel.c0; c <= sel.c1; c++) {
        if (this.sheet.getRaw(r, c) !== "") this.host.commit(r, c, "");
      }
    }
    this.render();
    this.host.onSelectionChange();
  }

  private filled(r: number, c: number): boolean {
    const cl = this.sheet.getCell(r, c);
    return !!cl && cl.value !== null && cl.value !== "";
  }

  // Ctrl+Arrow: move to the edge of the current data region in a direction.
  private jumpEdge(dr: number, dc: number, extend: boolean) {
    const maxR = this.sheet.maxRow;
    const maxC = this.sheet.maxCol;
    const inB = (r: number, c: number) => r >= 0 && c >= 0 && r <= maxR && c <= maxC;
    let r = this.active.row;
    let c = this.active.col;
    if (!inB(r + dr, c + dc)) {
      this.setActive(r, c, extend);
      return;
    }
    if (this.filled(r, c) && this.filled(r + dr, c + dc)) {
      // travel to the last filled cell before a gap
      while (inB(r + dr, c + dc) && this.filled(r + dr, c + dc)) {
        r += dr;
        c += dc;
      }
    } else {
      // skip the gap to the next filled cell (or the edge)
      let nr = r + dr;
      let nc = c + dc;
      while (inB(nr, nc) && !this.filled(nr, nc)) {
        nr += dr;
        nc += dc;
      }
      if (inB(nr, nc)) {
        r = nr;
        c = nc;
      } else {
        r = Math.max(0, Math.min(maxR, nr));
        c = Math.max(0, Math.min(maxC, nc));
      }
    }
    this.setActive(r, c, extend);
  }

  // Ctrl+A: select the used range (bounding box of all non-empty cells).
  // Ctrl/Cmd+A: select the whole sheet (its logical extent).
  selectAll() {
    if (this.editing) return;
    this.active = { row: 0, col: 0 };
    this.anchor = { row: this.sheet.maxRow, col: this.sheet.maxCol };
    this.render();
    this.host.onSelectionChange();
  }

  // Cycle the cell reference at the editor cursor through the four $ states
  // (A1 → $A$1 → A$1 → $A1 → A1), like Excel's F4.
  cycleReferenceInEditor() {
    const el = this.editor;
    const text = el.value;
    const pos = el.selectionStart ?? text.length;
    const re = /(\$?)([A-Za-z]{1,3})(\$?)(\d+)/g;
    let m: RegExpExecArray | null;
    let target: RegExpExecArray | null = null;
    while ((m = re.exec(text)) !== null) {
      if (pos >= m.index && pos <= m.index + m[0].length) {
        target = m;
        break;
      }
      if (m.index > pos) break;
      target = m; // remember the last ref before the cursor
    }
    if (!target) return;
    const colAbs = target[1] === "$";
    const rowAbs = target[3] === "$";
    let nc: boolean;
    let nr: boolean;
    if (!colAbs && !rowAbs) { nc = true; nr = true; }
    else if (colAbs && rowAbs) { nc = false; nr = true; }
    else if (!colAbs && rowAbs) { nc = true; nr = false; }
    else { nc = false; nr = false; }
    const rep = (nc ? "$" : "") + target[2] + (nr ? "$" : "") + target[4];
    el.value = text.slice(0, target.index) + rep + text.slice(target.index + target[0].length);
    const caret = target.index + rep.length;
    el.setSelectionRange(caret, caret);
  }

  // The destination range when the fill handle is double-clicked: fill down to
  // match the length of data in the adjacent column.
  private fillExtent(src: SelRange): SelRange {
    const refCol = src.c0 > 0 ? src.c0 - 1 : src.c1 + 1;
    let last = src.r1;
    for (let r = src.r1 + 1; r <= this.sheet.maxRow + 1; r++) {
      if (!this.filled(r, refCol)) break;
      last = r;
    }
    return { ...src, r1: last };
  }
}
