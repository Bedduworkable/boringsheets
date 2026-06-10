import { Cell, CellFormat } from "./types.js";
import { cellKey, parseKey } from "../engine/references.js";
import { ConditionalRule } from "../engine/conditional.js";
import { ChartSpec } from "../charts/render.js";

// A chart placed on the sheet: its data source range, spec, and pixel geometry
// (position/size of the floating overlay, in grid client coordinates).
export interface SheetChart {
  id: string;
  spec: ChartSpec;
  dataRange: Rect;
  byRows: boolean; // series taken from rows vs columns
  x: number;
  y: number;
  w: number;
  h: number;
}

export const DEFAULT_COL_WIDTH = 88;
export const DEFAULT_ROW_HEIGHT = 22;
export const HEADER_WIDTH = 46; // row-number gutter
export const HEADER_HEIGHT = 24; // column-letter header

// A rectangular merged-cell region. The top-left (r0,c0) holds the value; the
// other covered cells render as part of it.
export interface MergeRegion {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

export interface Rect {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

// A Google-Sheets-style filter over a range. `cols[colIndex]` holds the display
// values that are UNCHECKED (hidden) for that column; columns combine with AND.
export interface SheetFilter {
  range: Rect;
  cols: Record<number, string[]>;
}

// A data-validation rule applied to a rectangular range.
export interface DataValidation {
  range: Rect;
  type: "list" | "number" | "textLength";
  // list:
  source?: string[];
  // number / textLength:
  operator?: "between" | "notBetween" | "gt" | "lt" | "gte" | "lte" | "eq" | "ne";
  min?: number;
  max?: number;
  allowBlank?: boolean;
  errorMessage?: string;
}

export interface SheetSnapshot {
  cells: [string, Cell][];
  colWidths: [number, number][];
  rowHeights: [number, number][];
  merges: MergeRegion[];
  validations: DataValidation[];
  conditionalRules: ConditionalRule[];
  charts: SheetChart[];
  filter: SheetFilter | null;
  maxRow: number;
  maxCol: number;
  frozenRows: number;
  frozenCols: number;
  hiddenRows: number[];
  hiddenCols: number[];
  manualRows: number[];
}

let SHEET_SEQ = 0;

// A single worksheet: sparse cell storage plus row/column sizing. Cells that
// were never touched simply don't exist in the map (keeps memory tiny).
export class Sheet {
  // Stable identity for the calc engine's cross-sheet dependency graph. Does
  // not change when the sheet is renamed.
  readonly id = ++SHEET_SEQ;
  name: string;
  cells = new Map<string, Cell>();
  colWidths = new Map<number, number>();
  rowHeights = new Map<number, number>();
  merges: MergeRegion[] = [];
  validations: DataValidation[] = [];
  conditionalRules: ConditionalRule[] = [];
  charts: SheetChart[] = [];
  filter: SheetFilter | null = null;
  frozenRows = 0;
  frozenCols = 0;
  // Rows hidden by an active filter (rendered with zero height).
  hiddenRows = new Set<number>();
  // Columns hidden by the user (rendered with zero width).
  hiddenCols = new Set<number>();
  // Rows whose height the user set manually (excluded from auto-fit).
  manualRows = new Set<number>();
  // Logical extent for scrollbars; grows as the user navigates/edits.
  maxRow = 200;
  maxCol = 52;

  constructor(name: string) {
    this.name = name;
  }

  getCell(row: number, col: number): Cell | undefined {
    return this.cells.get(cellKey(row, col));
  }

  getRaw(row: number, col: number): string {
    return this.cells.get(cellKey(row, col))?.raw ?? "";
  }

  // Create-or-return a mutable cell.
  ensureCell(row: number, col: number): Cell {
    const k = cellKey(row, col);
    let c = this.cells.get(k);
    if (!c) {
      c = { raw: "", value: null };
      this.cells.set(k, c);
    }
    if (row > this.maxRow) this.maxRow = row;
    if (col > this.maxCol) this.maxCol = col;
    return c;
  }

  deleteCellIfEmpty(row: number, col: number) {
    const k = cellKey(row, col);
    const c = this.cells.get(k);
    if (c && c.raw === "" && !c.format && !c.note) this.cells.delete(k);
  }

  getFormat(row: number, col: number): CellFormat | undefined {
    return this.cells.get(cellKey(row, col))?.format;
  }

  colWidth(col: number): number {
    if (this.hiddenCols.has(col)) return 0; // hidden column
    return this.colWidths.get(col) ?? DEFAULT_COL_WIDTH;
  }

  rowHeight(row: number): number {
    if (this.hiddenRows.has(row)) return 0; // filtered-out row
    return this.rowHeights.get(row) ?? DEFAULT_ROW_HEIGHT;
  }

  // Cumulative x of a column's left edge (excluding the header gutter).
  colX(col: number): number {
    let x = 0;
    for (let c = 0; c < col; c++) x += this.colWidth(c);
    return x;
  }

  rowY(row: number): number {
    let y = 0;
    for (let r = 0; r < row; r++) y += this.rowHeight(r);
    return y;
  }

  // --- merges ---
  mergeAt(row: number, col: number): MergeRegion | undefined {
    return this.merges.find(
      (m) => row >= m.r0 && row <= m.r1 && col >= m.c0 && col <= m.c1
    );
  }

  addMerge(region: MergeRegion) {
    // drop any existing merges overlapping the new region
    this.merges = this.merges.filter(
      (m) => !(region.r0 <= m.r1 && region.r1 >= m.r0 && region.c0 <= m.c1 && region.c1 >= m.c0)
    );
    if (region.r1 > region.r0 || region.c1 > region.c0) this.merges.push(region);
  }

  removeMergeAt(row: number, col: number) {
    this.merges = this.merges.filter((m) => !(row >= m.r0 && row <= m.r1 && col >= m.c0 && col <= m.c1));
  }

  // --- data validation ---
  validationAt(row: number, col: number): DataValidation | undefined {
    // last matching rule wins (most recently added)
    for (let i = this.validations.length - 1; i >= 0; i--) {
      const v = this.validations[i];
      if (row >= v.range.r0 && row <= v.range.r1 && col >= v.range.c0 && col <= v.range.c1) return v;
    }
    return undefined;
  }

  // --- snapshot / restore (used for undo of structural edits) ---
  snapshot(): SheetSnapshot {
    const cells: [string, Cell][] = [];
    for (const [k, c] of this.cells) {
      cells.push([k, { raw: c.raw, value: c.value, format: c.format ? { ...c.format } : undefined, note: c.note }]);
    }
    return {
      cells,
      colWidths: [...this.colWidths],
      rowHeights: [...this.rowHeights],
      merges: this.merges.map((m) => ({ ...m })),
      validations: this.validations.map((v) => ({ ...v, range: { ...v.range }, source: v.source ? [...v.source] : undefined })),
      conditionalRules: JSON.parse(JSON.stringify(this.conditionalRules)),
      charts: JSON.parse(JSON.stringify(this.charts)),
      filter: this.filter ? JSON.parse(JSON.stringify(this.filter)) : null,
      maxRow: this.maxRow,
      maxCol: this.maxCol,
      frozenRows: this.frozenRows,
      frozenCols: this.frozenCols,
      hiddenRows: [...this.hiddenRows],
      hiddenCols: [...this.hiddenCols],
      manualRows: [...this.manualRows],
    };
  }

  restore(snap: SheetSnapshot) {
    this.cells = new Map(snap.cells.map(([k, c]) => [k, { raw: c.raw, value: c.value, format: c.format ? { ...c.format } : undefined, note: c.note }]));
    this.colWidths = new Map(snap.colWidths);
    this.rowHeights = new Map(snap.rowHeights);
    this.merges = snap.merges.map((m) => ({ ...m }));
    this.validations = snap.validations.map((v) => ({ ...v, range: { ...v.range }, source: v.source ? [...v.source] : undefined }));
    this.conditionalRules = JSON.parse(JSON.stringify(snap.conditionalRules));
    this.charts = JSON.parse(JSON.stringify(snap.charts));
    this.filter = snap.filter ? JSON.parse(JSON.stringify(snap.filter)) : null;
    this.maxRow = snap.maxRow;
    this.maxCol = snap.maxCol;
    this.frozenRows = snap.frozenRows;
    this.frozenCols = snap.frozenCols;
    this.hiddenRows = new Set(snap.hiddenRows);
    this.hiddenCols = new Set(snap.hiddenCols);
    this.manualRows = new Set(snap.manualRows ?? []);
  }

  // --- structural edits (cells + sizes + merges only; the controller rewrites
  // formula references separately) ---

  insertRows(at: number, count: number) {
    this.remap((r, c) => (r >= at ? { row: r + count, col: c } : { row: r, col: c }));
    this.shiftSizes(this.rowHeights, at, count);
    this.adjustMergesRow(at, count);
    this.maxRow += count;
  }

  deleteRows(at: number, count: number) {
    this.remap((r, c) => {
      if (r >= at && r < at + count) return null;
      return { row: r >= at + count ? r - count : r, col: c };
    });
    this.unshiftSizes(this.rowHeights, at, count);
    this.adjustMergesRow(at, -count);
    this.maxRow = Math.max(0, this.maxRow - count);
  }

  insertCols(at: number, count: number) {
    this.remap((r, c) => (c >= at ? { row: r, col: c + count } : { row: r, col: c }));
    this.shiftSizes(this.colWidths, at, count);
    this.adjustMergesCol(at, count);
    this.maxCol += count;
  }

  deleteCols(at: number, count: number) {
    this.remap((r, c) => {
      if (c >= at && c < at + count) return null;
      return { row: r, col: c >= at + count ? c - count : c };
    });
    this.unshiftSizes(this.colWidths, at, count);
    this.adjustMergesCol(at, -count);
    this.maxCol = Math.max(0, this.maxCol - count);
  }

  private remap(fn: (row: number, col: number) => { row: number; col: number } | null) {
    const next = new Map<string, Cell>();
    for (const [k, cell] of this.cells) {
      const { row, col } = parseKey(k);
      const moved = fn(row, col);
      if (moved) next.set(cellKey(moved.row, moved.col), cell);
    }
    this.cells = next;
  }

  private shiftSizes(map: Map<number, number>, at: number, count: number) {
    const next = new Map<number, number>();
    for (const [idx, size] of map) next.set(idx >= at ? idx + count : idx, size);
    map.clear();
    for (const [k, v] of next) map.set(k, v);
  }

  private unshiftSizes(map: Map<number, number>, at: number, count: number) {
    const next = new Map<number, number>();
    for (const [idx, size] of map) {
      if (idx >= at && idx < at + count) continue;
      next.set(idx >= at + count ? idx - count : idx, size);
    }
    map.clear();
    for (const [k, v] of next) map.set(k, v);
  }

  private adjustMergesRow(at: number, delta: number) {
    const out: MergeRegion[] = [];
    for (const m of this.merges) {
      if (delta < 0) {
        const dc = -delta;
        // drop merges entirely inside the deleted band
        if (m.r0 >= at && m.r1 < at + dc) continue;
        const r0 = m.r0 >= at + dc ? m.r0 - dc : m.r0 >= at ? at : m.r0;
        const r1 = m.r1 >= at + dc ? m.r1 - dc : m.r1 >= at ? at - 1 : m.r1;
        if (r1 > r0 || m.c1 > m.c0) out.push({ ...m, r0, r1: Math.max(r0, r1) });
      } else {
        const r0 = m.r0 >= at ? m.r0 + delta : m.r0;
        const r1 = m.r1 >= at ? m.r1 + delta : m.r1;
        out.push({ ...m, r0, r1 });
      }
    }
    this.merges = out;
  }

  private adjustMergesCol(at: number, delta: number) {
    const out: MergeRegion[] = [];
    for (const m of this.merges) {
      if (delta < 0) {
        const dc = -delta;
        if (m.c0 >= at && m.c1 < at + dc) continue;
        const c0 = m.c0 >= at + dc ? m.c0 - dc : m.c0 >= at ? at : m.c0;
        const c1 = m.c1 >= at + dc ? m.c1 - dc : m.c1 >= at ? at - 1 : m.c1;
        if (c1 > c0 || m.r1 > m.r0) out.push({ ...m, c0, c1: Math.max(c0, c1) });
      } else {
        const c0 = m.c0 >= at ? m.c0 + delta : m.c0;
        const c1 = m.c1 >= at ? m.c1 + delta : m.c1;
        out.push({ ...m, c0, c1 });
      }
    }
    this.merges = out;
  }
}
