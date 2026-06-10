// Evaluates an AST node against a cell-resolution context. A node evaluates to
// either a scalar CellValue or a RangeValue (a 2D block of cells), which
// functions like SUM consume.

import { Node } from "./ast.js";
import { CellValue, CellError } from "../model/types.js";
import { CellRef } from "./references.js";
import { FUNCTIONS } from "./functions.js";

export interface RangeValue {
  range: true;
  r0: number;
  c0: number;
  r1: number;
  c1: number;
  values: CellValue[][]; // [row][col]
}

export type EvalResult = CellValue | RangeValue;

export function isRange(v: EvalResult): v is RangeValue {
  return typeof v === "object" && v !== null && (v as RangeValue).range === true;
}

export interface EvalContext {
  // Resolve a single cell's value. `sheet` is the qualifier as written in the
  // formula (e.g. "Sheet2"); undefined means the formula's own home sheet.
  getCellValue(sheet: string | undefined, row: number, col: number): CellValue;
  // Resolve a named range to its values (or null if the name is undefined).
  resolveName(name: string): EvalResult | null;
  // The cell currently being evaluated (for ROW()/COLUMN() with no args).
  currentCell?: { row: number; col: number };
  // The used extent of a sheet (for whole-column/row references like A:A).
  extent(sheet: string | undefined): { maxRow: number; maxCol: number };
}

export class Evaluator {
  constructor(private ctx: EvalContext) {}

  // Exposed to functions that need reference/position info (ROW, OFFSET, …).
  current() {
    return this.ctx.currentCell;
  }
  extentOf(sheet: string | undefined) {
    return this.ctx.extent(sheet);
  }

  evalNode(node: Node): EvalResult {
    switch (node.kind) {
      case "num":
        return node.value;
      case "str":
        return node.value;
      case "bool":
        return node.value;
      case "ref":
        return this.ctx.getCellValue(node.sheet, node.ref.row, node.ref.col);
      case "range": {
        let r0 = node.start.row;
        let r1 = node.end.row;
        let c0 = node.start.col;
        let c1 = node.end.col;
        if (node.fullCol || node.fullRow) {
          const ext = this.ctx.extent(node.sheet);
          if (node.fullCol) { r0 = 0; r1 = ext.maxRow; }
          if (node.fullRow) { c0 = 0; c1 = ext.maxCol; }
        }
        return this.evalRange({ ...node.start, row: r0, col: c0 }, { ...node.end, row: r1, col: c1 }, node.sheet);
      }
      case "name": {
        const r = this.ctx.resolveName(node.name);
        return r === null ? new CellError("#NAME?", `Unknown name ${node.name}`) : r;
      }
      case "unary":
        return this.evalUnary(node.op, this.evalNode(node.operand));
      case "binary":
        return this.evalBinary(node.op, node.left, node.right);
      case "call": {
        const fn = FUNCTIONS[node.name];
        if (!fn) return new CellError("#NAME?", `Unknown function ${node.name}`);
        return fn(node.args, this);
      }
    }
  }

  private evalRange(start: CellRef, end: CellRef, sheet?: string): RangeValue {
    const r0 = Math.min(start.row, end.row);
    const r1 = Math.max(start.row, end.row);
    const c0 = Math.min(start.col, end.col);
    const c1 = Math.max(start.col, end.col);
    const values: CellValue[][] = [];
    for (let r = r0; r <= r1; r++) {
      const row: CellValue[] = [];
      for (let c = c0; c <= c1; c++) row.push(this.ctx.getCellValue(sheet, r, c));
      values.push(row);
    }
    return { range: true, r0, c0, r1, c1, values };
  }

  private evalUnary(op: string, v: EvalResult): EvalResult {
    const s = asScalar(v);
    if (s instanceof CellError) return s;
    if (op === "%") {
      const n = toNumber(s);
      return n instanceof CellError ? n : n / 100;
    }
    if (op === "-") {
      const n = toNumber(s);
      return n instanceof CellError ? n : -n;
    }
    // unary +
    const n = toNumber(s);
    return n instanceof CellError ? n : n;
  }

  private evalBinary(op: string, leftN: Node, rightN: Node): EvalResult {
    const l = asScalar(this.evalNode(leftN));
    if (l instanceof CellError) return l;
    const r = asScalar(this.evalNode(rightN));
    if (r instanceof CellError) return r;

    if (op === "&") {
      return toText(l) + toText(r);
    }

    if (["=", "<>", "<", ">", "<=", ">="].includes(op)) {
      return compare(op, l, r);
    }

    // arithmetic
    const a = toNumber(l);
    if (a instanceof CellError) return a;
    const b = toNumber(r);
    if (b instanceof CellError) return b;
    switch (op) {
      case "+":
        return a + b;
      case "-":
        return a - b;
      case "*":
        return a * b;
      case "/":
        return b === 0 ? new CellError("#DIV/0!") : a / b;
      case "^":
        return Math.pow(a, b);
    }
    return new CellError("#VALUE!", `Unknown operator ${op}`);
  }
}

// --- Coercions (exported for the function library) ---

// Collapse a range to a scalar via implicit intersection (1x1 ok, else #VALUE!).
export function asScalar(v: EvalResult): CellValue {
  if (isRange(v)) {
    if (v.values.length === 1 && v.values[0].length === 1) return v.values[0][0];
    return new CellError("#VALUE!", "Expected a single value");
  }
  return v;
}

export function toNumber(v: CellValue): number | CellError {
  if (v instanceof CellError) return v;
  if (v === null || v === "") return 0;
  if (typeof v === "number") return v;
  if (typeof v === "boolean") return v ? 1 : 0;
  if (typeof v === "string") {
    const t = v.trim();
    if (t === "") return 0;
    const n = Number(t);
    if (!Number.isNaN(n)) return n;
    return new CellError("#VALUE!", `Cannot convert "${v}" to a number`);
  }
  return new CellError("#VALUE!");
}

export function toText(v: CellValue): string {
  if (v instanceof CellError) return v.kind;
  if (v === null) return "";
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "number") return numToText(v);
  return v;
}

export function numToText(n: number): string {
  if (Number.isInteger(n)) return String(n);
  // Trim floating point noise similar to spreadsheet display.
  return String(Number(n.toPrecision(15)));
}

export function toBool(v: CellValue): boolean | CellError {
  if (v instanceof CellError) return v;
  if (typeof v === "boolean") return v;
  if (v === null || v === "") return false;
  if (typeof v === "number") return v !== 0;
  if (typeof v === "string") {
    const t = v.trim().toUpperCase();
    if (t === "TRUE") return true;
    if (t === "FALSE") return false;
    return new CellError("#VALUE!");
  }
  return new CellError("#VALUE!");
}

function compare(op: string, l: CellValue, r: CellValue): CellValue {
  let cmp: number;
  // Numbers compare numerically; otherwise compare as text (case-insensitive).
  const ln = typeof l === "number" || typeof l === "boolean";
  const rn = typeof r === "number" || typeof r === "boolean";
  if ((ln || l === null) && (rn || r === null)) {
    const a = l === null ? 0 : typeof l === "boolean" ? (l ? 1 : 0) : (l as number);
    const b = r === null ? 0 : typeof r === "boolean" ? (r ? 1 : 0) : (r as number);
    cmp = a < b ? -1 : a > b ? 1 : 0;
  } else {
    const a = toText(l).toUpperCase();
    const b = toText(r).toUpperCase();
    cmp = a < b ? -1 : a > b ? 1 : 0;
  }
  switch (op) {
    case "=":
      return cmp === 0;
    case "<>":
      return cmp !== 0;
    case "<":
      return cmp < 0;
    case ">":
      return cmp > 0;
    case "<=":
      return cmp <= 0;
    case ">=":
      return cmp >= 0;
  }
  return new CellError("#VALUE!");
}
