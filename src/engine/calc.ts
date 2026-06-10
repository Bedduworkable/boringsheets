// The calculation engine. Workbook-aware: owns formula parsing, a dependency
// graph that spans every sheet (so Sheet1 formulas can depend on Sheet2 cells),
// and incremental recalculation with circular-reference detection.

import { Workbook } from "../model/workbook.js";
import { Sheet } from "../model/sheet.js";
import { CellValue, CellError } from "../model/types.js";
import { Node } from "./ast.js";
import { parseFormula, ParseError } from "./parser.js";
import { Evaluator, asScalar } from "./evaluator.js";
import { cellKey, parseKey } from "./references.js";

interface FormulaEntry {
  ast: Node;
  home: Sheet;
  row: number;
  col: number;
  precedents: string[]; // global keys this formula reads
}

export class CalcEngine {
  // global key = `${sheetId}|${row},${col}`
  private formulas = new Map<string, FormulaEntry>();
  private dependents = new Map<string, Set<string>>();

  constructor(private wb: Workbook) {}

  // Kept for call-site compatibility; the graph is workbook-wide so switching
  // the visible sheet needs no engine work.
  setSheet(_sheet: Sheet) {}

  private gkey(sheetId: number, row: number, col: number): string {
    return `${sheetId}|${row},${col}`;
  }

  private resolveSheet(name: string | undefined, home: Sheet): Sheet | null {
    if (name === undefined) return home;
    const lower = name.toLowerCase();
    return this.wb.sheets.find((s) => s.name.toLowerCase() === lower) ?? null;
  }

  // Rebuild the entire workbook's graph + values from scratch.
  rebuild() {
    this.formulas.clear();
    this.dependents.clear();
    for (const sheet of this.wb.sheets) {
      for (const [k, cell] of sheet.cells) {
        const { row, col } = parseKey(k);
        if (cell.raw.startsWith("=")) this.installFormula(sheet, row, col, cell.raw);
        else cell.value = parseLiteral(cell.raw);
      }
    }
    this.recompute(new Set(this.formulas.keys()));
  }

  // Set a cell's raw text on the active sheet; returns the global keys whose
  // value changed (the renderer only cares about the active sheet's subset).
  setCellRaw(row: number, col: number, raw: string): Set<string> {
    const sheet = this.wb.active;
    const gk = this.gkey(sheet.id, row, col);
    this.removeFormula(gk);

    if (raw === "") {
      const c = sheet.cells.get(cellKey(row, col));
      if (c) {
        c.raw = "";
        c.value = null;
        if (!c.format) sheet.cells.delete(cellKey(row, col));
      }
    } else {
      const cell = sheet.ensureCell(row, col);
      cell.raw = raw;
      if (raw.startsWith("=")) this.installFormula(sheet, row, col, raw);
      else cell.value = parseLiteral(raw);
    }

    return this.recompute(new Set([gk]));
  }

  private installFormula(sheet: Sheet, row: number, col: number, raw: string) {
    const gk = this.gkey(sheet.id, row, col);
    let ast: Node;
    try {
      ast = parseFormula(raw.slice(1));
    } catch (e) {
      const cell = sheet.cells.get(cellKey(row, col))!;
      cell.value = new CellError("#NAME?", e instanceof ParseError ? e.message : String(e));
      return;
    }
    const precedents = this.extractPrecedents(ast, sheet);
    this.formulas.set(gk, { ast, home: sheet, row, col, precedents });
    for (const p of precedents) {
      let set = this.dependents.get(p);
      if (!set) this.dependents.set(p, (set = new Set()));
      set.add(gk);
    }
  }

  private removeFormula(gk: string) {
    const f = this.formulas.get(gk);
    if (!f) return;
    for (const p of f.precedents) this.dependents.get(p)?.delete(gk);
    this.formulas.delete(gk);
  }

  private recompute(changed: Set<string>): Set<string> {
    // 1. transitive closure over dependents
    const affected = new Set<string>(changed);
    const queue = [...changed];
    while (queue.length) {
      const k = queue.shift()!;
      const deps = this.dependents.get(k);
      if (deps) for (const d of deps) if (!affected.has(d)) { affected.add(d); queue.push(d); }
    }

    // 2. evaluate affected formula cells on demand, with cycle detection
    const computed = new Set<string>();
    const visiting = new Set<string>();

    const readCell = (target: Sheet, row: number, col: number): CellValue => {
      const tk = this.gkey(target.id, row, col);
      if (affected.has(tk) && this.formulas.has(tk)) return compute(tk);
      return target.cells.get(cellKey(row, col))?.value ?? null;
    };

    const valueOf = (home: Sheet, sheetName: string | undefined, row: number, col: number): CellValue => {
      const target = this.resolveSheet(sheetName, home);
      if (!target) return new CellError("#REF!", `Unknown sheet ${sheetName}`);
      return readCell(target, row, col);
    };

    const resolveName = (name: string) => {
      const nr = this.wb.names.find((n) => n.name.toLowerCase() === name.toLowerCase());
      if (!nr) return null;
      const target = this.wb.sheets.find((s) => s.id === nr.sheetId);
      if (!target) return new CellError("#REF!");
      const values: CellValue[][] = [];
      for (let r = nr.r0; r <= nr.r1; r++) {
        const rowv: CellValue[] = [];
        for (let c = nr.c0; c <= nr.c1; c++) rowv.push(readCell(target, r, c));
        values.push(rowv);
      }
      return { range: true as const, r0: nr.r0, c0: nr.c0, r1: nr.r1, c1: nr.c1, values };
    };

    const compute = (gk: string): CellValue => {
      const f = this.formulas.get(gk);
      if (!f) return null;
      const cell = f.home.cells.get(cellKey(f.row, f.col));
      if (computed.has(gk)) return cell?.value ?? null;
      if (visiting.has(gk)) {
        if (cell) cell.value = new CellError("#CIRCULAR!");
        return cell?.value ?? new CellError("#CIRCULAR!");
      }
      visiting.add(gk);
      const ev = new Evaluator({
        getCellValue: (sn, r, c) => valueOf(f.home, sn, r, c),
        resolveName,
        currentCell: { row: f.row, col: f.col },
        extent: (sn) => {
          const sheet = this.resolveSheet(sn, f.home);
          return { maxRow: sheet?.maxRow ?? 0, maxCol: sheet?.maxCol ?? 0 };
        },
      });
      let result: CellValue;
      try {
        result = asScalar(ev.evalNode(f.ast));
      } catch (e) {
        result = new CellError("#VALUE!", String(e));
      }
      visiting.delete(gk);
      computed.add(gk);
      if (cell) cell.value = result;
      return result;
    };

    for (const gk of affected) if (this.formulas.has(gk)) compute(gk);
    return affected;
  }

  private extractPrecedents(node: Node, home: Sheet): string[] {
    const keys = new Set<string>();
    const add = (sheetName: string | undefined, row: number, col: number) => {
      const target = this.resolveSheet(sheetName, home);
      if (target) keys.add(this.gkey(target.id, row, col));
    };
    const walk = (n: Node) => {
      switch (n.kind) {
        case "ref":
          add(n.sheet, n.ref.row, n.ref.col);
          break;
        case "range": {
          const target = this.resolveSheet(n.sheet, home);
          const r0 = n.fullCol ? 0 : Math.min(n.start.row, n.end.row);
          const r1 = n.fullCol ? target?.maxRow ?? 0 : Math.max(n.start.row, n.end.row);
          const c0 = n.fullRow ? 0 : Math.min(n.start.col, n.end.col);
          const c1 = n.fullRow ? target?.maxCol ?? 0 : Math.max(n.start.col, n.end.col);
          for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) add(n.sheet, r, c);
          break;
        }
        case "unary":
          walk(n.operand);
          break;
        case "binary":
          walk(n.left);
          walk(n.right);
          break;
        case "call":
          n.args.forEach(walk);
          break;
        case "name": {
          const nr = this.wb.names.find((x) => x.name.toLowerCase() === n.name.toLowerCase());
          if (nr) {
            const target = this.wb.sheets.find((s) => s.id === nr.sheetId);
            if (target) {
              for (let r = nr.r0; r <= nr.r1; r++)
                for (let c = nr.c0; c <= nr.c1; c++) keys.add(this.gkey(target.id, r, c));
            }
          }
          break;
        }
      }
    };
    walk(node);
    return [...keys];
  }
}

const ERROR_KINDS = new Set(["#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#NULL!", "#CIRCULAR!"]);

// Parse a non-formula literal into a typed value.
export function parseLiteral(raw: string): CellValue {
  if (raw === "") return null;
  const t = raw.trim();
  const upper = t.toUpperCase();
  if (upper === "TRUE") return true;
  if (upper === "FALSE") return false;
  // Excel error literals (also what a loaded error cell's raw text is).
  if (ERROR_KINDS.has(upper)) return new CellError(upper as ConstructorParameters<typeof CellError>[0]);

  if (/^-?\d*\.?\d+\s*%$/.test(t)) return parseFloat(t) / 100;

  const stripped = t.replace(/^[$£€]\s*/, "").replace(/,/g, "");
  if (isNumericLiteral(stripped) && (t !== stripped || isNumericLiteral(t))) {
    const n = Number(stripped);
    if (!Number.isNaN(n)) return n;
  }
  if (isNumericLiteral(t)) return Number(t);
  return raw;
}

function isNumericLiteral(s: string): boolean {
  if (s === "") return false;
  return /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(s);
}
