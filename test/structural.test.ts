// Tests for structural edits: AST printing, reference rewriting on
// insert/delete, fill-handle offsetting, and end-to-end recalculation after a
// row insert (mirrors what App.insertRows does).

import { Sheet } from "../src/model/sheet.js";
import { Workbook } from "../src/model/workbook.js";
import { CalcEngine } from "../src/engine/calc.js";
import { rewriteFormulaRefs, offsetFormula } from "../src/engine/printer.js";
import { CellError } from "../src/model/types.js";

let failures = 0;
let count = 0;
function eq(actual: unknown, expected: unknown, label: string) {
  count++;
  const a = actual instanceof CellError ? actual.kind : actual;
  if (a !== expected) {
    failures++;
    console.error(`✗ ${label}: expected ${JSON.stringify(expected)}, got ${JSON.stringify(a)}`);
  } else {
    console.log(`✓ ${label}`);
  }
}

// --- reference rewriting ---
const insertRow0 = (r: number, c: number) => ({ row: r >= 0 ? r + 1 : r, col: c });
eq(rewriteFormulaRefs("=A1+B2", insertRow0), "=A2+B3", "insert row shifts refs down");
eq(rewriteFormulaRefs("=SUM(A1:A3)", insertRow0), "=SUM(A2:A4)", "insert shifts range");

const delRow0 = (r: number, c: number) => (r === 0 ? null : { row: r - 1, col: c });
eq(rewriteFormulaRefs("=A2+B3", delRow0), "=A1+B2", "delete row shifts refs up");
eq(rewriteFormulaRefs("=A1", delRow0), "=#REF!", "deleting referenced row → #REF!");

// --- fill offset (relative moves, absolute stays) ---
eq(offsetFormula("=A1", 1, 0), "=A2", "fill down offsets relative ref");
eq(offsetFormula("=$A$1", 1, 0), "=$A$1", "absolute ref unchanged");
eq(offsetFormula("=A$1+$B2", 1, 1), "=B$1+$B3", "mixed absolute/relative");
eq(offsetFormula("=SUM(A1:A2)", 0, 1), "=SUM(B1:B2)", "fill right offsets range");

// --- Sheet structural data move ---
{
  const s = new Sheet("S");
  s.ensureCell(2, 0).raw = "hello"; // A3
  s.ensureCell(2, 0).value = "hello";
  s.insertRows(0, 1);
  eq(s.getRaw(3, 0), "hello", "insertRows moves cell data down");
  s.deleteRows(0, 1);
  eq(s.getRaw(2, 0), "hello", "deleteRows moves cell data back");
}

// --- end-to-end: insert a row and confirm formulas recompute correctly ---
{
  const wb = new Workbook();
  const sheet = wb.active;
  const engine = new CalcEngine(wb);
  engine.setCellRaw(0, 0, "10"); // A1
  engine.setCellRaw(1, 0, "20"); // A2
  engine.setCellRaw(2, 0, "=SUM(A1:A2)"); // A3
  eq(sheet.getCell(2, 0)?.value, 30, "baseline SUM");

  // App.insertRows(0,1): move data, rewrite refs, rebuild
  sheet.insertRows(0, 1);
  for (const [, cell] of sheet.cells) {
    if (cell.raw.startsWith("=")) cell.raw = rewriteFormulaRefs(cell.raw, insertRow0);
  }
  engine.rebuild();

  eq(sheet.getRaw(3, 0), "=SUM(A2:A3)", "formula references rewritten after insert");
  eq(sheet.getCell(3, 0)?.value, 30, "formula recomputes to same total after insert");
  eq(sheet.getCell(1, 0)?.value, 10, "data shifted to A2");
}

// --- merges survive a column insert ---
{
  const s = new Sheet("M");
  s.addMerge({ r0: 0, c0: 1, r1: 0, c1: 2 }); // B1:C1
  s.insertCols(0, 1); // insert col A → merge should become C1:D1
  const m = s.mergeAt(0, 2);
  eq(m?.c0, 2, "merge left edge shifted by column insert");
  eq(m?.c1, 3, "merge right edge shifted by column insert");
}

console.log(`\n${count - failures}/${count} passed`);
if (failures) process.exit(1);
