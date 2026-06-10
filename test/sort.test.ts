// Verifies the sort algorithm used by App.sortRange: rows are reordered by a key
// column, and each row's relative formulas are offset by how far the row moved
// so they keep pointing at their own row's data.

import { Workbook } from "../src/model/workbook.js";
import { CalcEngine } from "../src/engine/calc.js";
import { offsetFormula } from "../src/engine/printer.js";
import { CellValue } from "../src/model/types.js";

let failures = 0;
let count = 0;
function eq(actual: unknown, expected: unknown, label: string) {
  count++;
  if (actual !== expected) {
    failures++;
    console.error(`✗ ${label}: expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
  } else {
    console.log(`✓ ${label}`);
  }
}

const wb = new Workbook();
const sheet = wb.active;
const engine = new CalcEngine(wb);

// A = key, B = formula referencing same row's A
engine.setCellRaw(0, 0, "3"); engine.setCellRaw(0, 1, "=A1*10");
engine.setCellRaw(1, 0, "1"); engine.setCellRaw(1, 1, "=A2*10");
engine.setCellRaw(2, 0, "2"); engine.setCellRaw(2, 1, "=A3*10");

// --- replicate App.sortRange core for the range A1:B3, key col 0, ascending ---
const r0 = 0, r1 = 2, c0 = 0, c1 = 1, keyCol = 0;
interface RowData { origRow: number; key: CellValue; cells: { raw: string }[] }
const rows: RowData[] = [];
for (let r = r0; r <= r1; r++) {
  const cells: { raw: string }[] = [];
  for (let c = c0; c <= c1; c++) cells.push({ raw: sheet.getRaw(r, c) });
  rows.push({ origRow: r, key: sheet.getCell(r, keyCol)?.value ?? null, cells });
}
const cmp = (a: CellValue, b: CellValue) => (a as number) - (b as number);
const sorted = rows.map((r, i) => ({ r, i })).sort((a, b) => cmp(a.r.key, b.r.key) || a.i - b.i).map((x) => x.r);
sorted.forEach((rowData, idx) => {
  const newRow = r0 + idx;
  const dRow = newRow - rowData.origRow;
  rowData.cells.forEach((cd, ci) => {
    const c = c0 + ci;
    const raw = cd.raw.startsWith("=") && dRow !== 0 ? offsetFormula(cd.raw, dRow, 0) : cd.raw;
    engine.setCellRaw(newRow, c, raw);
  });
});

eq(sheet.getCell(0, 0)?.value, 1, "sorted key A1 = 1");
eq(sheet.getCell(1, 0)?.value, 2, "sorted key A2 = 2");
eq(sheet.getCell(2, 0)?.value, 3, "sorted key A3 = 3");
eq(sheet.getRaw(0, 1), "=A1*10", "B1 formula re-anchored to its own row");
eq(sheet.getCell(0, 1)?.value, 10, "B1 recomputes to 10");
eq(sheet.getCell(1, 1)?.value, 20, "B2 recomputes to 20");
eq(sheet.getCell(2, 1)?.value, 30, "B3 recomputes to 30");

// --- hidden-row (filter) round-trip via snapshot ---
sheet.hiddenRows.add(1);
eq(sheet.rowHeight(1), 0, "filtered row has zero height");
const snap = sheet.snapshot();
sheet.hiddenRows.clear();
eq(sheet.rowHeight(1) > 0, true, "row visible after clearing filter");
sheet.restore(snap);
eq(sheet.hiddenRows.has(1), true, "snapshot restores hidden rows");

console.log(`\n${count - failures}/${count} passed`);
if (failures) process.exit(1);
