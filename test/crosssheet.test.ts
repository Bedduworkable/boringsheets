// Cross-sheet reference tests: parsing Sheet2!A1 and 'My Sheet'!A1, evaluating
// across sheets, recalculating through the workbook-wide dependency graph,
// circular detection across sheets, and the rename-sheet formula rewrite.

import { Workbook } from "../src/model/workbook.js";
import { CalcEngine } from "../src/engine/calc.js";
import { CellError } from "../src/model/types.js";
import { renameSheetRefs } from "../src/engine/printer.js";

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

const wb = new Workbook(); // Sheet1 at index 0
wb.addSheet("Sheet2"); // index 1
const engine = new CalcEngine(wb);

const rc = (a1: string) => {
  const m = /^([A-Z]+)(\d+)$/.exec(a1)!;
  const col = m[1].split("").reduce((n, c) => n * 26 + (c.charCodeAt(0) - 64), 0) - 1;
  return { row: parseInt(m[2], 10) - 1, col };
};
const setOn = (idx: number, a1: string, raw: string) => {
  wb.activeIndex = idx;
  const { row, col } = rc(a1);
  engine.setCellRaw(row, col, raw);
};
const valOn = (idx: number, a1: string) => {
  const { row, col } = rc(a1);
  return wb.sheets[idx].getCell(row, col)?.value ?? null;
};

// data on Sheet2
setOn(1, "A1", "100");
setOn(1, "A2", "200");
setOn(1, "A3", "300");

// Sheet1 reads Sheet2
setOn(0, "B1", "=Sheet2!A1");
eq(valOn(0, "B1"), 100, "cross-sheet single reference");
setOn(0, "C1", "=SUM(Sheet2!A1:A3)");
eq(valOn(0, "C1"), 600, "cross-sheet range");
setOn(0, "C2", "=Sheet2!A1*2+B1");
eq(valOn(0, "C2"), 300, "mixed cross-sheet + local reference");

// change a Sheet2 cell → Sheet1 formulas recalc through the cross-sheet graph
setOn(1, "A1", "150");
eq(valOn(0, "B1"), 150, "dependent on another sheet recomputed");
eq(valOn(0, "C1"), 650, "cross-sheet range recomputed");
eq(valOn(0, "C2"), 450, "mixed recomputed");

// unknown sheet → #REF!
setOn(0, "D1", "=Nope!A1");
eq(valOn(0, "D1"), "#REF!", "reference to missing sheet is #REF!");

// quoted sheet names with spaces
wb.sheets[1].name = "My Data";
setOn(0, "E1", "='My Data'!A2");
eq(valOn(0, "E1"), 200, "quoted sheet name reference");

// circular across sheets
setOn(0, "F1", "='My Data'!Z1");
setOn(1, "Z1", "=Sheet1!F1");
eq(valOn(0, "F1"), "#CIRCULAR!", "cross-sheet circular reference detected");

// printer: rename rewrites references and re-quotes when needed
eq(renameSheetRefs("=Sheet2!A1+1", "Sheet2", "My Data"), "='My Data'!A1+1", "renameSheetRefs quotes new name");
eq(renameSheetRefs("=SUM(Data!A1:A3)", "Data", "Numbers"), "=SUM(Numbers!A1:A3)", "renameSheetRefs on range");
eq(renameSheetRefs("=Other!A1", "Data", "Numbers"), "=Other!A1", "renameSheetRefs leaves other sheets alone");

console.log(`\n${count - failures}/${count} passed`);
if (failures) process.exit(1);
