// Verifies the .xlsx WRITE→READ round-trip (the read path uses DOMParser, which
// we inject here from @xmldom/xmldom so it can run under Node). This is the path
// exercised when a user opens an .xlsx file in the app.

import { DOMParser } from "@xmldom/xmldom";
(globalThis as unknown as { DOMParser: unknown }).DOMParser = DOMParser;

import { Workbook } from "../src/model/workbook.js";
import { CalcEngine } from "../src/engine/calc.js";
import { writeXlsx, readXlsx } from "../src/io/xlsx.js";
import { CellError } from "../src/model/types.js";

let failures = 0;
let count = 0;
function eq(actual: unknown, expected: unknown, label: string) {
  count++;
  const a = actual instanceof CellError ? actual.kind : actual;
  const ok = a === expected || (typeof a === "number" && typeof expected === "number" && Math.abs(a - expected) < 1e-9);
  if (!ok) {
    failures++;
    console.error(`✗ ${label}: expected ${JSON.stringify(expected)}, got ${JSON.stringify(a)}`);
  } else {
    console.log(`✓ ${label}`);
  }
}

// --- build a workbook with a bit of everything ---
const wb = new Workbook();
wb.active.name = "Data";
wb.addSheet("Summary");
const engine = new CalcEngine(wb);

wb.activeIndex = 0; // Data
engine.setCellRaw(0, 0, "Item");
engine.setCellRaw(0, 1, "Qty");
engine.setCellRaw(1, 0, "Apples");
engine.setCellRaw(1, 1, "10");
engine.setCellRaw(2, 0, "Pears");
engine.setCellRaw(2, 1, "20");
engine.setCellRaw(3, 1, "=SUM(B2:B3)"); // 30
wb.sheets[0].ensureCell(0, 0).format = { bold: true, color: "#ff0000", bg: "#ffff00" };
wb.sheets[0].ensureCell(1, 1).format = { numFmt: "0.00" };

wb.activeIndex = 1; // Summary
engine.setCellRaw(0, 0, "Total");
engine.setCellRaw(0, 1, "=Data!B4"); // cross-sheet → 30

// --- write then read back ---
const bytes = writeXlsx(wb);
const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
const wb2 = readXlsx(ab);
const engine2 = new CalcEngine(wb2);
engine2.rebuild();

// --- assertions ---
eq(wb2.sheets.length, 2, "two sheets read back");
eq(wb2.sheets[0].name, "Data", "first sheet name");
eq(wb2.sheets[1].name, "Summary", "second sheet name");

const data = wb2.sheets[0];
eq(data.getCell(0, 0)?.value, "Item", "string cell value");
eq(data.getCell(1, 1)?.value, 10, "numeric cell value");
eq(data.getRaw(3, 1), "=SUM(B2:B3)", "formula raw preserved");
eq(data.getCell(3, 1)?.value, 30, "formula recomputes after read");

const f = data.getCell(0, 0)?.format;
eq(f?.bold, true, "bold format preserved");
eq(f?.color?.toLowerCase(), "#ff0000", "text color preserved");
eq(f?.bg?.toLowerCase(), "#ffff00", "fill color preserved");
eq(data.getCell(1, 1)?.format?.numFmt, "0.00", "number format preserved");

const summary = wb2.sheets[1];
eq(summary.getRaw(0, 1), "=Data!B4", "cross-sheet formula raw preserved");
eq(summary.getCell(0, 1)?.value, 30, "cross-sheet formula recomputes after read");

console.log(`\n${count - failures}/${count} passed`);
if (failures) process.exit(1);
