// Full-fidelity .xlsx round-trip: every feature that used to be lost on save
// (merges, borders, wrap, freeze, hidden rows/cols, row heights, data
// validation, conditional formatting, named ranges, notes) must survive
// write → read. DOMParser is injected from @xmldom/xmldom for Node.

import { DOMParser } from "@xmldom/xmldom";
(globalThis as unknown as { DOMParser: unknown }).DOMParser = DOMParser;

import { Workbook } from "../src/model/workbook.js";
import { CalcEngine } from "../src/engine/calc.js";
import { writeXlsx, readXlsx } from "../src/io/xlsx.js";

let failures = 0;
let count = 0;
function ok(cond: boolean, label: string) {
  count++;
  if (!cond) {
    failures++;
    console.error(`✗ ${label}`);
  } else {
    console.log(`✓ ${label}`);
  }
}

// --- build a workbook exercising everything ---
const wb = new Workbook();
wb.active.name = "Main";
wb.addSheet("Two");
const engine = new CalcEngine(wb);

wb.activeIndex = 0;
const s = wb.active;
engine.setCellRaw(0, 0, "Header");
engine.setCellRaw(1, 0, "10");
engine.setCellRaw(2, 0, "20");
engine.setCellRaw(3, 0, "=SUM(A2:A3)");
engine.setCellRaw(0, 1, "wrapped long text here");

// formatting
s.ensureCell(0, 0).format = { bold: true, border: { top: true, bottom: true, left: true, right: true, color: "#ff0000" } };
s.ensureCell(0, 1).format = { wrap: true };
// merge, freeze, hidden, row height
s.addMerge({ r0: 5, c0: 0, r1: 5, c1: 2 });
engine.setCellRaw(5, 0, "Merged title");
s.frozenRows = 1;
s.frozenCols = 1;
s.hiddenRows.add(2);
s.hiddenCols.add(3);
s.rowHeights.set(0, 40);
// note
s.ensureCell(1, 0).note = "this is the first value";
// data validation (list + number)
s.validations.push({ range: { r0: 10, c0: 0, r1: 12, c1: 0 }, type: "list", source: ["A", "B", "C"], allowBlank: true });
s.validations.push({ range: { r0: 10, c0: 1, r1: 12, c1: 1 }, type: "number", operator: "between", min: 1, max: 100 });
// conditional formatting (cellIs + colorScale + dataBar)
s.conditionalRules.push({ id: "1", range: { r0: 1, c0: 0, r1: 3, c1: 0 }, type: "greaterThan", value1: 15, format: { bg: "#ffff00", bold: true } });
s.conditionalRules.push({ id: "2", range: { r0: 1, c0: 0, r1: 3, c1: 0 }, type: "colorScale", minColor: "#ff0000", maxColor: "#00ff00" });
// named range
wb.names.push({ name: "Numbers", sheetId: s.id, r0: 1, c0: 0, r1: 2, c1: 0 });

// --- round-trip ---
const bytes = writeXlsx(wb);
const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
const wb2 = readXlsx(ab);
new CalcEngine(wb2).rebuild();
const s2 = wb2.sheets[0];

// --- assertions ---
ok(s2.getCell(3, 0)?.value === 30, "formula recomputes after load");
ok(s2.getCell(0, 0)?.format?.bold === true, "bold preserved");
ok(!!s2.getCell(0, 0)?.format?.border?.top, "border top preserved");
ok(s2.getCell(0, 0)?.format?.border?.color?.toLowerCase() === "#ff0000", "border color preserved");
ok(s2.getCell(0, 1)?.format?.wrap === true, "wrap preserved");
ok(s2.merges.length === 1 && s2.merges[0].r0 === 5 && s2.merges[0].c1 === 2, "merge preserved");
ok(s2.frozenRows === 1 && s2.frozenCols === 1, "freeze panes preserved");
ok(s2.hiddenRows.has(2), "hidden row preserved");
ok(s2.hiddenCols.has(3), "hidden column preserved");
ok(s2.rowHeights.get(0) !== undefined && Math.abs(s2.rowHeights.get(0)! - 40) <= 2, "row height preserved");
ok(s2.getCell(1, 0)?.note === "this is the first value", "note preserved");
ok(s2.validations.length === 2, "two validations preserved");
const listV = s2.validations.find((v) => v.type === "list");
ok(!!listV && JSON.stringify(listV.source) === JSON.stringify(["A", "B", "C"]), "list validation values preserved");
const numV = s2.validations.find((v) => v.type === "number");
ok(!!numV && numV.operator === "between" && numV.min === 1 && numV.max === 100, "number validation preserved");
ok(s2.conditionalRules.length === 2, "two conditional rules preserved");
const cis = s2.conditionalRules.find((r) => r.type === "greaterThan");
ok(!!cis && cis.value1 === 15 && cis.format?.bg?.toLowerCase() === "#ffff00", "cellIs rule + dxf format preserved");
const cscale = s2.conditionalRules.find((r) => r.type === "colorScale");
ok(!!cscale && cscale.minColor?.toLowerCase() === "#ff0000", "color scale preserved");
ok(wb2.names.length === 1 && wb2.names[0].name === "Numbers", "named range preserved");
const nr = wb2.names[0];
ok(nr.r0 === 1 && nr.c0 === 0 && nr.r1 === 2 && nr.c1 === 0, "named range coords preserved");
ok(wb2.sheets.length === 2 && wb2.sheets[1].name === "Two", "second sheet preserved");

console.log(`\n${count - failures}/${count} passed`);
if (failures) process.exit(1);
