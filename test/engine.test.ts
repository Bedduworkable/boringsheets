// Headless smoke test for the calc engine: parsing, evaluation, dependency
// recalculation, functions, and formatting. Bundled with esbuild and run under
// Node (see the command in the chat). Exits non-zero on failure.

import { Workbook } from "../src/model/workbook.js";
import { CalcEngine } from "../src/engine/calc.js";
import { CellError } from "../src/model/types.js";
import { formatValue } from "../src/engine/format.js";

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

const wb = new Workbook();
const sheet = wb.active;
const engine = new CalcEngine(wb);
const set = (a1: string, raw: string) => {
  const m = /^([A-Z]+)(\d+)$/.exec(a1)!;
  const col = m[1].split("").reduce((n, c) => n * 26 + (c.charCodeAt(0) - 64), 0) - 1;
  const row = parseInt(m[2], 10) - 1;
  engine.setCellRaw(row, col, raw);
};
const val = (a1: string) => {
  const m = /^([A-Z]+)(\d+)$/.exec(a1)!;
  const col = m[1].split("").reduce((n, c) => n * 26 + (c.charCodeAt(0) - 64), 0) - 1;
  const row = parseInt(m[2], 10) - 1;
  return sheet.getCell(row, col)?.value ?? null;
};

// --- literals ---
set("A1", "10");
set("A2", "20");
set("A3", "30");
eq(val("A1"), 10, "numeric literal");

// --- arithmetic + precedence ---
set("B1", "=2+3*4");
eq(val("B1"), 14, "operator precedence");
set("B2", "=(2+3)*4");
eq(val("B2"), 20, "parentheses");
set("B3", "=2^3^2");
eq(val("B3"), 512, "right-assoc exponent");
set("B4", "=-2^2");
eq(val("B4"), 4, "unary minus binds tighter than ^ in our grammar");

// --- references + ranges ---
set("C1", "=A1+A2");
eq(val("C1"), 30, "cell reference addition");
set("C2", "=SUM(A1:A3)");
eq(val("C2"), 60, "SUM range");
set("C3", "=AVERAGE(A1:A3)");
eq(val("C3"), 20, "AVERAGE range");
set("C4", "=MAX(A1:A3)");
eq(val("C4"), 30, "MAX");
set("C5", "=MIN(A1:A3)*2");
eq(val("C5"), 20, "MIN with arithmetic");

// --- dependency recalculation ---
set("D1", "=A1*2");
eq(val("D1"), 20, "dependent before change");
set("A1", "100");
eq(val("D1"), 200, "dependent recalculated after precedent change");
eq(val("C2"), 150, "SUM recalculated after member change");

// --- logical + text ---
set("E1", '=IF(A1>50,"big","small")');
eq(val("E1"), "big", "IF true branch");
set("E2", "=IF(A2>50,1,0)");
eq(val("E2"), 0, "IF false branch");
set("E3", '=CONCATENATE("a","b","c")');
eq(val("E3"), "abc", "CONCATENATE");
set("E4", '=UPPER("hello")');
eq(val("E4"), "HELLO", "UPPER");
set("E5", '=LEN("hello")');
eq(val("E5"), 5, "LEN");
set("E6", "=ROUND(3.14159,2)");
eq(val("E6"), 3.14, "ROUND");

// --- errors ---
set("F1", "=1/0");
eq(val("F1"), "#DIV/0!", "division by zero error");
set("F2", "=NOPE(1)");
eq(val("F2"), "#NAME?", "unknown function error");
set("F3", '=1+"abc"');
eq(val("F3"), "#VALUE!", "type coercion error");

// --- circular reference ---
set("G1", "=G2");
set("G2", "=G1");
eq(val("G1"), "#CIRCULAR!", "circular reference detected");

// --- conditional aggregation ---
set("H1", "5");
set("H2", "15");
set("H3", "25");
set("I1", "=SUMIF(H1:H3,\">10\")");
eq(val("I1"), 40, "SUMIF with criteria");
set("I2", "=COUNTIF(H1:H3,\">10\")");
eq(val("I2"), 2, "COUNTIF with criteria");

// --- lookup ---
set("J1", "1"); set("K1", "one");
set("J2", "2"); set("K2", "two");
set("J3", "3"); set("K3", "three");
set("L1", "=VLOOKUP(2,J1:K3,2,FALSE)");
eq(val("L1"), "two", "VLOOKUP exact match");

// --- formatting ---
eq(formatValue(1234.5, "#,##0.00"), "1,234.50", "number format thousands");
eq(formatValue(0.25, "0%"), "25%", "percent format");
eq(formatValue(1234.5, '"$"#,##0.00'), "$1,234.50", "currency format");
eq(formatValue(44927, "yyyy-mm-dd"), "2023-01-01", "date serial format");

console.log(`\n${count - failures}/${count} passed`);
if (failures) {
  console.error(`${failures} FAILED`);
  process.exit(1);
}
