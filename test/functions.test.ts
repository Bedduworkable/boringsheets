// Headless test for the expanded built-in function library. Mirrors the harness
// in engine.test.ts. Bundled with esbuild and run under Node.

import { Workbook } from "../src/model/workbook.js";
import { CalcEngine } from "../src/engine/calc.js";
import { CellError } from "../src/model/types.js";

let failures = 0;
let count = 0;

function eq(actual: unknown, expected: unknown, label: string) {
  count++;
  const a = actual instanceof CellError ? actual.kind : actual;
  const ok =
    a === expected ||
    (typeof a === "number" &&
      typeof expected === "number" &&
      // relative tolerance: expected values in this file are hand-rounded
      Math.abs(a - expected) <= 1e-5 * Math.max(1, Math.abs(expected)));
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
// Evaluate a formula directly in a scratch cell and read it back.
let scratch = 0;
const f = (formula: string) => {
  const cell = `Z${++scratch}`;
  set(cell, formula);
  return val(cell);
};

// --- Math / trig ---
eq(f("=TRUNC(3.78)"), 3, "TRUNC");
eq(f("=TRUNC(-3.78,1)"), -3.7, "TRUNC with digits");
eq(f("=EVEN(3)"), 4, "EVEN");
eq(f("=ODD(2)"), 3, "ODD");
eq(f("=QUOTIENT(17,5)"), 3, "QUOTIENT");
eq(f("=GCD(24,36)"), 12, "GCD");
eq(f("=LCM(4,6)"), 12, "LCM");
eq(f("=FACT(5)"), 120, "FACT");
eq(f("=COMBIN(5,2)"), 10, "COMBIN");
eq(f("=PERMUT(5,2)"), 20, "PERMUT");
eq(f("=SUMSQ(3,4)"), 25, "SUMSQ");
eq(f("=MROUND(10,3)"), 9, "MROUND");
eq(f("=MROUND(5,-2)"), "#NUM!", "MROUND sign mismatch error");
eq(f("=DEGREES(PI())"), 180, "DEGREES");
eq(f("=RADIANS(180)"), Math.PI, "RADIANS");
eq(f("=SIN(0)"), 0, "SIN");
eq(f("=COS(0)"), 1, "COS");
eq(f("=ATAN2(1,1)"), Math.PI / 4, "ATAN2");
eq(f("=LOG(8,2)"), 3, "LOG base 2");
eq(f("=LOG(100)"), 2, "LOG default base 10");

// --- Statistical ---
set("A1", "1"); set("A2", "2"); set("A3", "3"); set("A4", "4"); set("A5", "100");
eq(f("=MEDIAN(A1:A5)"), 3, "MEDIAN");
eq(f("=MEDIAN(A1:A4)"), 2.5, "MEDIAN even count");
set("B1", "2"); set("B2", "4"); set("B3", "4"); set("B4", "4"); set("B5", "5");
set("B6", "5"); set("B7", "7"); set("B8", "9");
eq(f("=STDEV(B1:B8)"), 2.138089935, "STDEV sample");
eq(f("=STDEVP(B1:B8)"), 2, "STDEVP population");
eq(f("=VAR(B1:B8)"), 4.571428571, "VAR sample");
eq(f("=VARP(B1:B8)"), 4, "VARP population");
eq(f("=MODE(B1:B8)"), 4, "MODE");
eq(f("=LARGE(A1:A5,2)"), 4, "LARGE 2nd");
eq(f("=SMALL(A1:A5,2)"), 2, "SMALL 2nd");
eq(f("=RANK(4,A1:A5)"), 2, "RANK descending default");
eq(f("=MEDIAN(B1:B8)"), 4.5, "MEDIAN of B");

// --- Text ---
eq(f("=CHAR(65)"), "A", "CHAR");
eq(f('=CODE("A")'), 65, "CODE");
eq(f('=EXACT("abc","abc")'), true, "EXACT true");
eq(f('=EXACT("abc","ABC")'), false, "EXACT case-sensitive");
eq(f('=FIXED(1234.567,1)'), "1,234.6", "FIXED");
eq(f('=DOLLAR(1234.5,2)'), "$1,234.50", "DOLLAR");
eq(f('=TEXTBEFORE("a-b-c","-")'), "a", "TEXTBEFORE");
eq(f('=TEXTAFTER("a-b-c","-")'), "b-c", "TEXTAFTER");
eq(f('=NUMBERVALUE("1,234.5")'), 1234.5, "NUMBERVALUE");

// --- Logical ---
eq(f('=SWITCH(2,1,"a",2,"b",3,"c")'), "b", "SWITCH match");
eq(f('=SWITCH(9,1,"a",2,"b","def")'), "def", "SWITCH default");
eq(f('=SWITCH(9,1,"a",2,"b")'), "#N/A", "SWITCH no match no default");

// --- Lookup / reference ---
eq(f('=CHOOSE(2,"x","y","z")'), "y", "CHOOSE");
eq(f("=CHOOSE(5,1,2)"), "#VALUE!", "CHOOSE out of range");
set("C1", "10"); set("C2", "20"); set("C3", "30");
set("D1", "ten"); set("D2", "twenty"); set("D3", "thirty");
eq(f("=XLOOKUP(20,C1:C3,D1:D3)"), "twenty", "XLOOKUP hit");
eq(f('=XLOOKUP(99,C1:C3,D1:D3,"none")'), "none", "XLOOKUP not found");
eq(f("=XLOOKUP(99,C1:C3,D1:D3)"), "#N/A", "XLOOKUP not found no default");
eq(f("=LOOKUP(20,C1:C3,D1:D3)"), "twenty", "LOOKUP vector");
eq(f("=ROWS(A1:A5)"), 5, "ROWS");
eq(f("=COLUMNS(C1:D3)"), 2, "COLUMNS");

// --- SUMIFS / COUNTIFS family ---
set("E1", "10"); set("E2", "20"); set("E3", "30"); set("E4", "40");
set("F1", "x"); set("F2", "y"); set("F3", "x"); set("F4", "y");
set("G1", "1"); set("G2", "1"); set("G3", "2"); set("G4", "2");
eq(f('=SUMIFS(E1:E4,F1:F4,"x")'), 40, "SUMIFS one criterion");
eq(f('=SUMIFS(E1:E4,F1:F4,"x",G1:G4,1)'), 10, "SUMIFS two criteria");
eq(f('=COUNTIFS(F1:F4,"y")'), 2, "COUNTIFS");
eq(f('=COUNTIFS(F1:F4,"y",G1:G4,2)'), 1, "COUNTIFS two criteria");
eq(f('=AVERAGEIFS(E1:E4,F1:F4,"y")'), 30, "AVERAGEIFS");
eq(f('=MAXIFS(E1:E4,F1:F4,"x")'), 30, "MAXIFS");
eq(f('=MINIFS(E1:E4,F1:F4,"y")'), 20, "MINIFS");

// --- Date / time ---
eq(f("=DATE(2023,1,1)"), 44927, "DATE serial");
eq(f("=EOMONTH(DATE(2023,1,15),0)"), 44957, "EOMONTH end of Jan 2023");
eq(f("=EDATE(DATE(2023,1,31),1)"), 44985, "EDATE Feb clamp");
eq(f("=DATEDIF(DATE(2020,1,1),DATE(2023,3,1),\"Y\")"), 3, "DATEDIF years");
eq(f("=DATEDIF(DATE(2020,1,1),DATE(2023,3,1),\"M\")"), 38, "DATEDIF months");
eq(f("=DATEDIF(DATE(2020,1,1),DATE(2023,3,15),\"YM\")"), 2, "DATEDIF YM");
eq(f("=DAYS(DATE(2023,1,31),DATE(2023,1,1))"), 30, "DAYS");
eq(f("=HOUR(0.5)"), 12, "HOUR noon");
eq(f("=MINUTE(TIME(8,30,0))"), 30, "MINUTE");
eq(f("=NETWORKDAYS(DATE(2023,1,2),DATE(2023,1,6))"), 5, "NETWORKDAYS full week");
eq(f("=WORKDAY(DATE(2023,1,2),5)"), 44935, "WORKDAY +5");
eq(f('=DATEVALUE("2023-01-01")'), 44927, "DATEVALUE ISO");
eq(f("=YEARFRAC(DATE(2023,1,1),DATE(2023,7,1))"), 0.5, "YEARFRAC 30/360");

// --- Information ---
set("H1", "5"); set("H2", "hello");
eq(f("=ISNUMBER(H1)"), true, "ISNUMBER true");
eq(f("=ISNUMBER(H2)"), false, "ISNUMBER false");
eq(f("=ISTEXT(H2)"), true, "ISTEXT");
eq(f("=ISBLANK(H9)"), true, "ISBLANK on empty");
eq(f("=ISERROR(1/0)"), true, "ISERROR");
eq(f("=ISNA(NA())"), true, "ISNA");
eq(f("=ISEVEN(4)"), true, "ISEVEN");
eq(f("=ISODD(3)"), true, "ISODD");
eq(f("=N(7)"), 7, "N number");
eq(f("=TYPE(H1)"), 1, "TYPE number");
eq(f("=TYPE(H2)"), 2, "TYPE text");
eq(f("=ERROR.TYPE(1/0)"), 2, "ERROR.TYPE div0");

// --- Financial ---
eq(f("=PMT(0.005,360,200000)"), -1199.101049, "PMT monthly payment");
eq(f("=FV(0.005,12,-100,0)"), 1233.555976, "FV");
eq(f("=PV(0.005,12,-100)"), 1161.892213, "PV");
eq(f("=NPV(0.1,100,100,100)") < 250, true, "NPV less than undiscounted sum");
eq(f("=PPMT(0.005,1,360,200000)") + f("=IPMT(0.005,1,360,200000)"), -1199.101049, "PPMT+IPMT = PMT");

console.log(`\n${count - failures}/${count} passed`);
if (failures) {
  console.error(`${failures} FAILED`);
  process.exit(1);
}
