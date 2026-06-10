// Headless smoke test for the conditional-formatting engine. Bundled with
// esbuild and run under Node. Exits non-zero on failure.

import { ConditionalEngine, ConditionalRule } from "../src/engine/conditional.js";
import { CellValue } from "../src/model/types.js";

let failures = 0;
let count = 0;

function eq(actual: unknown, expected: unknown, label: string) {
  count++;
  const ok = actual === expected
    || (typeof actual === "number" && typeof expected === "number" && Math.abs(actual - expected) < 1e-9);
  if (!ok) {
    failures++;
    console.error(`✗ ${label}: expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
  } else {
    console.log(`✓ ${label}`);
  }
}

// Helper: build a getValue from a dense 2D grid.
function gridGet(grid: CellValue[][]): (r: number, c: number) => CellValue {
  return (r, c) => (grid[r] && grid[r][c] !== undefined ? grid[r][c] : null);
}

// A single-column range of numbers 10,20,30,40,50 at rows 0..4, col 0.
const colGrid: CellValue[][] = [[10], [20], [30], [40], [50]];
const colGet = gridGet(colGrid);

// --- greaterThan ---
{
  const rules: ConditionalRule[] = [{
    id: "gt", range: { r0: 0, c0: 0, r1: 4, c1: 0 },
    type: "greaterThan", value1: 25, format: { bold: true },
  }];
  const e = new ConditionalEngine(rules);
  eq(e.resolve(0, 0, 10, colGet), null, "greaterThan: 10 not > 25 → null");
  const v = e.resolve(2, 0, 30, colGet);
  eq(v?.format?.bold, true, "greaterThan: 30 > 25 → bold format");
}

// --- between ---
{
  const rules: ConditionalRule[] = [{
    id: "bt", range: { r0: 0, c0: 0, r1: 4, c1: 0 },
    type: "between", value1: 20, value2: 40, format: { color: "#ff0000" },
  }];
  const e = new ConditionalEngine(rules);
  eq(e.resolve(0, 0, 10, colGet), null, "between: 10 outside [20,40] → null");
  eq(e.resolve(2, 0, 30, colGet)?.format?.color, "#ff0000", "between: 30 inside [20,40]");
  eq(e.resolve(4, 0, 50, colGet), null, "between: 50 outside [20,40] → null");
}

// --- textContains (case-insensitive) ---
{
  const rules: ConditionalRule[] = [{
    id: "tc", range: { r0: 0, c0: 0, r1: 0, c1: 2 },
    type: "textContains", value1: "err", format: { bg: "#ffcccc" },
  }];
  const grid: CellValue[][] = [["Error here", "fine", "ERRoneous"]];
  const g = gridGet(grid);
  const e = new ConditionalEngine(rules);
  eq(e.resolve(0, 0, "Error here", g)?.format?.bg, "#ffcccc", "textContains: matches 'Error'");
  eq(e.resolve(0, 1, "fine", g), null, "textContains: 'fine' no match → null");
  eq(e.resolve(0, 2, "ERRoneous", g)?.format?.bg, "#ffcccc", "textContains: case-insensitive match");
}

// --- duplicate ---
{
  const grid: CellValue[][] = [["a"], ["b"], ["a"], ["c"]];
  const g = gridGet(grid);
  const rules: ConditionalRule[] = [{
    id: "dup", range: { r0: 0, c0: 0, r1: 3, c1: 0 },
    type: "duplicate", format: { italic: true },
  }];
  const e = new ConditionalEngine(rules);
  eq(e.resolve(0, 0, "a", g)?.format?.italic, true, "duplicate: 'a' appears twice → italic");
  eq(e.resolve(1, 0, "b", g), null, "duplicate: 'b' unique → null");
}

// --- top / bottom ---
{
  const rules: ConditionalRule[] = [{
    id: "top2", range: { r0: 0, c0: 0, r1: 4, c1: 0 },
    type: "top", n: 2, format: { bold: true },
  }];
  const e = new ConditionalEngine(rules);
  eq(e.resolve(4, 0, 50, colGet)?.format?.bold, true, "top2: 50 in top 2 → bold");
  eq(e.resolve(3, 0, 40, colGet)?.format?.bold, true, "top2: 40 in top 2 → bold");
  eq(e.resolve(2, 0, 30, colGet), null, "top2: 30 not in top 2 → null");
}
{
  const rules: ConditionalRule[] = [{
    id: "bot1", range: { r0: 0, c0: 0, r1: 4, c1: 0 },
    type: "bottom", n: 1, format: { strike: true },
  }];
  const e = new ConditionalEngine(rules);
  eq(e.resolve(0, 0, 10, colGet)?.format?.strike, true, "bottom1: 10 is lowest → strike");
  eq(e.resolve(1, 0, 20, colGet), null, "bottom1: 20 not lowest → null");
}

// --- colorScale endpoints (2-color) ---
{
  const rules: ConditionalRule[] = [{
    id: "cs", range: { r0: 0, c0: 0, r1: 4, c1: 0 },
    type: "colorScale", minColor: "#000000", maxColor: "#ffffff",
  }];
  const e = new ConditionalEngine(rules);
  eq(e.resolve(0, 0, 10, colGet)?.fillColor, "#000000", "colorScale: min value → minColor");
  eq(e.resolve(4, 0, 50, colGet)?.fillColor, "#ffffff", "colorScale: max value → maxColor");
  // midpoint 30 → halfway → #808080 (128)
  eq(e.resolve(2, 0, 30, colGet)?.fillColor, "#808080", "colorScale: midpoint → grey");
}

// --- dataBar fractions ---
{
  const rules: ConditionalRule[] = [{
    id: "db", range: { r0: 0, c0: 0, r1: 4, c1: 0 },
    type: "dataBar", color: "#00ff00",
  }];
  const e = new ConditionalEngine(rules);
  eq(e.resolve(0, 0, 10, colGet)?.dataBar?.fraction, 0, "dataBar: min → fraction 0");
  eq(e.resolve(4, 0, 50, colGet)?.dataBar?.fraction, 1, "dataBar: max → fraction 1");
  eq(e.resolve(2, 0, 30, colGet)?.dataBar?.fraction, 0.5, "dataBar: midpoint → fraction 0.5");
  eq(e.resolve(2, 0, 30, colGet)?.dataBar?.color, "#00ff00", "dataBar: color propagated");
}

// --- out-of-range cell → null ---
{
  const rules: ConditionalRule[] = [{
    id: "oor", range: { r0: 0, c0: 0, r1: 1, c1: 1 },
    type: "greaterThan", value1: 0, format: { bold: true },
  }];
  const e = new ConditionalEngine(rules);
  eq(e.resolve(5, 5, 100, colGet), null, "out-of-range cell → null");
}

// --- precedence / merge ---
{
  const grid: CellValue[][] = [[10], [20], [30], [40], [50]];
  const g = gridGet(grid);
  const rules: ConditionalRule[] = [
    { id: "r1", range: { r0: 0, c0: 0, r1: 4, c1: 0 }, type: "greaterThan", value1: 0, format: { bold: true, color: "#111111" } },
    { id: "r2", range: { r0: 0, c0: 0, r1: 4, c1: 0 }, type: "greaterThan", value1: 0, format: { color: "#222222" } },
    { id: "r3", range: { r0: 0, c0: 0, r1: 4, c1: 0 }, type: "dataBar", color: "#abcdef" },
  ];
  const e = new ConditionalEngine(rules);
  const v = e.resolve(4, 0, 50, g);
  eq(v?.format?.bold, true, "merge: bold from first rule kept");
  eq(v?.format?.color, "#222222", "merge: later format color overrides earlier");
  eq(v?.dataBar?.fraction, 1, "merge: dataBar combined with format");
}

console.log(`\n${count - failures}/${count} passed`);
if (failures) {
  console.error(`${failures} FAILED`);
  process.exit(1);
}
