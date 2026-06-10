// Built-in function library. Each entry receives the *unevaluated* arg nodes
// plus the evaluator, so functions control evaluation (needed for short-circuit
// logic like IF/IFERROR and for consuming ranges).

import { Node } from "./ast.js";
import { CellValue, CellError } from "../model/types.js";
import { formatA1, parseRef } from "./references.js";
import {
  Evaluator,
  EvalResult,
  RangeValue,
  isRange,
  asScalar,
  toNumber,
  toText,
  toBool,
} from "./evaluator.js";

export type FnImpl = (args: Node[], ev: Evaluator) => EvalResult;

// --- argument helpers ---

function scalar(ev: Evaluator, node: Node): CellValue {
  return asScalar(ev.evalNode(node));
}

// Flatten a list of arg nodes into a flat array of scalar cell values,
// expanding ranges. Used by aggregate functions.
function flatten(ev: Evaluator, args: Node[]): CellValue[] {
  const out: CellValue[] = [];
  for (const a of args) {
    const v = ev.evalNode(a);
    if (isRange(v)) {
      for (const row of v.values) for (const cell of row) out.push(cell);
    } else {
      out.push(v);
    }
  }
  return out;
}

// Numbers only (ignoring blanks/text/bools as Excel aggregates do), but
// propagate the first error encountered.
function numbers(values: CellValue[]): number[] | CellError {
  const nums: number[] = [];
  for (const v of values) {
    if (v instanceof CellError) return v;
    if (typeof v === "number") nums.push(v);
    // strings, booleans, blanks are skipped in range aggregation
  }
  return nums;
}

function num1(ev: Evaluator, node: Node): number | CellError {
  return toNumber(scalar(ev, node));
}

// --- criteria matching for SUMIF / COUNTIF ---

function makeMatcher(criteria: CellValue): (v: CellValue) => boolean {
  const text = toText(criteria).trim();
  const m = /^(<=|>=|<>|<|>|=)?(.*)$/.exec(text);
  const op = m?.[1] ?? "";
  const rhs = m?.[2] ?? "";
  const rhsNum = Number(rhs);
  const isNum = rhs !== "" && !Number.isNaN(rhsNum);

  return (v: CellValue) => {
    if (v instanceof CellError) return false;
    if (op === "" || op === "=") {
      if (isNum && typeof v === "number") return v === rhsNum;
      // wildcard support: * and ?
      if (rhs.includes("*") || rhs.includes("?")) {
        const re = new RegExp(
          "^" + rhs.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".") + "$",
          "i"
        );
        return re.test(toText(v));
      }
      return toText(v).toLowerCase() === rhs.toLowerCase();
    }
    if (op === "<>") return toText(v).toLowerCase() !== rhs.toLowerCase();
    if (isNum && typeof v === "number") {
      switch (op) {
        case "<":
          return v < rhsNum;
        case ">":
          return v > rhsNum;
        case "<=":
          return v <= rhsNum;
        case ">=":
          return v >= rhsNum;
      }
    }
    return false;
  };
}

// --- date serials (Excel-compatible epoch 1899-12-30) ---
const MS_PER_DAY = 86400000;
const EPOCH = Date.UTC(1899, 11, 30);

function dateToSerial(y: number, m: number, d: number): number {
  return Math.round((Date.UTC(y, m - 1, d) - EPOCH) / MS_PER_DAY);
}
function serialToDate(serial: number): Date {
  return new Date(EPOCH + Math.round(serial) * MS_PER_DAY);
}

// function_num → aggregate name, for SUBTOTAL / AGGREGATE.
const SUBTOTAL_FNS: Record<number, string> = {
  1: "AVERAGE", 2: "COUNT", 3: "COUNTA", 4: "MAX", 5: "MIN",
  6: "PRODUCT", 7: "STDEV", 8: "STDEVP", 9: "SUM", 10: "VAR", 11: "VARP",
};
const AGGREGATE_FNS: Record<number, string> = {
  ...SUBTOTAL_FNS,
  12: "MEDIAN", 13: "MODE", 14: "LARGE", 15: "SMALL", 16: "PERCENTILE", 17: "QUARTILE",
};

export const FUNCTIONS: Record<string, FnImpl> = {
  // ===== Math / aggregation =====
  SUM(args, ev) {
    const vals = flatten(ev, args);
    const nums = numbers(vals);
    if (nums instanceof CellError) return nums;
    return nums.reduce((a, b) => a + b, 0);
  },
  AVERAGE(args, ev) {
    const nums = numbers(flatten(ev, args));
    if (nums instanceof CellError) return nums;
    if (nums.length === 0) return new CellError("#DIV/0!");
    return nums.reduce((a, b) => a + b, 0) / nums.length;
  },
  COUNT(args, ev) {
    // COUNT tallies numeric cells and ignores errors/text/blanks.
    return flatten(ev, args).filter((v) => typeof v === "number").length;
  },
  COUNTA(args, ev) {
    return flatten(ev, args).filter((v) => v !== null && v !== "").length;
  },
  COUNTBLANK(args, ev) {
    return flatten(ev, args).filter((v) => v === null || v === "").length;
  },
  MIN(args, ev) {
    const nums = numbers(flatten(ev, args));
    if (nums instanceof CellError) return nums;
    return nums.length ? Math.min(...nums) : 0;
  },
  MAX(args, ev) {
    const nums = numbers(flatten(ev, args));
    if (nums instanceof CellError) return nums;
    return nums.length ? Math.max(...nums) : 0;
  },
  PRODUCT(args, ev) {
    const nums = numbers(flatten(ev, args));
    if (nums instanceof CellError) return nums;
    return nums.length ? nums.reduce((a, b) => a * b, 1) : 0;
  },
  ABS: (a, ev) => unaryMath(a, ev, Math.abs),
  SQRT: (a, ev) =>
    unaryMath(a, ev, (n) => (n < 0 ? new CellError("#NUM!") : Math.sqrt(n))),
  INT: (a, ev) => unaryMath(a, ev, Math.floor),
  SIGN: (a, ev) => unaryMath(a, ev, Math.sign),
  EXP: (a, ev) => unaryMath(a, ev, Math.exp),
  LN: (a, ev) => unaryMath(a, ev, (n) => (n <= 0 ? new CellError("#NUM!") : Math.log(n))),
  LOG10: (a, ev) => unaryMath(a, ev, (n) => (n <= 0 ? new CellError("#NUM!") : Math.log10(n))),
  ROUND(args, ev) {
    const n = num1(ev, args[0]);
    if (n instanceof CellError) return n;
    const d = args[1] ? num1(ev, args[1]) : 0;
    if (d instanceof CellError) return d;
    const f = Math.pow(10, d);
    return Math.round(n * f) / f;
  },
  ROUNDUP(args, ev) {
    const n = num1(ev, args[0]);
    if (n instanceof CellError) return n;
    const d = args[1] ? num1(ev, args[1]) : 0;
    if (d instanceof CellError) return d;
    const f = Math.pow(10, d);
    return (n < 0 ? -Math.ceil(-n * f) : Math.ceil(n * f)) / f;
  },
  ROUNDDOWN(args, ev) {
    const n = num1(ev, args[0]);
    if (n instanceof CellError) return n;
    const d = args[1] ? num1(ev, args[1]) : 0;
    if (d instanceof CellError) return d;
    const f = Math.pow(10, d);
    return (n < 0 ? -Math.floor(-n * f) : Math.floor(n * f)) / f;
  },
  MOD(args, ev) {
    const a = num1(ev, args[0]);
    if (a instanceof CellError) return a;
    const b = num1(ev, args[1]);
    if (b instanceof CellError) return b;
    if (b === 0) return new CellError("#DIV/0!");
    return a - b * Math.floor(a / b);
  },
  POWER(args, ev) {
    const a = num1(ev, args[0]);
    if (a instanceof CellError) return a;
    const b = num1(ev, args[1]);
    if (b instanceof CellError) return b;
    return Math.pow(a, b);
  },
  CEILING(args, ev) {
    const a = num1(ev, args[0]);
    if (a instanceof CellError) return a;
    const sig = args[1] ? num1(ev, args[1]) : 1;
    if (sig instanceof CellError) return sig;
    if (sig === 0) return 0;
    return Math.ceil(a / sig) * sig;
  },
  FLOOR(args, ev) {
    const a = num1(ev, args[0]);
    if (a instanceof CellError) return a;
    const sig = args[1] ? num1(ev, args[1]) : 1;
    if (sig instanceof CellError) return sig;
    if (sig === 0) return new CellError("#DIV/0!");
    return Math.floor(a / sig) * sig;
  },
  PI: () => Math.PI,
  SUMIF(args, ev) {
    const range = asRange(ev, args[0]);
    if (!range) return new CellError("#VALUE!");
    const crit = scalar(ev, args[1]);
    const match = makeMatcher(crit);
    const sumRange = args[2] ? asRange(ev, args[2]) : range;
    const flat = range.values.flat();
    const sumFlat = sumRange ? sumRange.values.flat() : flat;
    let total = 0;
    for (let i = 0; i < flat.length; i++) {
      if (match(flat[i])) {
        const v = sumFlat[i];
        if (typeof v === "number") total += v;
      }
    }
    return total;
  },
  COUNTIF(args, ev) {
    const range = asRange(ev, args[0]);
    if (!range) return new CellError("#VALUE!");
    const match = makeMatcher(scalar(ev, args[1]));
    return range.values.flat().filter(match).length;
  },
  AVERAGEIF(args, ev) {
    const range = asRange(ev, args[0]);
    if (!range) return new CellError("#VALUE!");
    const match = makeMatcher(scalar(ev, args[1]));
    const avgRange = args[2] ? asRange(ev, args[2]) : range;
    const flat = range.values.flat();
    const avgFlat = avgRange ? avgRange.values.flat() : flat;
    let total = 0;
    let count = 0;
    for (let i = 0; i < flat.length; i++) {
      if (match(flat[i]) && typeof avgFlat[i] === "number") {
        total += avgFlat[i] as number;
        count++;
      }
    }
    return count === 0 ? new CellError("#DIV/0!") : total / count;
  },
  SUMPRODUCT(args, ev) {
    const ranges = args.map((a) => asRange(ev, a));
    if (ranges.some((r) => !r)) return new CellError("#VALUE!");
    const arrays = (ranges as RangeValue[]).map((r) => r.values.flat());
    const len = arrays[0].length;
    if (arrays.some((a) => a.length !== len)) return new CellError("#VALUE!");
    let total = 0;
    for (let i = 0; i < len; i++) {
      let prod = 1;
      for (const arr of arrays) prod *= typeof arr[i] === "number" ? (arr[i] as number) : 0;
      total += prod;
    }
    return total;
  },

  // ===== Logical =====
  IF(args, ev) {
    const cond = toBool(scalar(ev, args[0]));
    if (cond instanceof CellError) return cond;
    if (cond) return ev.evalNode(args[1]);
    return args[2] ? ev.evalNode(args[2]) : false;
  },
  IFERROR(args, ev) {
    const v = ev.evalNode(args[0]);
    if (asScalar(v) instanceof CellError) return ev.evalNode(args[1]);
    return v;
  },
  IFNA(args, ev) {
    const v = ev.evalNode(args[0]);
    const s = asScalar(v);
    if (s instanceof CellError && s.kind === "#N/A") return ev.evalNode(args[1]);
    return v;
  },
  IFS(args, ev) {
    for (let i = 0; i + 1 < args.length; i += 2) {
      const cond = toBool(scalar(ev, args[i]));
      if (cond instanceof CellError) return cond;
      if (cond) return ev.evalNode(args[i + 1]);
    }
    return new CellError("#N/A");
  },
  AND(args, ev) {
    for (const v of flatten(ev, args)) {
      const b = toBool(v);
      if (b instanceof CellError) return b;
      if (!b) return false;
    }
    return true;
  },
  OR(args, ev) {
    for (const v of flatten(ev, args)) {
      const b = toBool(v);
      if (b instanceof CellError) return b;
      if (b) return true;
    }
    return false;
  },
  XOR(args, ev) {
    let count = 0;
    for (const v of flatten(ev, args)) {
      const b = toBool(v);
      if (b instanceof CellError) return b;
      if (b) count++;
    }
    return count % 2 === 1;
  },
  NOT(args, ev) {
    const b = toBool(scalar(ev, args[0]));
    return b instanceof CellError ? b : !b;
  },
  TRUE: () => true,
  FALSE: () => false,
  NA: () => new CellError("#N/A"),

  // ===== Text =====
  CONCATENATE: (a, ev) => flatten(ev, a).map(toText).join(""),
  CONCAT: (a, ev) => flatten(ev, a).map(toText).join(""),
  TEXTJOIN(args, ev) {
    const delim = toText(scalar(ev, args[0]));
    const skipEmpty = toBool(scalar(ev, args[1]));
    if (skipEmpty instanceof CellError) return skipEmpty;
    const parts = flatten(ev, args.slice(2))
      .filter((v) => (skipEmpty ? v !== null && v !== "" : true))
      .map(toText);
    return parts.join(delim);
  },
  LEN: (a, ev) => toText(scalar(ev, a[0])).length,
  LOWER: (a, ev) => toText(scalar(ev, a[0])).toLowerCase(),
  UPPER: (a, ev) => toText(scalar(ev, a[0])).toUpperCase(),
  TRIM: (a, ev) => toText(scalar(ev, a[0])).replace(/\s+/g, " ").trim(),
  PROPER: (a, ev) =>
    toText(scalar(ev, a[0])).replace(/\b\w/g, (c) => c.toUpperCase()).replace(/\B\w/g, (c) => c.toLowerCase()),
  LEFT(args, ev) {
    const s = toText(scalar(ev, args[0]));
    const n = args[1] ? num1(ev, args[1]) : 1;
    if (n instanceof CellError) return n;
    return s.slice(0, Math.max(0, n));
  },
  RIGHT(args, ev) {
    const s = toText(scalar(ev, args[0]));
    const n = args[1] ? num1(ev, args[1]) : 1;
    if (n instanceof CellError) return n;
    return n <= 0 ? "" : s.slice(-n);
  },
  MID(args, ev) {
    const s = toText(scalar(ev, args[0]));
    const start = num1(ev, args[1]);
    if (start instanceof CellError) return start;
    const len = num1(ev, args[2]);
    if (len instanceof CellError) return len;
    return s.slice(Math.max(0, start - 1), Math.max(0, start - 1) + Math.max(0, len));
  },
  REPT(args, ev) {
    const s = toText(scalar(ev, args[0]));
    const n = num1(ev, args[1]);
    if (n instanceof CellError) return n;
    return n > 0 ? s.repeat(Math.floor(n)) : "";
  },
  SUBSTITUTE(args, ev) {
    const s = toText(scalar(ev, args[0]));
    const oldT = toText(scalar(ev, args[1]));
    const newT = toText(scalar(ev, args[2]));
    if (oldT === "") return s;
    if (args[3]) {
      const which = num1(ev, args[3]);
      if (which instanceof CellError) return which;
      let idx = -1;
      let count = 0;
      let from = 0;
      while ((idx = s.indexOf(oldT, from)) !== -1) {
        count++;
        if (count === which) return s.slice(0, idx) + newT + s.slice(idx + oldT.length);
        from = idx + oldT.length;
      }
      return s;
    }
    return s.split(oldT).join(newT);
  },
  REPLACE(args, ev) {
    const s = toText(scalar(ev, args[0]));
    const start = num1(ev, args[1]);
    if (start instanceof CellError) return start;
    const len = num1(ev, args[2]);
    if (len instanceof CellError) return len;
    const newT = toText(scalar(ev, args[3]));
    return s.slice(0, start - 1) + newT + s.slice(start - 1 + len);
  },
  FIND(args, ev) {
    const find = toText(scalar(ev, args[0]));
    const within = toText(scalar(ev, args[1]));
    const start = args[2] ? num1(ev, args[2]) : 1;
    if (start instanceof CellError) return start;
    const idx = within.indexOf(find, start - 1);
    return idx === -1 ? new CellError("#VALUE!") : idx + 1;
  },
  SEARCH(args, ev) {
    const find = toText(scalar(ev, args[0])).toLowerCase();
    const within = toText(scalar(ev, args[1])).toLowerCase();
    const start = args[2] ? num1(ev, args[2]) : 1;
    if (start instanceof CellError) return start;
    const idx = within.indexOf(find, start - 1);
    return idx === -1 ? new CellError("#VALUE!") : idx + 1;
  },
  VALUE(args, ev) {
    return toNumber(scalar(ev, args[0]));
  },
  TEXT(args, ev) {
    const v = scalar(ev, args[0]);
    const fmt = toText(scalar(ev, args[1]));
    // Delegated to the shared number formatter at call sites that have it; here
    // we do a minimal subset to stay dependency-free in the engine.
    const n = toNumber(v);
    if (n instanceof CellError) return toText(v);
    if (fmt.includes("%")) return (n * 100).toFixed((fmt.split(".")[1] || "").length) + "%";
    const decimals = (fmt.split(".")[1] || "").replace(/[^0#]/g, "").length;
    if (fmt.includes("0") || fmt.includes("#")) {
      const s = n.toFixed(decimals);
      return fmt.includes(",") ? addThousands(s) : s;
    }
    return toText(v);
  },

  // ===== Lookup =====
  VLOOKUP(args, ev) {
    const key = scalar(ev, args[0]);
    const table = asRange(ev, args[1]);
    if (!table) return new CellError("#VALUE!");
    const colIdx = num1(ev, args[2]);
    if (colIdx instanceof CellError) return colIdx;
    const approx = args[3] ? toBool(scalar(ev, args[3])) : true;
    if (approx instanceof CellError) return approx;
    const ci = colIdx - 1;
    if (ci < 0 || ci >= table.values[0].length) return new CellError("#REF!");
    if (approx) {
      let found: CellValue = new CellError("#N/A");
      for (const row of table.values) {
        if (compareLE(row[0], key)) found = row[ci];
        else break;
      }
      return found;
    }
    for (const row of table.values) {
      if (looseEqual(row[0], key)) return row[ci];
    }
    return new CellError("#N/A");
  },
  HLOOKUP(args, ev) {
    const key = scalar(ev, args[0]);
    const table = asRange(ev, args[1]);
    if (!table) return new CellError("#VALUE!");
    const rowIdx = num1(ev, args[2]);
    if (rowIdx instanceof CellError) return rowIdx;
    const ri = rowIdx - 1;
    if (ri < 0 || ri >= table.values.length) return new CellError("#REF!");
    const header = table.values[0];
    for (let c = 0; c < header.length; c++) {
      if (looseEqual(header[c], key)) return table.values[ri][c];
    }
    return new CellError("#N/A");
  },
  INDEX(args, ev) {
    const range = asRange(ev, args[0]);
    if (!range) return new CellError("#VALUE!");
    const r = num1(ev, args[1]);
    if (r instanceof CellError) return r;
    const c = args[2] ? num1(ev, args[2]) : 1;
    if (c instanceof CellError) return c;
    const rows = range.values;
    // single-row or single-column ranges accept a single index
    if (rows.length === 1) return rows[0][(c || r) - 1] ?? new CellError("#REF!");
    if (rows[0].length === 1 && !args[2]) return rows[r - 1]?.[0] ?? new CellError("#REF!");
    const rr = rows[r - 1];
    if (!rr) return new CellError("#REF!");
    return rr[c - 1] ?? new CellError("#REF!");
  },
  MATCH(args, ev) {
    const key = scalar(ev, args[0]);
    const range = asRange(ev, args[1]);
    if (!range) return new CellError("#VALUE!");
    const type = args[2] ? num1(ev, args[2]) : 1;
    if (type instanceof CellError) return type;
    const flat = range.values.flat();
    if (type === 0) {
      for (let i = 0; i < flat.length; i++) if (looseEqual(flat[i], key)) return i + 1;
      return new CellError("#N/A");
    }
    // approximate: largest value <= key (type 1) assuming ascending
    let result = -1;
    for (let i = 0; i < flat.length; i++) {
      if (type === 1 && compareLE(flat[i], key)) result = i + 1;
      if (type === -1 && compareLE(key, flat[i])) result = i + 1;
    }
    return result === -1 ? new CellError("#N/A") : result;
  },

  // ===== Date / time =====
  TODAY: () => {
    const now = new Date();
    return dateToSerial(now.getFullYear(), now.getMonth() + 1, now.getDate());
  },
  NOW: () => {
    const now = new Date();
    const days = dateToSerial(now.getFullYear(), now.getMonth() + 1, now.getDate());
    const frac = (now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds()) / 86400;
    return days + frac;
  },
  DATE(args, ev) {
    const y = num1(ev, args[0]);
    if (y instanceof CellError) return y;
    const m = num1(ev, args[1]);
    if (m instanceof CellError) return m;
    const d = num1(ev, args[2]);
    if (d instanceof CellError) return d;
    return dateToSerial(y, m, d);
  },
  YEAR: (a, ev) => datePart(a, ev, (d) => d.getUTCFullYear()),
  MONTH: (a, ev) => datePart(a, ev, (d) => d.getUTCMonth() + 1),
  DAY: (a, ev) => datePart(a, ev, (d) => d.getUTCDate()),
  WEEKDAY(args, ev) {
    const n = num1(ev, args[0]);
    if (n instanceof CellError) return n;
    return serialToDate(n).getUTCDay() + 1;
  },

  // ===== Added: Math / trig =====
  TRUNC(args, ev) {
    const n = num1(ev, args[0]);
    if (n instanceof CellError) return n;
    const d = args[1] ? num1(ev, args[1]) : 0;
    if (d instanceof CellError) return d;
    const f = Math.pow(10, d);
    return Math.trunc(n * f) / f;
  },
  EVEN(args, ev) {
    const n = num1(ev, args[0]);
    if (n instanceof CellError) return n;
    const up = Math.ceil(Math.abs(n) / 2) * 2;
    return n < 0 ? -up : up;
  },
  ODD(args, ev) {
    const n = num1(ev, args[0]);
    if (n instanceof CellError) return n;
    let up = Math.ceil(Math.abs(n));
    if (up % 2 === 0) up += 1;
    if (up === 0) up = 1;
    return n < 0 ? -up : up;
  },
  QUOTIENT(args, ev) {
    const a = num1(ev, args[0]);
    if (a instanceof CellError) return a;
    const b = num1(ev, args[1]);
    if (b instanceof CellError) return b;
    if (b === 0) return new CellError("#DIV/0!");
    return Math.trunc(a / b);
  },
  GCD(args, ev) {
    const nums = numbers(flatten(ev, args));
    if (nums instanceof CellError) return nums;
    if (nums.some((n) => n < 0)) return new CellError("#NUM!");
    const ints = nums.map((n) => Math.floor(n));
    return ints.reduce((a, b) => gcd2(a, b), 0);
  },
  LCM(args, ev) {
    const nums = numbers(flatten(ev, args));
    if (nums instanceof CellError) return nums;
    if (nums.some((n) => n < 0)) return new CellError("#NUM!");
    const ints = nums.map((n) => Math.floor(n));
    return ints.reduce((a, b) => {
      if (a === 0 || b === 0) return 0;
      return Math.abs(a * b) / gcd2(a, b);
    }, 1);
  },
  FACT(args, ev) {
    const n = num1(ev, args[0]);
    if (n instanceof CellError) return n;
    const f = Math.floor(n);
    if (f < 0) return new CellError("#NUM!");
    return factorial(f);
  },
  FACTDOUBLE(args, ev) {
    const n = num1(ev, args[0]);
    if (n instanceof CellError) return n;
    const f = Math.floor(n);
    if (f < -1) return new CellError("#NUM!");
    let result = 1;
    for (let i = f; i > 1; i -= 2) result *= i;
    return result;
  },
  COMBIN(args, ev) {
    const n = num1(ev, args[0]);
    if (n instanceof CellError) return n;
    const k = num1(ev, args[1]);
    if (k instanceof CellError) return k;
    const ni = Math.floor(n);
    const ki = Math.floor(k);
    if (ni < 0 || ki < 0 || ki > ni) return new CellError("#NUM!");
    return Math.round(factorial(ni) / (factorial(ki) * factorial(ni - ki)));
  },
  PERMUT(args, ev) {
    const n = num1(ev, args[0]);
    if (n instanceof CellError) return n;
    const k = num1(ev, args[1]);
    if (k instanceof CellError) return k;
    const ni = Math.floor(n);
    const ki = Math.floor(k);
    if (ni < 0 || ki < 0 || ki > ni) return new CellError("#NUM!");
    return Math.round(factorial(ni) / factorial(ni - ki));
  },
  SUMSQ(args, ev) {
    const nums = numbers(flatten(ev, args));
    if (nums instanceof CellError) return nums;
    return nums.reduce((a, b) => a + b * b, 0);
  },
  MROUND(args, ev) {
    const n = num1(ev, args[0]);
    if (n instanceof CellError) return n;
    const m = num1(ev, args[1]);
    if (m instanceof CellError) return m;
    if (m === 0) return 0;
    if ((n < 0 && m > 0) || (n > 0 && m < 0)) return new CellError("#NUM!");
    return Math.round(n / m) * m;
  },
  SQRTPI(args, ev) {
    const n = num1(ev, args[0]);
    if (n instanceof CellError) return n;
    if (n < 0) return new CellError("#NUM!");
    return Math.sqrt(n * Math.PI);
  },
  RADIANS: (a, ev) => unaryMath(a, ev, (n) => (n * Math.PI) / 180),
  DEGREES: (a, ev) => unaryMath(a, ev, (n) => (n * 180) / Math.PI),
  SIN: (a, ev) => unaryMath(a, ev, Math.sin),
  COS: (a, ev) => unaryMath(a, ev, Math.cos),
  TAN: (a, ev) => unaryMath(a, ev, Math.tan),
  ASIN: (a, ev) => unaryMath(a, ev, (n) => (n < -1 || n > 1 ? new CellError("#NUM!") : Math.asin(n))),
  ACOS: (a, ev) => unaryMath(a, ev, (n) => (n < -1 || n > 1 ? new CellError("#NUM!") : Math.acos(n))),
  ATAN: (a, ev) => unaryMath(a, ev, Math.atan),
  ATAN2(args, ev) {
    const x = num1(ev, args[0]);
    if (x instanceof CellError) return x;
    const y = num1(ev, args[1]);
    if (y instanceof CellError) return y;
    if (x === 0 && y === 0) return new CellError("#DIV/0!");
    return Math.atan2(y, x);
  },
  SINH: (a, ev) => unaryMath(a, ev, Math.sinh),
  COSH: (a, ev) => unaryMath(a, ev, Math.cosh),
  TANH: (a, ev) => unaryMath(a, ev, Math.tanh),
  LOG(args, ev) {
    const n = num1(ev, args[0]);
    if (n instanceof CellError) return n;
    if (n <= 0) return new CellError("#NUM!");
    const base = args[1] ? num1(ev, args[1]) : 10;
    if (base instanceof CellError) return base;
    if (base <= 0 || base === 1) return new CellError("#NUM!");
    return Math.log(n) / Math.log(base);
  },
  RAND: () => Math.random(),
  RANDBETWEEN(args, ev) {
    const lo = num1(ev, args[0]);
    if (lo instanceof CellError) return lo;
    const hi = num1(ev, args[1]);
    if (hi instanceof CellError) return hi;
    const l = Math.ceil(lo);
    const h = Math.floor(hi);
    if (l > h) return new CellError("#NUM!");
    return l + Math.floor(Math.random() * (h - l + 1));
  },
  SUMIFS(args, ev) {
    return ifsAggregate(args, ev, 1, "sum");
  },
  COUNTIFS(args, ev) {
    return ifsAggregate(args, ev, 0, "count");
  },
  AVERAGEIFS(args, ev) {
    return ifsAggregate(args, ev, 1, "avg");
  },
  MAXIFS(args, ev) {
    return ifsAggregate(args, ev, 1, "max");
  },
  MINIFS(args, ev) {
    return ifsAggregate(args, ev, 1, "min");
  },

  // ===== Added: Statistical =====
  MEDIAN(args, ev) {
    const nums = numbers(flatten(ev, args));
    if (nums instanceof CellError) return nums;
    if (nums.length === 0) return new CellError("#NUM!");
    const s = [...nums].sort((a, b) => a - b);
    const mid = Math.floor(s.length / 2);
    return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
  },
  MODE(args, ev) {
    return modeImpl(args, ev);
  },
  "MODE.SNGL"(args, ev) {
    return modeImpl(args, ev);
  },
  STDEV(args, ev) {
    return varianceImpl(args, ev, true, true);
  },
  STDEVP(args, ev) {
    return varianceImpl(args, ev, false, true);
  },
  VAR(args, ev) {
    return varianceImpl(args, ev, true, false);
  },
  VARP(args, ev) {
    return varianceImpl(args, ev, false, false);
  },
  LARGE(args, ev) {
    const nums = numbers(flatten(ev, [args[0]]));
    if (nums instanceof CellError) return nums;
    const k = num1(ev, args[1]);
    if (k instanceof CellError) return k;
    const ki = Math.floor(k);
    if (ki < 1 || ki > nums.length) return new CellError("#NUM!");
    const s = [...nums].sort((a, b) => b - a);
    return s[ki - 1];
  },
  SMALL(args, ev) {
    const nums = numbers(flatten(ev, [args[0]]));
    if (nums instanceof CellError) return nums;
    const k = num1(ev, args[1]);
    if (k instanceof CellError) return k;
    const ki = Math.floor(k);
    if (ki < 1 || ki > nums.length) return new CellError("#NUM!");
    const s = [...nums].sort((a, b) => a - b);
    return s[ki - 1];
  },
  RANK(args, ev) {
    const v = num1(ev, args[0]);
    if (v instanceof CellError) return v;
    const nums = numbers(flatten(ev, [args[1]]));
    if (nums instanceof CellError) return nums;
    const order = args[2] ? num1(ev, args[2]) : 0;
    if (order instanceof CellError) return order;
    const asc = order !== 0;
    const s = [...nums].sort((a, b) => (asc ? a - b : b - a));
    const idx = s.indexOf(v);
    return idx === -1 ? new CellError("#N/A") : idx + 1;
  },
  PERCENTILE(args, ev) {
    const nums = numbers(flatten(ev, [args[0]]));
    if (nums instanceof CellError) return nums;
    const k = num1(ev, args[1]);
    if (k instanceof CellError) return k;
    return percentile(nums, k);
  },
  QUARTILE(args, ev) {
    const nums = numbers(flatten(ev, [args[0]]));
    if (nums instanceof CellError) return nums;
    const q = num1(ev, args[1]);
    if (q instanceof CellError) return q;
    const qi = Math.floor(q);
    if (qi < 0 || qi > 4) return new CellError("#NUM!");
    return percentile(nums, qi / 4);
  },
  AVERAGEA(args, ev) {
    const vals = flatten(ev, args);
    let total = 0;
    let count = 0;
    for (const v of vals) {
      if (v instanceof CellError) return v;
      if (v === null) continue;
      if (typeof v === "number") total += v;
      else if (typeof v === "boolean") total += v ? 1 : 0;
      else if (typeof v === "string") {
        if (v === "") continue;
        total += 0;
      }
      count++;
    }
    return count === 0 ? new CellError("#DIV/0!") : total / count;
  },
  GEOMEAN(args, ev) {
    const nums = numbers(flatten(ev, args));
    if (nums instanceof CellError) return nums;
    if (nums.length === 0 || nums.some((n) => n <= 0)) return new CellError("#NUM!");
    const logSum = nums.reduce((a, b) => a + Math.log(b), 0);
    return Math.exp(logSum / nums.length);
  },
  HARMEAN(args, ev) {
    const nums = numbers(flatten(ev, args));
    if (nums instanceof CellError) return nums;
    if (nums.length === 0 || nums.some((n) => n <= 0)) return new CellError("#NUM!");
    const recip = nums.reduce((a, b) => a + 1 / b, 0);
    return nums.length / recip;
  },
  DEVSQ(args, ev) {
    const nums = numbers(flatten(ev, args));
    if (nums instanceof CellError) return nums;
    if (nums.length === 0) return new CellError("#NUM!");
    const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
    return nums.reduce((a, b) => a + (b - mean) * (b - mean), 0);
  },
  AVEDEV(args, ev) {
    const nums = numbers(flatten(ev, args));
    if (nums instanceof CellError) return nums;
    if (nums.length === 0) return new CellError("#NUM!");
    const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
    return nums.reduce((a, b) => a + Math.abs(b - mean), 0) / nums.length;
  },

  // ===== Added: Text =====
  CHAR(args, ev) {
    const n = num1(ev, args[0]);
    if (n instanceof CellError) return n;
    const c = Math.floor(n);
    if (c < 1 || c > 255) return new CellError("#VALUE!");
    return String.fromCharCode(c);
  },
  CODE(args, ev) {
    const s = toText(scalar(ev, args[0]));
    if (s.length === 0) return new CellError("#VALUE!");
    return s.charCodeAt(0);
  },
  UNICHAR(args, ev) {
    const n = num1(ev, args[0]);
    if (n instanceof CellError) return n;
    const c = Math.floor(n);
    if (c < 1) return new CellError("#VALUE!");
    return String.fromCodePoint(c);
  },
  UNICODE(args, ev) {
    const s = toText(scalar(ev, args[0]));
    if (s.length === 0) return new CellError("#VALUE!");
    return s.codePointAt(0) ?? new CellError("#VALUE!");
  },
  EXACT(args, ev) {
    return toText(scalar(ev, args[0])) === toText(scalar(ev, args[1]));
  },
  T(args, ev) {
    const v = scalar(ev, args[0]);
    if (v instanceof CellError) return v;
    return typeof v === "string" ? v : "";
  },
  CLEAN(args, ev) {
    const s = toText(scalar(ev, args[0]));
    // eslint-disable-next-line no-control-regex
    return s.replace(/[\x00-\x1f]/g, "");
  },
  DOLLAR(args, ev) {
    const n = num1(ev, args[0]);
    if (n instanceof CellError) return n;
    const d = args[1] ? num1(ev, args[1]) : 2;
    if (d instanceof CellError) return d;
    const dec = Math.max(0, Math.floor(d));
    const f = Math.pow(10, dec);
    const rounded = Math.round(n * f) / f;
    const sign = rounded < 0 ? "-" : "";
    return sign + "$" + addThousands(Math.abs(rounded).toFixed(dec));
  },
  FIXED(args, ev) {
    const n = num1(ev, args[0]);
    if (n instanceof CellError) return n;
    const d = args[1] ? num1(ev, args[1]) : 2;
    if (d instanceof CellError) return d;
    const noCommas = args[2] ? toBool(scalar(ev, args[2])) : false;
    if (noCommas instanceof CellError) return noCommas;
    const dec = Math.max(0, Math.floor(d));
    const s = n.toFixed(dec);
    return noCommas ? s : addThousands(s);
  },
  NUMBERVALUE(args, ev) {
    let s = toText(scalar(ev, args[0]));
    const dec = args[1] ? toText(scalar(ev, args[1])) : ".";
    const grp = args[2] ? toText(scalar(ev, args[2])) : ",";
    if (grp) s = s.split(grp).join("");
    if (dec && dec !== ".") s = s.split(dec).join(".");
    s = s.trim();
    if (s === "") return 0;
    const n = Number(s);
    return Number.isNaN(n) ? new CellError("#VALUE!") : n;
  },
  TEXTBEFORE(args, ev) {
    const s = toText(scalar(ev, args[0]));
    const delim = toText(scalar(ev, args[1]));
    if (delim === "") return "";
    const idx = s.indexOf(delim);
    return idx === -1 ? new CellError("#N/A") : s.slice(0, idx);
  },
  TEXTAFTER(args, ev) {
    const s = toText(scalar(ev, args[0]));
    const delim = toText(scalar(ev, args[1]));
    if (delim === "") return s;
    const idx = s.indexOf(delim);
    return idx === -1 ? new CellError("#N/A") : s.slice(idx + delim.length);
  },

  // ===== Added: Logical =====
  SWITCH(args, ev) {
    const target = scalar(ev, args[0]);
    if (target instanceof CellError) return target;
    let i = 1;
    for (; i + 1 < args.length; i += 2) {
      const candidate = scalar(ev, args[i]);
      if (looseEqual(candidate, target)) return ev.evalNode(args[i + 1]);
    }
    // trailing odd arg is the default
    if (i < args.length) return ev.evalNode(args[i]);
    return new CellError("#N/A");
  },

  // ===== Added: Lookup / reference =====
  CHOOSE(args, ev) {
    const idx = num1(ev, args[0]);
    if (idx instanceof CellError) return idx;
    const i = Math.floor(idx);
    if (i < 1 || i >= args.length) return new CellError("#VALUE!");
    return ev.evalNode(args[i]);
  },
  XLOOKUP(args, ev) {
    const key = scalar(ev, args[0]);
    const lookup = asRange(ev, args[1]);
    const ret = asRange(ev, args[2]);
    if (!lookup || !ret) return new CellError("#VALUE!");
    const lookFlat = lookup.values.flat();
    const retVals = ret.values;
    const retFlat = retVals.flat();
    for (let i = 0; i < lookFlat.length; i++) {
      if (looseEqual(lookFlat[i], key)) {
        return retFlat[i] ?? new CellError("#N/A");
      }
    }
    if (args[3]) return ev.evalNode(args[3]);
    return new CellError("#N/A");
  },
  LOOKUP(args, ev) {
    const key = scalar(ev, args[0]);
    const lookup = asRange(ev, args[1]);
    if (!lookup) return new CellError("#VALUE!");
    const lookFlat = lookup.values.flat();
    const result = args[2] ? asRange(ev, args[2]) : lookup;
    const resFlat = result ? result.values.flat() : lookFlat;
    // approximate match: largest value <= key (assumes ascending)
    let found = -1;
    for (let i = 0; i < lookFlat.length; i++) {
      if (compareLE(lookFlat[i], key)) found = i;
      else break;
    }
    return found === -1 ? new CellError("#N/A") : resFlat[found] ?? new CellError("#N/A");
  },
  ROWS(args, ev) {
    const r = asRange(ev, args[0]);
    if (!r) return new CellError("#VALUE!");
    return r.values.length;
  },
  COLUMNS(args, ev) {
    const r = asRange(ev, args[0]);
    if (!r) return new CellError("#VALUE!");
    return r.values[0]?.length ?? 0;
  },

  // ===== Reference functions =====
  ROW(args, ev) {
    if (!args[0]) {
      const cur = ev.current();
      return cur ? cur.row + 1 : new CellError("#VALUE!");
    }
    const n = args[0];
    if (n.kind === "ref") return n.ref.row + 1;
    if (n.kind === "range") return Math.min(n.start.row, n.end.row) + 1;
    return new CellError("#VALUE!");
  },
  COLUMN(args, ev) {
    if (!args[0]) {
      const cur = ev.current();
      return cur ? cur.col + 1 : new CellError("#VALUE!");
    }
    const n = args[0];
    if (n.kind === "ref") return n.ref.col + 1;
    if (n.kind === "range") return Math.min(n.start.col, n.end.col) + 1;
    return new CellError("#VALUE!");
  },
  ADDRESS(args, ev) {
    const r = num1(ev, args[0]);
    if (r instanceof CellError) return r;
    const c = num1(ev, args[1]);
    if (c instanceof CellError) return c;
    const absNum = args[2] ? num1(ev, args[2]) : 1;
    if (absNum instanceof CellError) return absNum;
    const absRow = absNum === 1 || absNum === 2;
    const absCol = absNum === 1 || absNum === 3;
    return formatA1(r - 1, c - 1, absRow, absCol);
  },
  INDIRECT(args, ev) {
    const text = toText(scalar(ev, args[0])).trim();
    const [a, b] = text.split(":");
    const p0 = parseRef(a);
    if (!p0) return new CellError("#REF!");
    if (b) {
      const p1 = parseRef(b);
      if (!p1) return new CellError("#REF!");
      return ev.evalNode({ kind: "range", sheet: p0.sheet, start: p0.ref, end: p1.ref });
    }
    return ev.evalNode({ kind: "ref", sheet: p0.sheet, ref: p0.ref });
  },
  OFFSET(args, ev) {
    const base = args[0];
    let br: number;
    let bc: number;
    let sheet: string | undefined;
    if (base.kind === "ref") { br = base.ref.row; bc = base.ref.col; sheet = base.sheet; }
    else if (base.kind === "range") { br = Math.min(base.start.row, base.end.row); bc = Math.min(base.start.col, base.end.col); sheet = base.sheet; }
    else return new CellError("#REF!");
    const dr = num1(ev, args[1]);
    if (dr instanceof CellError) return dr;
    const dc = num1(ev, args[2]);
    if (dc instanceof CellError) return dc;
    const h = args[3] ? num1(ev, args[3]) : 1;
    if (h instanceof CellError) return h;
    const w = args[4] ? num1(ev, args[4]) : 1;
    if (w instanceof CellError) return w;
    const r0 = br + dr;
    const c0 = bc + dc;
    if (r0 < 0 || c0 < 0 || h < 1 || w < 1) return new CellError("#REF!");
    const mk = (row: number, col: number) => ({ row, col, absRow: false, absCol: false });
    if (h === 1 && w === 1) return ev.evalNode({ kind: "ref", sheet, ref: mk(r0, c0) });
    return ev.evalNode({ kind: "range", sheet, start: mk(r0, c0), end: mk(r0 + h - 1, c0 + w - 1) });
  },
  SUBTOTAL(args, ev) {
    const fn = num1(ev, args[0]);
    if (fn instanceof CellError) return fn;
    const base = fn > 100 ? fn - 100 : fn;
    const name = SUBTOTAL_FNS[base];
    if (!name) return new CellError("#VALUE!");
    return FUNCTIONS[name](args.slice(1), ev);
  },
  AGGREGATE(args, ev) {
    const fn = num1(ev, args[0]);
    if (fn instanceof CellError) return fn;
    const name = AGGREGATE_FNS[fn];
    if (!name) return new CellError("#VALUE!");
    // AGGREGATE(funcNum, options, range, [k]) — drop funcNum + options
    return FUNCTIONS[name](args.slice(2), ev);
  },

  // ===== Added: Date / time =====
  HOUR(args, ev) {
    const n = num1(ev, args[0]);
    if (n instanceof CellError) return n;
    return Math.floor(timeFrac(n) * 24) % 24;
  },
  MINUTE(args, ev) {
    const n = num1(ev, args[0]);
    if (n instanceof CellError) return n;
    return Math.floor(timeFrac(n) * 1440) % 60;
  },
  SECOND(args, ev) {
    const n = num1(ev, args[0]);
    if (n instanceof CellError) return n;
    return Math.round(timeFrac(n) * 86400) % 60;
  },
  TIME(args, ev) {
    const h = num1(ev, args[0]);
    if (h instanceof CellError) return h;
    const m = num1(ev, args[1]);
    if (m instanceof CellError) return m;
    const s = num1(ev, args[2]);
    if (s instanceof CellError) return s;
    const frac = ((h * 3600 + m * 60 + s) % 86400) / 86400;
    return frac < 0 ? frac + 1 : frac;
  },
  EDATE(args, ev) {
    const start = num1(ev, args[0]);
    if (start instanceof CellError) return start;
    const months = num1(ev, args[1]);
    if (months instanceof CellError) return months;
    return edate(start, Math.floor(months));
  },
  EOMONTH(args, ev) {
    const start = num1(ev, args[0]);
    if (start instanceof CellError) return start;
    const months = num1(ev, args[1]);
    if (months instanceof CellError) return months;
    const d = serialToDate(start);
    const y = d.getUTCFullYear();
    const m = d.getUTCMonth() + 1 + Math.floor(months);
    // day 0 of next month = last day of target month
    return dateToSerial(y, m + 1, 0);
  },
  DATEDIF(args, ev) {
    const start = num1(ev, args[0]);
    if (start instanceof CellError) return start;
    const end = num1(ev, args[1]);
    if (end instanceof CellError) return end;
    const unit = toText(scalar(ev, args[2])).toUpperCase();
    if (end < start) return new CellError("#NUM!");
    const sd = serialToDate(start);
    const ed = serialToDate(end);
    const sy = sd.getUTCFullYear(), sm = sd.getUTCMonth() + 1, sday = sd.getUTCDate();
    const ey = ed.getUTCFullYear(), em = ed.getUTCMonth() + 1, eday = ed.getUTCDate();
    switch (unit) {
      case "Y": {
        let y = ey - sy;
        if (em < sm || (em === sm && eday < sday)) y--;
        return y;
      }
      case "M": {
        let m = (ey - sy) * 12 + (em - sm);
        if (eday < sday) m--;
        return m;
      }
      case "D":
        return Math.floor(end) - Math.floor(start);
      case "MD": {
        let d = eday - sday;
        if (d < 0) {
          // borrow days from the month preceding the end date's month
          const daysPrevMonth = new Date(Date.UTC(ey, em - 1, 0)).getUTCDate();
          d = daysPrevMonth - sday + eday;
        }
        return d;
      }
      case "YM": {
        let m = (em - sm);
        if (eday < sday) m--;
        if (m < 0) m += 12;
        return m;
      }
      case "YD": {
        // difference in days ignoring years
        let anchor = dateToSerial(ey, sm, sday);
        if (anchor > end) anchor = dateToSerial(ey - 1, sm, sday);
        return Math.floor(end) - anchor;
      }
      default:
        return new CellError("#NUM!");
    }
  },
  DAYS(args, ev) {
    const end = num1(ev, args[0]);
    if (end instanceof CellError) return end;
    const start = num1(ev, args[1]);
    if (start instanceof CellError) return start;
    return Math.floor(end) - Math.floor(start);
  },
  WEEKNUM(args, ev) {
    const n = num1(ev, args[0]);
    if (n instanceof CellError) return n;
    const d = serialToDate(n);
    const year = d.getUTCFullYear();
    const jan1 = dateToSerial(year, 1, 1);
    const jan1Dow = serialToDate(jan1).getUTCDay(); // 0=Sun
    const days = Math.floor(n) - jan1 + jan1Dow;
    return Math.floor(days / 7) + 1;
  },
  NETWORKDAYS(args, ev) {
    const start = num1(ev, args[0]);
    if (start instanceof CellError) return start;
    const end = num1(ev, args[1]);
    if (end instanceof CellError) return end;
    const holidays = args[2] ? numbers(flatten(ev, [args[2]])) : [];
    if (holidays instanceof CellError) return holidays;
    const holidaySet = new Set(holidays.map((h) => Math.floor(h)));
    let s = Math.floor(start);
    let e = Math.floor(end);
    const sign = s <= e ? 1 : -1;
    if (sign < 0) { const t = s; s = e; e = t; }
    let count = 0;
    for (let day = s; day <= e; day++) {
      const dow = serialToDate(day).getUTCDay();
      if (dow !== 0 && dow !== 6 && !holidaySet.has(day)) count++;
    }
    return count * sign;
  },
  WORKDAY(args, ev) {
    const start = num1(ev, args[0]);
    if (start instanceof CellError) return start;
    const days = num1(ev, args[1]);
    if (days instanceof CellError) return days;
    const holidays = args[2] ? numbers(flatten(ev, [args[2]])) : [];
    if (holidays instanceof CellError) return holidays;
    const holidaySet = new Set(holidays.map((h) => Math.floor(h)));
    let day = Math.floor(start);
    let remaining = Math.floor(days);
    const step = remaining >= 0 ? 1 : -1;
    remaining = Math.abs(remaining);
    while (remaining > 0) {
      day += step;
      const dow = serialToDate(day).getUTCDay();
      if (dow !== 0 && dow !== 6 && !holidaySet.has(day)) remaining--;
    }
    return day;
  },
  DATEVALUE(args, ev) {
    const s = toText(scalar(ev, args[0])).trim();
    const serial = parseDate(s);
    return serial === null ? new CellError("#VALUE!") : serial;
  },
  YEARFRAC(args, ev) {
    const start = num1(ev, args[0]);
    if (start instanceof CellError) return start;
    const end = num1(ev, args[1]);
    if (end instanceof CellError) return end;
    // basis 0: US 30/360
    const sd = serialToDate(start);
    const ed = serialToDate(end);
    let sday = sd.getUTCDate();
    let eday = ed.getUTCDate();
    const sm = sd.getUTCMonth() + 1, sy = sd.getUTCFullYear();
    const em = ed.getUTCMonth() + 1, ey = ed.getUTCFullYear();
    if (sday === 31) sday = 30;
    if (eday === 31 && sday === 30) eday = 30;
    const days = (ey - sy) * 360 + (em - sm) * 30 + (eday - sday);
    return days / 360;
  },

  // ===== Added: Information =====
  ISBLANK(args, ev) {
    return scalar(ev, args[0]) === null;
  },
  ISNUMBER(args, ev) {
    return typeof scalar(ev, args[0]) === "number";
  },
  ISTEXT(args, ev) {
    return typeof scalar(ev, args[0]) === "string";
  },
  ISNONTEXT(args, ev) {
    return typeof scalar(ev, args[0]) !== "string";
  },
  ISLOGICAL(args, ev) {
    return typeof scalar(ev, args[0]) === "boolean";
  },
  ISERROR(args, ev) {
    return scalar(ev, args[0]) instanceof CellError;
  },
  ISERR(args, ev) {
    const v = scalar(ev, args[0]);
    return v instanceof CellError && v.kind !== "#N/A";
  },
  ISNA(args, ev) {
    const v = scalar(ev, args[0]);
    return v instanceof CellError && v.kind === "#N/A";
  },
  ISEVEN(args, ev) {
    const n = num1(ev, args[0]);
    if (n instanceof CellError) return n;
    return Math.floor(Math.abs(n)) % 2 === 0;
  },
  ISODD(args, ev) {
    const n = num1(ev, args[0]);
    if (n instanceof CellError) return n;
    return Math.floor(Math.abs(n)) % 2 === 1;
  },
  N(args, ev) {
    const v = scalar(ev, args[0]);
    if (v instanceof CellError) return v;
    if (typeof v === "number") return v;
    if (typeof v === "boolean") return v ? 1 : 0;
    return 0;
  },
  TYPE(args, ev) {
    const v = ev.evalNode(args[0]);
    if (isRange(v)) return 64;
    if (v instanceof CellError) return 16;
    if (typeof v === "number") return 1;
    if (typeof v === "string") return 2;
    if (typeof v === "boolean") return 4;
    return 1; // blank treated as number
  },
  "ERROR.TYPE"(args, ev) {
    const v = scalar(ev, args[0]);
    if (!(v instanceof CellError)) return new CellError("#N/A");
    const map: Record<string, number> = {
      "#NULL!": 1,
      "#DIV/0!": 2,
      "#VALUE!": 3,
      "#REF!": 4,
      "#NAME?": 5,
      "#NUM!": 6,
      "#N/A": 7,
    };
    return map[v.kind] ?? new CellError("#N/A");
  },

  // ===== Added: Financial =====
  PMT(args, ev) {
    const rate = num1(ev, args[0]);
    if (rate instanceof CellError) return rate;
    const nper = num1(ev, args[1]);
    if (nper instanceof CellError) return nper;
    const pv = num1(ev, args[2]);
    if (pv instanceof CellError) return pv;
    const fv = args[3] ? num1(ev, args[3]) : 0;
    if (fv instanceof CellError) return fv;
    const type = args[4] ? num1(ev, args[4]) : 0;
    if (type instanceof CellError) return type;
    return pmtCalc(rate, nper, pv, fv, type);
  },
  FV(args, ev) {
    const rate = num1(ev, args[0]);
    if (rate instanceof CellError) return rate;
    const nper = num1(ev, args[1]);
    if (nper instanceof CellError) return nper;
    const pmt = num1(ev, args[2]);
    if (pmt instanceof CellError) return pmt;
    const pv = args[3] ? num1(ev, args[3]) : 0;
    if (pv instanceof CellError) return pv;
    const type = args[4] ? num1(ev, args[4]) : 0;
    if (type instanceof CellError) return type;
    if (rate === 0) return -(pv + pmt * nper);
    const pow = Math.pow(1 + rate, nper);
    return -(pv * pow + pmt * (1 + rate * type) * (pow - 1) / rate);
  },
  PV(args, ev) {
    const rate = num1(ev, args[0]);
    if (rate instanceof CellError) return rate;
    const nper = num1(ev, args[1]);
    if (nper instanceof CellError) return nper;
    const pmt = num1(ev, args[2]);
    if (pmt instanceof CellError) return pmt;
    const fv = args[3] ? num1(ev, args[3]) : 0;
    if (fv instanceof CellError) return fv;
    const type = args[4] ? num1(ev, args[4]) : 0;
    if (type instanceof CellError) return type;
    if (rate === 0) return -(fv + pmt * nper);
    const pow = Math.pow(1 + rate, nper);
    return -(fv + pmt * (1 + rate * type) * (pow - 1) / rate) / pow;
  },
  NPER(args, ev) {
    const rate = num1(ev, args[0]);
    if (rate instanceof CellError) return rate;
    const pmt = num1(ev, args[1]);
    if (pmt instanceof CellError) return pmt;
    const pv = num1(ev, args[2]);
    if (pv instanceof CellError) return pv;
    const fv = args[3] ? num1(ev, args[3]) : 0;
    if (fv instanceof CellError) return fv;
    const type = args[4] ? num1(ev, args[4]) : 0;
    if (type instanceof CellError) return type;
    if (rate === 0) {
      if (pmt === 0) return new CellError("#NUM!");
      return -(pv + fv) / pmt;
    }
    const a = pmt * (1 + rate * type);
    const num = (a - fv * rate);
    const den = (pv * rate + a);
    if (num / den <= 0) return new CellError("#NUM!");
    return Math.log(num / den) / Math.log(1 + rate);
  },
  NPV(args, ev) {
    const rate = num1(ev, args[0]);
    if (rate instanceof CellError) return rate;
    const vals = numbers(flatten(ev, args.slice(1)));
    if (vals instanceof CellError) return vals;
    let total = 0;
    for (let i = 0; i < vals.length; i++) {
      total += vals[i] / Math.pow(1 + rate, i + 1);
    }
    return total;
  },
  IPMT(args, ev) {
    const rate = num1(ev, args[0]);
    if (rate instanceof CellError) return rate;
    const per = num1(ev, args[1]);
    if (per instanceof CellError) return per;
    const nper = num1(ev, args[2]);
    if (nper instanceof CellError) return nper;
    const pv = num1(ev, args[3]);
    if (pv instanceof CellError) return pv;
    const fv = args[4] ? num1(ev, args[4]) : 0;
    if (fv instanceof CellError) return fv;
    const type = args[5] ? num1(ev, args[5]) : 0;
    if (type instanceof CellError) return type;
    if (per < 1 || per > nper) return new CellError("#NUM!");
    const r = ipmtCalc(rate, per, nper, pv, fv, type);
    return r;
  },
  PPMT(args, ev) {
    const rate = num1(ev, args[0]);
    if (rate instanceof CellError) return rate;
    const per = num1(ev, args[1]);
    if (per instanceof CellError) return per;
    const nper = num1(ev, args[2]);
    if (nper instanceof CellError) return nper;
    const pv = num1(ev, args[3]);
    if (pv instanceof CellError) return pv;
    const fv = args[4] ? num1(ev, args[4]) : 0;
    if (fv instanceof CellError) return fv;
    const type = args[5] ? num1(ev, args[5]) : 0;
    if (type instanceof CellError) return type;
    if (per < 1 || per > nper) return new CellError("#NUM!");
    const pmt = pmtCalc(rate, nper, pv, fv, type);
    const ipmt = ipmtCalc(rate, per, nper, pv, fv, type);
    return pmt - ipmt;
  },
};

// --- local helpers used by the table above ---

function unaryMath(args: Node[], ev: Evaluator, fn: (n: number) => number | CellError): EvalResult {
  const n = toNumber(scalar(ev, args[0]));
  if (n instanceof CellError) return n;
  return fn(n);
}

function asRange(ev: Evaluator, node: Node): RangeValue | null {
  const v = ev.evalNode(node);
  if (isRange(v)) return v;
  // a single scalar acts as a 1x1 range
  return { range: true, r0: 0, c0: 0, r1: 0, c1: 0, values: [[v as CellValue]] };
}

function datePart(args: Node[], ev: Evaluator, fn: (d: Date) => number): EvalResult {
  const n = toNumber(scalar(ev, args[0]));
  if (n instanceof CellError) return n;
  return fn(serialToDate(n));
}

function looseEqual(a: CellValue, b: CellValue): boolean {
  if (typeof a === "number" && typeof b === "number") return a === b;
  return toText(a).toLowerCase() === toText(b).toLowerCase();
}

function compareLE(a: CellValue, b: CellValue): boolean {
  if (typeof a === "number" && typeof b === "number") return a <= b;
  return toText(a).toLowerCase() <= toText(b).toLowerCase();
}

function addThousands(s: string): string {
  const [int, dec] = s.split(".");
  const sign = int.startsWith("-") ? "-" : "";
  const digits = int.replace("-", "").replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return sign + digits + (dec ? "." + dec : "");
}

// --- helpers for the added functions ---

function gcd2(a: number, b: number): number {
  a = Math.abs(a);
  b = Math.abs(b);
  while (b) {
    const t = b;
    b = a % b;
    a = t;
  }
  return a;
}

function factorial(n: number): number {
  let r = 1;
  for (let i = 2; i <= n; i++) r *= i;
  return r;
}

function modeImpl(args: Node[], ev: Evaluator): EvalResult {
  const nums = numbers(flatten(ev, args));
  if (nums instanceof CellError) return nums;
  const counts = new Map<number, number>();
  const firstIndex = new Map<number, number>();
  nums.forEach((n, i) => {
    counts.set(n, (counts.get(n) ?? 0) + 1);
    if (!firstIndex.has(n)) firstIndex.set(n, i);
  });
  let best: number | null = null;
  let bestCount = 1;
  for (const [val, c] of counts) {
    if (c > bestCount || (c === bestCount && best !== null && firstIndex.get(val)! < firstIndex.get(best)!)) {
      best = val;
      bestCount = c;
    }
  }
  if (best === null || bestCount < 2) return new CellError("#N/A");
  return best;
}

function varianceImpl(args: Node[], ev: Evaluator, sample: boolean, sqrt: boolean): EvalResult {
  const nums = numbers(flatten(ev, args));
  if (nums instanceof CellError) return nums;
  const n = nums.length;
  if (sample && n < 2) return new CellError("#DIV/0!");
  if (!sample && n < 1) return new CellError("#DIV/0!");
  const mean = nums.reduce((a, b) => a + b, 0) / n;
  const ss = nums.reduce((a, b) => a + (b - mean) * (b - mean), 0);
  const variance = ss / (sample ? n - 1 : n);
  return sqrt ? Math.sqrt(variance) : variance;
}

function percentile(nums: number[], k: number): EvalResult {
  if (nums.length === 0 || k < 0 || k > 1) return new CellError("#NUM!");
  const s = [...nums].sort((a, b) => a - b);
  if (s.length === 1) return s[0];
  const pos = k * (s.length - 1);
  const lo = Math.floor(pos);
  const hi = Math.ceil(pos);
  if (lo === hi) return s[lo];
  return s[lo] + (pos - lo) * (s[hi] - s[lo]);
}

function ifsAggregate(
  args: Node[],
  ev: Evaluator,
  pairStart: 0 | 1,
  mode: "sum" | "count" | "avg" | "max" | "min"
): EvalResult {
  // For COUNTIFS, pairs start at index 0; otherwise index 0 is the value range.
  let valueRange: RangeValue | null = null;
  let pairsFrom = pairStart;
  if (pairStart === 1) {
    valueRange = asRange(ev, args[0]);
    if (!valueRange) return new CellError("#VALUE!");
    pairsFrom = 1 as 0 | 1;
  }
  const matchers: { range: RangeValue; match: (v: CellValue) => boolean }[] = [];
  for (let i = pairsFrom; i + 1 < args.length; i += 2) {
    const r = asRange(ev, args[i]);
    if (!r) return new CellError("#VALUE!");
    const crit = scalar(ev, args[i + 1]);
    matchers.push({ range: r, match: makeMatcher(crit) });
  }
  if (matchers.length === 0) return new CellError("#VALUE!");
  const len = matchers[0].range.values.flat().length;
  const flats = matchers.map((m) => m.range.values.flat());
  if (flats.some((f) => f.length !== len)) return new CellError("#VALUE!");
  const valFlat = valueRange ? valueRange.values.flat() : null;
  if (valFlat && valFlat.length !== len) return new CellError("#VALUE!");

  let total = 0;
  let count = 0;
  let best: number | null = null;
  for (let i = 0; i < len; i++) {
    let ok = true;
    for (let m = 0; m < matchers.length; m++) {
      if (!matchers[m].match(flats[m][i])) {
        ok = false;
        break;
      }
    }
    if (!ok) continue;
    if (mode === "count") {
      count++;
      continue;
    }
    const v = valFlat![i];
    if (typeof v !== "number") continue;
    count++;
    total += v;
    if (best === null) best = v;
    else if (mode === "max") best = Math.max(best, v);
    else if (mode === "min") best = Math.min(best, v);
  }
  switch (mode) {
    case "count":
      return count;
    case "sum":
      return total;
    case "avg":
      return count === 0 ? new CellError("#DIV/0!") : total / count;
    case "max":
      return best === null ? 0 : best;
    case "min":
      return best === null ? 0 : best;
  }
}

function timeFrac(serial: number): number {
  const frac = serial - Math.floor(serial);
  return frac < 0 ? frac + 1 : frac;
}

function edate(startSerial: number, months: number): number {
  const d = serialToDate(startSerial);
  const y = d.getUTCFullYear();
  const m = d.getUTCMonth() + 1 + months;
  const day = d.getUTCDate();
  // clamp day to last day of target month
  const lastDay = new Date(Date.UTC(y, m, 0)).getUTCDate();
  return dateToSerial(y, m, Math.min(day, lastDay));
}

function parseDate(s: string): number | null {
  // ISO yyyy-mm-dd or yyyy/mm/dd
  let m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(s);
  if (m) return dateToSerial(+m[1], +m[2], +m[3]);
  // mm/dd/yyyy or mm-dd-yyyy
  m = /^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/.exec(s);
  if (m) return dateToSerial(+m[3], +m[1], +m[2]);
  // dd-Mon-yyyy or Mon dd, yyyy
  const months: Record<string, number> = {
    jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
    jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12,
  };
  m = /^(\d{1,2})[-\s]([A-Za-z]{3,})[-\s](\d{4})$/.exec(s);
  if (m) {
    const mon = months[m[2].slice(0, 3).toLowerCase()];
    if (mon) return dateToSerial(+m[3], mon, +m[1]);
  }
  m = /^([A-Za-z]{3,})\s+(\d{1,2}),?\s+(\d{4})$/.exec(s);
  if (m) {
    const mon = months[m[1].slice(0, 3).toLowerCase()];
    if (mon) return dateToSerial(+m[3], mon, +m[2]);
  }
  return null;
}

function pmtCalc(rate: number, nper: number, pv: number, fv: number, type: number): number {
  if (rate === 0) return -(pv + fv) / nper;
  const pow = Math.pow(1 + rate, nper);
  return -(rate * (fv + pv * pow)) / ((1 + rate * type) * (pow - 1));
}

function ipmtCalc(rate: number, per: number, nper: number, pv: number, fv: number, type: number): number {
  const pmt = pmtCalc(rate, nper, pv, fv, type);
  // balance at start of period `per`
  let bal = pv;
  let interest = 0;
  for (let p = 1; p <= per; p++) {
    if (type === 1 && p === 1) {
      interest = 0;
    } else {
      interest = bal * rate;
    }
    const principal = pmt - interest;
    bal += principal;
  }
  return interest;
}

// Sorted list of all built-in function names (for editor autocomplete).
export const FUNCTION_NAMES = Object.keys(FUNCTIONS).sort();
