// Pure conditional-formatting engine. Resolves a per-cell visual (text/fill
// format overrides, data bars, or color-scale fills) from a list of rules.
// Range statistics are computed lazily and cached per rule id.

import { CellValue, CellError, CellFormat } from "../model/types.js";

export type CondType =
  | "greaterThan" | "lessThan" | "between" | "equalTo" | "notEqualTo"
  | "textContains" | "duplicate" | "unique" | "top" | "bottom"
  | "colorScale" | "dataBar";

export interface ConditionalRule {
  id: string;
  range: { r0: number; c0: number; r1: number; c1: number };
  type: CondType;
  value1?: number | string;   // threshold / compare value / search text
  value2?: number;            // upper bound for "between"
  n?: number;                 // count for top/bottom
  format?: CellFormat;        // applied for boolean-style rules
  color?: string;             // dataBar fill color, or colorScale "max" color
  minColor?: string;          // colorScale low end
  midColor?: string;          // colorScale mid (optional 3-color scale)
  maxColor?: string;          // colorScale high end
}

export interface CondVisual {
  format?: CellFormat;                         // text/fill overrides to merge over the cell's own format
  dataBar?: { color: string; fraction: number }; // fraction in [0,1]
  fillColor?: string;                          // solid background from a color scale
}

// Cached per-rule range statistics.
interface RangeStats {
  numbers: number[];          // participating numeric values, ascending
  min: number;                // NaN when no numbers
  max: number;                // NaN when no numbers
  median: number;             // NaN when no numbers
  counts: Map<string, number>; // value key -> occurrence count (all participating cells)
}

// A cell participates in stats unless it is null, "" or a CellError.
function participates(v: CellValue): boolean {
  if (v === null) return false;
  if (v instanceof CellError) return false;
  if (typeof v === "string" && v === "") return false;
  return true;
}

function asNumber(v: CellValue): number | null {
  return typeof v === "number" ? v : null;
}

// A stable key for duplicate/unique counting. Booleans and numbers and strings
// are distinguished by type so 1 !== "1" !== true.
function valueKey(v: CellValue): string {
  if (typeof v === "number") return "n:" + v;
  if (typeof v === "boolean") return "b:" + (v ? "1" : "0");
  return "s:" + String(v);
}

function clamp01(x: number): number {
  if (x < 0) return 0;
  if (x > 1) return 1;
  return x;
}

function parseHex(hex: string): [number, number, number] {
  let h = hex.trim();
  if (h.startsWith("#")) h = h.slice(1);
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  return [r, g, b];
}

function toHex(c: [number, number, number]): string {
  const part = (x: number) => {
    const v = Math.max(0, Math.min(255, Math.round(x)));
    return v.toString(16).padStart(2, "0");
  };
  return "#" + part(c[0]) + part(c[1]) + part(c[2]);
}

function lerpColor(a: string, b: string, t: number): string {
  const ca = parseHex(a);
  const cb = parseHex(b);
  return toHex([
    ca[0] + (cb[0] - ca[0]) * t,
    ca[1] + (cb[1] - ca[1]) * t,
    ca[2] + (cb[2] - ca[2]) * t,
  ]);
}

export class ConditionalEngine {
  private rules: ConditionalRule[];
  private statsCache = new Map<string, RangeStats>();

  constructor(rules: ConditionalRule[]) {
    this.rules = rules;
  }

  private inRange(rule: ConditionalRule, row: number, col: number): boolean {
    const r = rule.range;
    return row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1;
  }

  private statsFor(rule: ConditionalRule, getValue: (r: number, c: number) => CellValue): RangeStats {
    const cached = this.statsCache.get(rule.id);
    if (cached) return cached;

    const numbers: number[] = [];
    const counts = new Map<string, number>();
    const r = rule.range;
    for (let row = r.r0; row <= r.r1; row++) {
      for (let col = r.c0; col <= r.c1; col++) {
        const v = getValue(row, col);
        if (!participates(v)) continue;
        counts.set(valueKey(v), (counts.get(valueKey(v)) ?? 0) + 1);
        const num = asNumber(v);
        if (num !== null) numbers.push(num);
      }
    }
    numbers.sort((a, b) => a - b);

    let min = NaN, max = NaN, median = NaN;
    if (numbers.length > 0) {
      min = numbers[0];
      max = numbers[numbers.length - 1];
      const mid = Math.floor(numbers.length / 2);
      median = numbers.length % 2 === 1
        ? numbers[mid]
        : (numbers[mid - 1] + numbers[mid]) / 2;
    }

    const stats: RangeStats = { numbers, min, max, median, counts };
    this.statsCache.set(rule.id, stats);
    return stats;
  }

  private boolMatch(
    rule: ConditionalRule,
    value: CellValue,
    getValue: (r: number, c: number) => CellValue,
  ): boolean {
    const num = asNumber(value);
    switch (rule.type) {
      case "greaterThan":
        return num !== null && typeof rule.value1 === "number" && num > rule.value1;
      case "lessThan":
        return num !== null && typeof rule.value1 === "number" && num < rule.value1;
      case "between":
        return num !== null && typeof rule.value1 === "number" && typeof rule.value2 === "number"
          && num >= rule.value1 && num <= rule.value2;
      case "equalTo":
      case "notEqualTo": {
        let isEqual: boolean;
        if (num !== null && typeof rule.value1 === "number") {
          isEqual = num === rule.value1;
        } else {
          isEqual = String(value ?? "").toLowerCase() === String(rule.value1 ?? "").toLowerCase();
        }
        return rule.type === "equalTo" ? isEqual : !isEqual;
      }
      case "textContains": {
        const hay = String(value ?? "").toLowerCase();
        const needle = String(rule.value1 ?? "").toLowerCase();
        return needle.length > 0 && hay.includes(needle);
      }
      case "duplicate": {
        if (!participates(value)) return false;
        const stats = this.statsFor(rule, getValue);
        return (stats.counts.get(valueKey(value)) ?? 0) > 1;
      }
      case "unique": {
        if (!participates(value)) return false;
        const stats = this.statsFor(rule, getValue);
        return (stats.counts.get(valueKey(value)) ?? 0) === 1;
      }
      case "top":
      case "bottom": {
        if (num === null) return false;
        const stats = this.statsFor(rule, getValue);
        const n = rule.n ?? 1;
        if (n <= 0 || stats.numbers.length === 0) return false;
        if (rule.type === "top") {
          // The n highest values: threshold is the n-th from the top.
          const idx = Math.max(0, stats.numbers.length - n);
          return num >= stats.numbers[idx];
        } else {
          const idx = Math.min(stats.numbers.length - 1, n - 1);
          return num <= stats.numbers[idx];
        }
      }
      default:
        return false;
    }
  }

  resolve(
    row: number,
    col: number,
    value: CellValue,
    getValue: (r: number, c: number) => CellValue,
  ): CondVisual | null {
    let format: CellFormat | undefined;
    let dataBar: { color: string; fraction: number } | undefined;
    let fillColor: string | undefined;
    let matched = false;

    for (const rule of this.rules) {
      if (!this.inRange(rule, row, col)) continue;

      switch (rule.type) {
        case "colorScale": {
          const num = asNumber(value);
          if (num === null) break;
          const stats = this.statsFor(rule, getValue);
          if (stats.numbers.length === 0 || Number.isNaN(stats.min) || Number.isNaN(stats.max)) break;
          const lo = rule.minColor ?? "#ffffff";
          const hi = rule.maxColor ?? rule.color ?? "#000000";
          if (stats.max === stats.min) {
            fillColor = lo;
          } else if (rule.midColor) {
            const med = stats.median;
            if (num <= med) {
              const span = med - stats.min;
              const t = span === 0 ? 0 : (num - stats.min) / span;
              fillColor = lerpColor(lo, rule.midColor, clamp01(t));
            } else {
              const span = stats.max - med;
              const t = span === 0 ? 1 : (num - med) / span;
              fillColor = lerpColor(rule.midColor, hi, clamp01(t));
            }
          } else {
            const t = (num - stats.min) / (stats.max - stats.min);
            fillColor = lerpColor(lo, hi, clamp01(t));
          }
          dataBar = undefined; // a later scale replaces an earlier bar
          matched = true;
          break;
        }
        case "dataBar": {
          const num = asNumber(value);
          if (num === null) break;
          const stats = this.statsFor(rule, getValue);
          if (stats.numbers.length === 0 || Number.isNaN(stats.min) || Number.isNaN(stats.max)) break;
          const fraction = stats.max === stats.min ? 0 : clamp01((num - stats.min) / (stats.max - stats.min));
          dataBar = { color: rule.color ?? "#4285f4", fraction };
          fillColor = undefined; // a later bar replaces an earlier scale
          matched = true;
          break;
        }
        default: {
          if (this.boolMatch(rule, value, getValue)) {
            format = { ...(format ?? {}), ...(rule.format ?? {}) };
            matched = true;
          }
          break;
        }
      }
    }

    if (!matched) return null;
    const out: CondVisual = {};
    if (format !== undefined) out.format = format;
    if (dataBar !== undefined) out.dataBar = dataBar;
    if (fillColor !== undefined) out.fillColor = fillColor;
    return out;
  }
}
