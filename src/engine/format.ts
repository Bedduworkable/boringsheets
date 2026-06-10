// Turns a computed CellValue into the string the grid displays, honoring an
// Excel-style number-format code. Implements a pragmatic, common subset:
// decimals, thousands separators, percent, currency, and date/time codes.

import { CellValue, CellError } from "../model/types.js";
import { numToText } from "./evaluator.js";

const MS_PER_DAY = 86400000;
const EPOCH = Date.UTC(1899, 11, 30);
const MONTHS = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];
const DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

function serialToDate(serial: number): Date {
  return new Date(EPOCH + Math.round(serial * MS_PER_DAY));
}

function isDateFormat(fmt: string): boolean {
  // strip quoted literals, then look for date/time tokens
  const f = fmt.replace(/"[^"]*"/g, "").toLowerCase();
  return /[ymdhs]/.test(f) && !/[#0]/.test(f.replace(/[ymdhs]/g, ""));
}

export function formatValue(value: CellValue, numFmt?: string): string {
  if (value === null) return "";
  if (value instanceof CellError) return value.kind;
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "string") return value;

  const n = value;
  if (!numFmt || numFmt === "General" || numFmt === "") return numToText(n);

  if (isDateFormat(numFmt)) return formatDate(n, numFmt);
  return formatNumber(n, numFmt);
}

function formatNumber(n: number, fmt: string): string {
  const isPercent = fmt.includes("%");
  const val = isPercent ? n * 100 : n;

  // leading literal: a quoted "..." segment or a bare currency symbol
  let prefix = "";
  const trimmed = fmt.trim();
  const lit = /^"([^"]*)"/.exec(trimmed);
  if (lit) prefix = lit[1];
  else {
    const cur = /^[$£€]/.exec(trimmed);
    if (cur) prefix = cur[0];
  }

  // decimal places = number of 0/# after the dot
  const dotIdx = fmt.indexOf(".");
  let decimals = 0;
  if (dotIdx !== -1) {
    const after = fmt.slice(dotIdx + 1);
    decimals = (after.match(/[0#]/g) || []).length;
  }

  const useThousands = /[#0],[#0]/.test(fmt) || fmt.includes("#,##0");

  const sign = val < 0 ? "-" : "";
  let s = Math.abs(val).toFixed(decimals);
  if (useThousands) s = addThousands(s);

  return sign + prefix + s + (isPercent ? "%" : "");
}

function addThousands(s: string): string {
  const [int, dec] = s.split(".");
  const grouped = int.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return dec ? `${grouped}.${dec}` : grouped;
}

function formatDate(serial: number, fmt: string): string {
  const d = serialToDate(serial);
  const frac = serial - Math.floor(serial);
  let secs = Math.round(frac * 86400);
  const hh = Math.floor(secs / 3600);
  secs -= hh * 3600;
  const mm = Math.floor(secs / 60);
  const ss = secs - mm * 60;

  const year = d.getUTCFullYear();
  const month = d.getUTCMonth() + 1;
  const day = d.getUTCDate();
  const dow = d.getUTCDay();
  const hasAmPm = /am\/pm|a\/p/i.test(fmt);
  const h12 = hh % 12 === 0 ? 12 : hh % 12;

  const pad = (x: number, w = 2) => String(x).padStart(w, "0");

  // Tokenize so each token is replaced exactly once. 'm'/'mm' is a minute when
  // adjacent to an hour/second token, otherwise a month.
  const tokenRe = /yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|am\/pm|a\/p|"[^"]*"|./gi;
  const parts = fmt.match(tokenRe) || [];
  const hasTime = parts.some((p) => /^h/i.test(p) || /^s/i.test(p));

  let out = "";
  parts.forEach((p, i) => {
    const lo = p.toLowerCase();
    const prev = (parts[i - 1] || "").toLowerCase();
    const next = (parts[i + 1] || "").toLowerCase();
    const minuteCtx = hasTime && (/^h/.test(prev) || /^s/.test(next) || /^h/.test(next));
    switch (lo) {
      case "yyyy": out += year; break;
      case "yy": out += pad(year % 100); break;
      case "mmmm": out += MONTHS[month - 1]; break;
      case "mmm": out += MONTHS[month - 1].slice(0, 3); break;
      case "mm": out += minuteCtx ? pad(mm) : pad(month); break;
      case "m": out += minuteCtx ? String(mm) : String(month); break;
      case "dddd": out += DAYS[dow]; break;
      case "ddd": out += DAYS[dow].slice(0, 3); break;
      case "dd": out += pad(day); break;
      case "d": out += String(day); break;
      case "hh": out += pad(hasAmPm ? h12 : hh); break;
      case "h": out += String(hasAmPm ? h12 : hh); break;
      case "ss": out += pad(ss); break;
      case "s": out += String(ss); break;
      case "am/pm": out += hh < 12 ? "AM" : "PM"; break;
      case "a/p": out += hh < 12 ? "A" : "P"; break;
      default:
        out += p.startsWith('"') ? p.slice(1, -1) : p;
    }
  });
  return out;
}
