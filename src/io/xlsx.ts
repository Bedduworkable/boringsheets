// Read and write real Excel .xlsx files (SpreadsheetML packaged in a zip).
// Reading uses the browser DOMParser; writing builds the XML by hand. This is a
// pragmatic subset: values, formulas, shared strings, and basic styles
// (number formats, bold/italic, text color, fill color, alignment).

import { zipSync, unzipSync, strToU8, strFromU8 } from "fflate";
import { Workbook } from "../model/workbook.js";
import { Sheet, DataValidation } from "../model/sheet.js";
import { CellFormat, CellValue, CellError, BorderSet } from "../model/types.js";
import { ConditionalRule, CondType } from "../engine/conditional.js";
import { parseA1, colToLetter, cellKey, formatA1, formatSheetName } from "../engine/references.js";
import { formatValue } from "../engine/format.js";

const PX_PER_PT = 96 / 72; // 1 point = 1.333px at 96dpi
const opMap: Record<string, string> = {
  between: "between",
  notBetween: "notBetween",
  gt: "greaterThan",
  lt: "lessThan",
  gte: "greaterThanOrEqual",
  lte: "lessThanOrEqual",
  eq: "equal",
  ne: "notEqual",
};
const opMapRev: Record<string, DataValidation["operator"]> = Object.fromEntries(
  Object.entries(opMap).map(([k, v]) => [v, k as DataValidation["operator"]])
);

// Built-in number-format ids Excel doesn't store inline.
const BUILTIN_NUMFMT: Record<number, string> = {
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  14: "mm/dd/yyyy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  22: "m/d/yyyy h:mm",
  37: "#,##0;(#,##0)",
  38: "#,##0;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)",
  40: "#,##0.00;[Red](#,##0.00)",
  44: '"$"#,##0.00',
  49: "@",
};

const xmlEscape = (s: string) =>
  s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

// ===================== READ =====================

export function readXlsx(data: ArrayBuffer): Workbook {
  const files = unzipSync(new Uint8Array(data));
  const parser = new DOMParser();
  const get = (path: string) => (files[path] ? strFromU8(files[path]) : null);

  const sharedStrings = parseSharedStrings(get("xl/sharedStrings.xml"), parser);
  const styles = parseStyles(get("xl/styles.xml"), parser);

  // workbook.xml lists the sheets and their names; the rels map sheet -> file.
  const wbXml = get("xl/workbook.xml");
  const relsXml = get("xl/_rels/workbook.xml.rels");
  const sheetList = parseWorkbookSheets(wbXml, relsXml, parser);

  const wb = new Workbook();
  wb.sheets = [];

  for (const { name, path } of sheetList) {
    const sheetXml = get(path);
    const sheet = new Sheet(name);
    if (sheetXml) {
      const comments = parseSheetComments(path, get, parser);
      parseSheet(sheetXml, parser, sheet, sharedStrings, styles, comments);
    }
    wb.sheets.push(sheet);
  }
  if (wb.sheets.length === 0) wb.sheets.push(new Sheet("Sheet1"));

  // workbook-level named ranges (resolve sheet names to the loaded sheets)
  if (wbXml) parseDefinedNames(parser.parseFromString(wbXml, "application/xml"), wb);

  wb.activeIndex = 0;
  return wb;
}

function resolvePath(base: string, rel: string): string {
  if (rel.startsWith("/")) return rel.slice(1);
  const parts = base.split("/");
  for (const seg of rel.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg !== "." && seg !== "") parts.push(seg);
  }
  return parts.join("/");
}

function parseSheetComments(sheetPath: string, get: (p: string) => string | null, parser: DOMParser): Map<string, string> {
  const map = new Map<string, string>();
  const slash = sheetPath.lastIndexOf("/");
  const dir = sheetPath.slice(0, slash);
  const fname = sheetPath.slice(slash + 1);
  const relsXml = get(`${dir}/_rels/${fname}.rels`);
  if (!relsXml) return map;
  const relsDoc = parser.parseFromString(relsXml, "application/xml");
  let target: string | null = null;
  for (const r of Array.from(relsDoc.getElementsByTagName("Relationship"))) {
    if ((r.getAttribute("Type") || "").endsWith("/comments")) {
      target = r.getAttribute("Target");
      break;
    }
  }
  if (!target) return map;
  const cxml = get(resolvePath(dir, target));
  if (!cxml) return map;
  const cdoc = parser.parseFromString(cxml, "application/xml");
  for (const cm of Array.from(cdoc.getElementsByTagName("comment"))) {
    const ref = cm.getAttribute("ref");
    const text = Array.from(cm.getElementsByTagName("t")).map((t) => t.textContent ?? "").join("");
    if (ref) map.set(ref, text);
  }
  return map;
}

function parseDefinedNames(wbDoc: Document, wb: Workbook) {
  for (const dn of Array.from(wbDoc.getElementsByTagName("definedName"))) {
    const name = dn.getAttribute("name") || "";
    if (!name || name.startsWith("_xlnm") || !/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) continue;
    const ref = (dn.textContent || "").trim();
    const bang = ref.lastIndexOf("!");
    if (bang < 0) continue;
    let sheetName = ref.slice(0, bang);
    if (sheetName.startsWith("'") && sheetName.endsWith("'")) sheetName = sheetName.slice(1, -1).replace(/''/g, "'");
    const sheet = wb.sheets.find((s) => s.name.toLowerCase() === sheetName.toLowerCase());
    if (!sheet) continue;
    const [a, b] = ref.slice(bang + 1).split(":");
    const p0 = parseA1(a);
    const p1 = parseA1(b || a);
    if (!p0 || !p1) continue;
    wb.names.push({
      name,
      sheetId: sheet.id,
      r0: Math.min(p0.row, p1.row),
      c0: Math.min(p0.col, p1.col),
      r1: Math.max(p0.row, p1.row),
      c1: Math.max(p0.col, p1.col),
    });
  }
}

function parseSharedStrings(xml: string | null, parser: DOMParser): string[] {
  if (!xml) return [];
  const doc = parser.parseFromString(xml, "application/xml");
  return Array.from(doc.getElementsByTagName("si")).map((si) => {
    // join all <t> descendants (rich text runs)
    return Array.from(si.getElementsByTagName("t"))
      .map((t) => t.textContent ?? "")
      .join("");
  });
}

interface StyleInfo {
  numFmts: Map<number, string>; // styleIndex -> format code
  formats: Map<number, CellFormat>; // styleIndex -> full format
  dxfs: CellFormat[]; // differential formats (for conditional formatting)
}

function parseFontFormat(font: Element, fmt: CellFormat) {
  if (font.getElementsByTagName("b").length) fmt.bold = true;
  if (font.getElementsByTagName("i").length) fmt.italic = true;
  if (font.getElementsByTagName("u").length) fmt.underline = true;
  if (font.getElementsByTagName("strike").length) fmt.strike = true;
  const color = font.getElementsByTagName("color")[0]?.getAttribute("rgb");
  if (color) fmt.color = argbToHex(color);
  const sz = font.getElementsByTagName("sz")[0]?.getAttribute("val");
  if (sz) fmt.fontSize = Math.round((parseFloat(sz) * 4) / 3); // points → px
  const name = font.getElementsByTagName("name")[0]?.getAttribute("val");
  if (name && name !== "Arial" && name !== "Calibri") fmt.fontFamily = name;
}

function parseBorder(el: Element): BorderSet | undefined {
  const b: BorderSet = {};
  let color: string | undefined;
  for (const tag of ["left", "right", "top", "bottom"] as const) {
    const side = el.getElementsByTagName(tag)[0];
    const style = side?.getAttribute("style");
    if (style && style !== "none") {
      b[tag] = true;
      const c = side.getElementsByTagName("color")[0]?.getAttribute("rgb");
      if (c && !color) color = argbToHex(c);
    }
  }
  if (!b.left && !b.right && !b.top && !b.bottom) return undefined;
  if (color) b.color = color;
  return b;
}

function parseStyles(xml: string | null, parser: DOMParser): StyleInfo {
  const numFmts = new Map<number, string>();
  const formats = new Map<number, CellFormat>();
  const dxfs: CellFormat[] = [];
  if (!xml) return { numFmts, formats, dxfs };
  const doc = parser.parseFromString(xml, "application/xml");

  const customNumFmt = new Map<number, string>();
  for (const nf of Array.from(doc.getElementsByTagName("numFmt"))) {
    customNumFmt.set(parseInt(nf.getAttribute("numFmtId") || "0", 10), nf.getAttribute("formatCode") || "");
  }

  const fonts = Array.from(doc.getElementsByTagName("fonts")[0]?.getElementsByTagName("font") || []);
  const fills = Array.from(doc.getElementsByTagName("fills")[0]?.getElementsByTagName("fill") || []);
  const borders = Array.from(doc.getElementsByTagName("borders")[0]?.getElementsByTagName("border") || []);

  // differential formats (conditional formatting targets these by dxfId)
  for (const dxf of Array.from(doc.getElementsByTagName("dxfs")[0]?.getElementsByTagName("dxf") || [])) {
    const f: CellFormat = {};
    const font = dxf.getElementsByTagName("font")[0];
    if (font) parseFontFormat(font, f);
    const fill = dxf.getElementsByTagName("fill")[0];
    const bg = fill?.getElementsByTagName("bgColor")[0]?.getAttribute("rgb");
    if (bg) f.bg = argbToHex(bg);
    dxfs.push(f);
  }

  const cellXfs = doc.getElementsByTagName("cellXfs")[0];
  if (!cellXfs) return { numFmts, formats, dxfs };
  const xfs = Array.from(cellXfs.getElementsByTagName("xf"));

  xfs.forEach((xf, i) => {
    const fmt: CellFormat = {};
    const numFmtId = parseInt(xf.getAttribute("numFmtId") || "0", 10);
    const code = customNumFmt.get(numFmtId) ?? BUILTIN_NUMFMT[numFmtId];
    if (code && code !== "General") {
      fmt.numFmt = code;
      numFmts.set(i, code);
    }
    const font = fonts[parseInt(xf.getAttribute("fontId") || "0", 10)];
    if (font && xf.getAttribute("fontId") !== "0") parseFontFormat(font, fmt);

    const fillId = parseInt(xf.getAttribute("fillId") || "0", 10);
    const fgColor = fills[fillId]?.getElementsByTagName("fgColor")[0]?.getAttribute("rgb");
    if (fgColor && fillId > 1) fmt.bg = argbToHex(fgColor);

    const borderId = parseInt(xf.getAttribute("borderId") || "0", 10);
    if (borderId > 0 && borders[borderId]) {
      const b = parseBorder(borders[borderId]);
      if (b) fmt.border = b;
    }

    const alignEl = xf.getElementsByTagName("alignment")[0];
    const h = alignEl?.getAttribute("horizontal");
    if (h === "left" || h === "center" || h === "right") fmt.align = h;
    if (alignEl?.getAttribute("wrapText") === "1") fmt.wrap = true;

    if (Object.keys(fmt).length) formats.set(i, fmt);
  });

  return { numFmts, formats, dxfs };
}

function parseWorkbookSheets(
  wbXml: string | null,
  relsXml: string | null,
  parser: DOMParser
): { name: string; path: string }[] {
  const result: { name: string; path: string }[] = [];
  if (!wbXml) return [{ name: "Sheet1", path: "xl/worksheets/sheet1.xml" }];
  const wbDoc = parser.parseFromString(wbXml, "application/xml");
  const relsDoc = relsXml ? parser.parseFromString(relsXml, "application/xml") : null;

  const relMap = new Map<string, string>();
  if (relsDoc) {
    for (const r of Array.from(relsDoc.getElementsByTagName("Relationship"))) {
      relMap.set(r.getAttribute("Id") || "", r.getAttribute("Target") || "");
    }
  }

  const sheets = Array.from(wbDoc.getElementsByTagName("sheet"));
  sheets.forEach((s, i) => {
    const name = s.getAttribute("name") || `Sheet${i + 1}`;
    const rid = s.getAttribute("r:id") || s.getAttributeNS("*", "id") || "";
    let target = relMap.get(rid) || `worksheets/sheet${i + 1}.xml`;
    if (!target.startsWith("xl/")) target = "xl/" + target.replace(/^\//, "");
    result.push({ name, path: target });
  });
  return result;
}

function parseSheet(
  xml: string,
  parser: DOMParser,
  sheet: Sheet,
  shared: string[],
  styles: StyleInfo,
  comments: Map<string, string>
) {
  const doc = parser.parseFromString(xml, "application/xml");

  // freeze panes
  const pane = doc.getElementsByTagName("pane")[0];
  if (pane && (pane.getAttribute("state") === "frozen" || !pane.getAttribute("state"))) {
    sheet.frozenCols = parseInt(pane.getAttribute("xSplit") || "0", 10) || 0;
    sheet.frozenRows = parseInt(pane.getAttribute("ySplit") || "0", 10) || 0;
  }

  // columns: widths + hidden
  for (const col of Array.from(doc.getElementsByTagName("col"))) {
    const min = parseInt(col.getAttribute("min") || "1", 10);
    const max = parseInt(col.getAttribute("max") || "1", 10);
    const width = parseFloat(col.getAttribute("width") || "0");
    const hidden = col.getAttribute("hidden") === "1";
    for (let c = min; c <= max; c++) {
      if (width > 0 && col.getAttribute("customWidth") === "1") sheet.colWidths.set(c - 1, Math.round(width * 7 + 5));
      if (hidden) sheet.hiddenCols.add(c - 1);
    }
  }

  // rows: heights + hidden
  for (const row of Array.from(doc.getElementsByTagName("row"))) {
    const r = parseInt(row.getAttribute("r") || "0", 10) - 1;
    if (r < 0) continue;
    if (row.getAttribute("hidden") === "1") sheet.hiddenRows.add(r);
    const ht = row.getAttribute("ht");
    if (ht && row.getAttribute("customHeight") === "1") sheet.rowHeights.set(r, Math.round(parseFloat(ht) * PX_PER_PT));
  }

  // merges
  for (const m of Array.from(doc.getElementsByTagName("mergeCell"))) {
    const ref = m.getAttribute("ref");
    if (!ref) continue;
    const [a, b] = ref.split(":");
    const p0 = parseA1(a);
    const p1 = parseA1(b || a);
    if (p0 && p1) sheet.merges.push({ r0: p0.row, c0: p0.col, r1: p1.row, c1: p1.col });
  }

  // data validations
  for (const dv of Array.from(doc.getElementsByTagName("dataValidation"))) {
    parseDataValidation(dv, sheet);
  }

  // conditional formatting
  for (const cf of Array.from(doc.getElementsByTagName("conditionalFormatting"))) {
    const sqref = cf.getAttribute("sqref") || "";
    const range = parseSqref(sqref);
    if (!range) continue;
    for (const rule of Array.from(cf.getElementsByTagName("cfRule"))) {
      const parsed = parseCfRule(rule, range, styles.dxfs);
      if (parsed) sheet.conditionalRules.push(parsed);
    }
  }

  for (const c of Array.from(doc.getElementsByTagName("c"))) {
    const ref = c.getAttribute("r");
    if (!ref) continue;
    const pos = parseA1(ref);
    if (!pos) continue;
    const t = c.getAttribute("t"); // s | b | str | inlineStr | e | n(default)
    const styleIdx = c.getAttribute("s") ? parseInt(c.getAttribute("s")!, 10) : -1;
    const fEl = c.getElementsByTagName("f")[0];
    const vEl = c.getElementsByTagName("v")[0];
    const vText = vEl?.textContent ?? "";
    const formulaText = (fEl?.textContent ?? "").trim();

    let raw = "";
    let value: CellValue = null;

    // A real formula has inline text. Excel "shared formula" DEPENDENTS have an
    // empty <f t="shared" si="N"/> — keep their cached value instead of a broken
    // "=" formula.
    if (fEl && formulaText !== "") {
      raw = "=" + fEl.textContent;
      // (value recomputed on rebuild; keep the cached <v> as a best-effort)
      if (t === "str") value = vText;
      else if (t === "b") value = vText === "1";
      else if (t === "e") value = new CellError((vText as CellError["kind"]) || "#VALUE!");
      else if (vText !== "") value = Number(vText);
    } else if (t === "s") {
      const idx = parseInt(vText, 10);
      raw = shared[idx] ?? "";
      value = raw;
    } else if (t === "b") {
      value = vText === "1";
      raw = value ? "TRUE" : "FALSE";
    } else if (t === "inlineStr") {
      const isEl = c.getElementsByTagName("is")[0];
      raw = isEl ? Array.from(isEl.getElementsByTagName("t")).map((x) => x.textContent ?? "").join("") : vText;
      value = raw;
    } else if (t === "str") {
      raw = vText;
      value = raw;
    } else if (t === "e") {
      value = new CellError((vText as CellError["kind"]) || "#VALUE!");
      raw = vText;
    } else {
      // numeric (default). Guard against non-numeric junk → treat as text.
      if (vText === "") {
        value = null;
      } else {
        const n = Number(vText);
        if (Number.isNaN(n)) {
          value = vText;
          raw = vText;
        } else {
          value = n;
          raw = vText;
        }
      }
    }

    const cell = sheet.ensureCell(pos.row, pos.col);
    cell.raw = raw;
    cell.value = value;
    if (styleIdx >= 0) {
      const f = styles.formats.get(styleIdx);
      if (f) cell.format = { ...f };
    }
  }

  // notes/comments
  for (const [ref, text] of comments) {
    const p = parseA1(ref);
    if (p) sheet.ensureCell(p.row, p.col).note = text;
  }
}

function parseSqref(sqref: string): { r0: number; c0: number; r1: number; c1: number } | null {
  const first = sqref.split(" ")[0]; // ignore multi-area; take first
  const [a, b] = first.split(":");
  const p0 = parseA1(a);
  const p1 = parseA1(b || a);
  if (!p0 || !p1) return null;
  return { r0: Math.min(p0.row, p1.row), c0: Math.min(p0.col, p1.col), r1: Math.max(p0.row, p1.row), c1: Math.max(p0.col, p1.col) };
}

function parseDataValidation(dv: Element, sheet: Sheet) {
  const range = parseSqref(dv.getAttribute("sqref") || "");
  if (!range) return;
  const type = dv.getAttribute("type") || "";
  const f1 = dv.getElementsByTagName("formula1")[0]?.textContent ?? "";
  const f2 = dv.getElementsByTagName("formula2")[0]?.textContent ?? "";
  if (type === "list") {
    const src = f1.replace(/^"|"$/g, "").split(",").map((s) => s.trim()).filter((s) => s !== "");
    sheet.validations.push({ range, type: "list", source: src, allowBlank: true });
  } else if (type === "textLength") {
    sheet.validations.push({ range, type: "textLength", operator: opMapRev[dv.getAttribute("operator") || "between"] || "between", min: parseFloat(f1) || 0, max: parseFloat(f2) || undefined });
  } else if (type === "whole" || type === "decimal") {
    sheet.validations.push({ range, type: "number", operator: opMapRev[dv.getAttribute("operator") || "between"] || "between", min: parseFloat(f1) || 0, max: f2 ? parseFloat(f2) : undefined });
  }
}

function parseCfRule(rule: Element, range: { r0: number; c0: number; r1: number; c1: number }, dxfs: CellFormat[]): ConditionalRule | null {
  const type = rule.getAttribute("type") || "";
  const dxfId = rule.getAttribute("dxfId");
  const format = dxfId !== null ? dxfs[parseInt(dxfId, 10)] : undefined;
  const id = `cf-${range.r0}-${range.c0}-${rule.getAttribute("priority") || "0"}`;
  const formulas = Array.from(rule.getElementsByTagName("formula")).map((f) => f.textContent ?? "");
  const base = { id, range, format };
  if (type === "cellIs") {
    const op = rule.getAttribute("operator") || "";
    const v1: number | string = formulas[0] ?? "";
    const map: Record<string, CondType> = { greaterThan: "greaterThan", lessThan: "lessThan", equal: "equalTo", notEqual: "notEqualTo", between: "between" };
    const t = map[op];
    if (!t) return null;
    if (t === "between") return { ...base, type: "between", value1: parseFloat(formulas[0]) || 0, value2: parseFloat(formulas[1]) || 0 };
    return { ...base, type: t, value1: isNaN(Number(v1)) ? v1 : Number(v1) };
  }
  if (type === "containsText") return { ...base, type: "textContains", value1: rule.getAttribute("text") || "" };
  if (type === "duplicateValues") return { ...base, type: "duplicate" };
  if (type === "uniqueValues") return { ...base, type: "unique" };
  if (type === "top10") {
    const n = parseInt(rule.getAttribute("rank") || "10", 10);
    return { ...base, type: rule.getAttribute("bottom") === "1" ? "bottom" : "top", n };
  }
  if (type === "colorScale") {
    const cs = rule.getElementsByTagName("colorScale")[0];
    const colors = Array.from(cs?.getElementsByTagName("color") || []).map((c) => argbToHex(c.getAttribute("rgb") || "FF000000"));
    if (colors.length === 3) return { id, range, type: "colorScale", minColor: colors[0], midColor: colors[1], maxColor: colors[2] };
    return { id, range, type: "colorScale", minColor: colors[0] || "#f8696b", maxColor: colors[colors.length - 1] || "#63be7b" };
  }
  if (type === "dataBar") {
    const color = rule.getElementsByTagName("dataBar")[0]?.getElementsByTagName("color")[0]?.getAttribute("rgb");
    return { id, range, type: "dataBar", color: color ? argbToHex(color) : "#638ec6" };
  }
  return null;
}

function argbToHex(argb: string): string {
  // "FFRRGGBB" -> "#RRGGBB"
  if (argb.length === 8) return "#" + argb.slice(2);
  if (argb.length === 6) return "#" + argb;
  return "#000000";
}

// ===================== WRITE =====================

export function writeXlsx(wb: Workbook): Uint8Array {
  const shared: string[] = [];
  const sharedIndex = new Map<string, number>();
  const internString = (s: string) => {
    let i = sharedIndex.get(s);
    if (i === undefined) {
      i = shared.length;
      shared.push(s);
      sharedIndex.set(s, i);
    }
    return i;
  };

  // Build the style table from unique cell formats across all sheets. Sheets
  // must be serialized BEFORE styles.toXml() (they register fonts/fills/borders/
  // dxfs into the style builder).
  const styleBuilder = new StyleBuilder();
  const files: Record<string, Uint8Array> = {};
  const commentSheets: boolean[] = [];

  wb.sheets.forEach((sheet, i) => {
    const hasComments = [...sheet.cells.values()].some((c) => c.note);
    commentSheets[i] = hasComments;
    files[`xl/worksheets/sheet${i + 1}.xml`] = strToU8(sheetToXml(sheet, internString, styleBuilder, hasComments));
    if (hasComments) {
      files[`xl/comments${i + 1}.xml`] = strToU8(commentsXml(sheet));
      files[`xl/worksheets/_rels/sheet${i + 1}.xml.rels`] = strToU8(worksheetRelsXml(i + 1));
    }
  });

  files["[Content_Types].xml"] = strToU8(contentTypesXml(wb.sheets.length, commentSheets));
  files["_rels/.rels"] = strToU8(rootRelsXml());
  files["xl/workbook.xml"] = strToU8(workbookXml(wb));
  files["xl/_rels/workbook.xml.rels"] = strToU8(workbookRelsXml(wb.sheets.length));
  files["xl/styles.xml"] = strToU8(styleBuilder.toXml());
  files["xl/sharedStrings.xml"] = strToU8(sharedStringsXml(shared));

  return zipSync(files, { level: 6 });
}

function worksheetRelsXml(n: number): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments${n}.xml"/>` +
    `</Relationships>`
  );
}

function sheetToXml(
  sheet: Sheet,
  internString: (s: string) => number,
  styles: StyleBuilder,
  hasComments: boolean
): string {
  // group cells by row
  const rowsMap = new Map<number, { col: number; key: string }[]>();
  let maxCol = 0;
  for (const key of sheet.cells.keys()) {
    const { row, col } = parseKeyLocal(key);
    if (!rowsMap.has(row)) rowsMap.set(row, []);
    rowsMap.get(row)!.push({ col, key });
    if (col > maxCol) maxCol = col;
  }
  // rows that need attributes even without cells (custom height / hidden)
  const rowSet = new Set<number>(rowsMap.keys());
  for (const r of sheet.rowHeights.keys()) rowSet.add(r);
  for (const r of sheet.hiddenRows) rowSet.add(r);
  const sortedRows = [...rowSet].sort((a, b) => a - b);

  let body = "";
  for (const r of sortedRows) {
    const cells = (rowsMap.get(r) || []).sort((a, b) => a.col - b.col);
    let rowXml = "";
    for (const { col, key } of cells) {
      const cell = sheet.cells.get(key)!;
      const ref = colToLetter(col) + (r + 1);
      const sIdx = cell.format ? styles.indexFor(cell.format) : 0;
      const sAttr = sIdx ? ` s="${sIdx}"` : "";
      rowXml += cellToXml(cell, ref, sAttr, internString);
    }
    const ht = sheet.rowHeights.get(r);
    const htAttr = ht !== undefined ? ` ht="${(ht / PX_PER_PT).toFixed(2)}" customHeight="1"` : "";
    const hidAttr = sheet.hiddenRows.has(r) ? ` hidden="1"` : "";
    body += `<row r="${r + 1}"${htAttr}${hidAttr}>${rowXml}</row>`;
  }

  // columns (widths + hidden)
  let colsXml = "";
  const colSet = new Set<number>([...sheet.colWidths.keys(), ...sheet.hiddenCols]);
  if (colSet.size) {
    let entries = "";
    for (const c of [...colSet].sort((a, b) => a - b)) {
      const px = sheet.colWidths.get(c);
      const widthAttr = px !== undefined ? ` width="${Math.max(0, (px - 5) / 7).toFixed(2)}" customWidth="1"` : ` width="8.43"`;
      const hidAttr = sheet.hiddenCols.has(c) ? ` hidden="1"` : "";
      entries += `<col min="${c + 1}" max="${c + 1}"${widthAttr}${hidAttr}/>`;
    }
    colsXml = `<cols>${entries}</cols>`;
  }

  // freeze panes
  let sheetViews = "";
  if (sheet.frozenRows || sheet.frozenCols) {
    const topLeft = formatA1(sheet.frozenRows, sheet.frozenCols);
    const x = sheet.frozenCols ? ` xSplit="${sheet.frozenCols}"` : "";
    const y = sheet.frozenRows ? ` ySplit="${sheet.frozenRows}"` : "";
    sheetViews = `<sheetViews><sheetView workbookViewId="0"><pane${x}${y} topLeftCell="${topLeft}" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>`;
  }

  // merges
  let mergeXml = "";
  if (sheet.merges.length) {
    const items = sheet.merges
      .map((m) => `<mergeCell ref="${formatA1(m.r0, m.c0)}:${formatA1(m.r1, m.c1)}"/>`)
      .join("");
    mergeXml = `<mergeCells count="${sheet.merges.length}">${items}</mergeCells>`;
  }

  // conditional formatting
  let cfXml = "";
  let priority = 1;
  for (const rule of sheet.conditionalRules) {
    const sqref = `${formatA1(rule.range.r0, rule.range.c0)}:${formatA1(rule.range.r1, rule.range.c1)}`;
    cfXml += `<conditionalFormatting sqref="${sqref}">${cfRuleXml(rule, styles, priority++)}</conditionalFormatting>`;
  }

  // data validation
  let dvXml = "";
  if (sheet.validations.length) {
    const items = sheet.validations.map((v) => dataValidationXml(v)).join("");
    dvXml = `<dataValidations count="${sheet.validations.length}">${items}</dataValidations>`;
  }

  void hasComments;
  const dim = `A1:${colToLetter(Math.max(maxCol, 0))}${(sortedRows[sortedRows.length - 1] ?? 0) + 1}`;
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<dimension ref="${dim}"/>${sheetViews}${colsXml}<sheetData>${body}</sheetData>` +
    `${mergeXml}${cfXml}${dvXml}</worksheet>`
  );
}

function dataValidationXml(v: DataValidation): string {
  const sqref = `${formatA1(v.range.r0, v.range.c0)}:${formatA1(v.range.r1, v.range.c1)}`;
  if (v.type === "list") {
    const list = (v.source || []).join(",");
    return `<dataValidation type="list" allowBlank="1" sqref="${sqref}"><formula1>"${xmlEscape(list)}"</formula1></dataValidation>`;
  }
  const type = v.type === "textLength" ? "textLength" : "decimal";
  const op = opMap[v.operator || "between"] || "between";
  const f1 = `<formula1>${v.min ?? 0}</formula1>`;
  const f2 = op === "between" || op === "notBetween" ? `<formula2>${v.max ?? 0}</formula2>` : "";
  return `<dataValidation type="${type}" operator="${op}" allowBlank="1" sqref="${sqref}">${f1}${f2}</dataValidation>`;
}

function cfRuleXml(rule: ConditionalRule, styles: StyleBuilder, priority: number): string {
  const dxf = () => (rule.format ? ` dxfId="${styles.dxfFor(rule.format)}"` : "");
  const p = ` priority="${priority}"`;
  switch (rule.type) {
    case "greaterThan":
    case "lessThan":
    case "equalTo":
    case "notEqualTo": {
      const opx = { greaterThan: "greaterThan", lessThan: "lessThan", equalTo: "equal", notEqualTo: "notEqual" }[rule.type];
      return `<cfRule type="cellIs" operator="${opx}"${dxf()}${p}><formula>${Number(rule.value1) || 0}</formula></cfRule>`;
    }
    case "between":
      return `<cfRule type="cellIs" operator="between"${dxf()}${p}><formula>${Number(rule.value1) || 0}</formula><formula>${rule.value2 ?? 0}</formula></cfRule>`;
    case "textContains":
      return `<cfRule type="containsText" operator="containsText" text="${xmlEscape(String(rule.value1 ?? ""))}"${dxf()}${p}><formula>NOT(ISERROR(SEARCH("${xmlEscape(String(rule.value1 ?? ""))}",${formatA1(rule.range.r0, rule.range.c0)})))</formula></cfRule>`;
    case "duplicate":
      return `<cfRule type="duplicateValues"${dxf()}${p}/>`;
    case "unique":
      return `<cfRule type="uniqueValues"${dxf()}${p}/>`;
    case "top":
      return `<cfRule type="top10" rank="${rule.n ?? 10}"${dxf()}${p}/>`;
    case "bottom":
      return `<cfRule type="top10" bottom="1" rank="${rule.n ?? 10}"${dxf()}${p}/>`;
    case "colorScale": {
      const min = hexToArgb(rule.minColor || "#f8696b");
      const mid = rule.midColor ? hexToArgb(rule.midColor) : null;
      const max = hexToArgb(rule.maxColor || "#63be7b");
      const cfvo = mid
        ? `<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>`
        : `<cfvo type="min"/><cfvo type="max"/>`;
      const colors = mid ? `<color rgb="${min}"/><color rgb="${mid}"/><color rgb="${max}"/>` : `<color rgb="${min}"/><color rgb="${max}"/>`;
      return `<cfRule type="colorScale"${p}><colorScale>${cfvo}${colors}</colorScale></cfRule>`;
    }
    case "dataBar":
      return `<cfRule type="dataBar"${p}><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="${hexToArgb(rule.color || "#638ec6")}"/></dataBar></cfRule>`;
    default:
      return "";
  }
}

function commentsXml(sheet: Sheet): string {
  let list = "";
  for (const [key, cell] of sheet.cells) {
    if (!cell.note) continue;
    const { row, col } = parseKeyLocal(key);
    list += `<comment ref="${formatA1(row, col)}" authorId="0"><text><r><t xml:space="preserve">${xmlEscape(cell.note)}</t></r></text></comment>`;
  }
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
    `<authors><author>BoringSheets</author></authors><commentList>${list}</commentList></comments>`
  );
}

function cellToXml(
  cell: { raw: string; value: CellValue },
  ref: string,
  sAttr: string,
  internString: (s: string) => number
): string {
  const raw = cell.raw;
  if (raw.startsWith("=")) {
    const formula = xmlEscape(raw.slice(1));
    const v = cell.value;
    let vXml = "";
    let tAttr = "";
    if (typeof v === "number") vXml = `<v>${v}</v>`;
    else if (typeof v === "boolean") {
      tAttr = ` t="b"`;
      vXml = `<v>${v ? 1 : 0}</v>`;
    } else if (v instanceof CellError) {
      tAttr = ` t="e"`;
      vXml = `<v>${v.kind}</v>`;
    } else if (typeof v === "string") {
      tAttr = ` t="str"`;
      vXml = `<v>${xmlEscape(v)}</v>`;
    }
    return `<c r="${ref}"${sAttr}${tAttr}><f>${formula}</f>${vXml}</c>`;
  }

  const v = cell.value;
  if (v === null) return sAttr ? `<c r="${ref}"${sAttr}/>` : "";
  if (typeof v === "number") return `<c r="${ref}"${sAttr}><v>${v}</v></c>`;
  if (typeof v === "boolean") return `<c r="${ref}"${sAttr} t="b"><v>${v ? 1 : 0}</v></c>`;
  if (v instanceof CellError) return `<c r="${ref}"${sAttr} t="e"><v>${v.kind}</v></c>`;
  // string -> shared string
  const idx = internString(String(v));
  return `<c r="${ref}"${sAttr} t="s"><v>${idx}</v></c>`;
}

// --- style table builder ---

class StyleBuilder {
  private numFmts: string[] = [];
  private fonts: string[] = ['<font><sz val="10"/><name val="Arial"/></font>'];
  private fills: string[] = [
    '<fill><patternFill patternType="none"/></fill>',
    '<fill><patternFill patternType="gray125"/></fill>',
  ];
  private borders: string[] = ["<border><left/><right/><top/><bottom/><diagonal/></border>"];
  private xfs: string[] = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  private dxfs: string[] = [];
  private cache = new Map<string, number>();
  private borderCache = new Map<string, number>();
  private dxfCache = new Map<string, number>();

  private borderId(b: BorderSet): number {
    const k = JSON.stringify(b);
    const c = this.borderCache.get(k);
    if (c !== undefined) return c;
    const argb = hexToArgb(b.color || "#000000");
    const side = (on?: boolean, tag = "left") =>
      on ? `<${tag} style="thin"><color rgb="${argb}"/></${tag}>` : `<${tag}/>`;
    const xml = `<border>${side(b.left, "left")}${side(b.right, "right")}${side(b.top, "top")}${side(b.bottom, "bottom")}<diagonal/></border>`;
    const id = this.borders.length;
    this.borders.push(xml);
    this.borderCache.set(k, id);
    return id;
  }

  // A differential format for conditional formatting (font + fill).
  dxfFor(fmt: CellFormat): number {
    const k = JSON.stringify(fmt);
    const c = this.dxfCache.get(k);
    if (c !== undefined) return c;
    let font = "";
    if (fmt.bold || fmt.italic || fmt.color) {
      font = `<font>${fmt.bold ? "<b/>" : ""}${fmt.italic ? "<i/>" : ""}${fmt.color ? `<color rgb="${hexToArgb(fmt.color)}"/>` : ""}</font>`;
    }
    // NOTE: conditional-format fills use <bgColor> (not fgColor).
    const fill = fmt.bg ? `<fill><patternFill><bgColor rgb="${hexToArgb(fmt.bg)}"/></patternFill></fill>` : "";
    const id = this.dxfs.length;
    this.dxfs.push(`<dxf>${font}${fill}</dxf>`);
    this.dxfCache.set(k, id);
    return id;
  }

  indexFor(fmt: CellFormat): number {
    const k = JSON.stringify(fmt);
    const cached = this.cache.get(k);
    if (cached !== undefined) return cached;

    let numFmtId = 0;
    if (fmt.numFmt && fmt.numFmt !== "General") {
      numFmtId = 164 + this.numFmts.length;
      this.numFmts.push(`<numFmt numFmtId="${numFmtId}" formatCode="${xmlEscape(fmt.numFmt)}"/>`);
    }

    let fontId = 0;
    if (fmt.bold || fmt.italic || fmt.underline || fmt.strike || fmt.color || fmt.fontSize || fmt.fontFamily) {
      const parts: string[] = [];
      if (fmt.bold) parts.push("<b/>");
      if (fmt.italic) parts.push("<i/>");
      if (fmt.underline) parts.push("<u/>");
      if (fmt.strike) parts.push("<strike/>");
      // model stores px; Excel font size is in points
      parts.push(`<sz val="${fmt.fontSize ? Math.round(fmt.fontSize * 0.75 * 100) / 100 : 10}"/>`);
      if (fmt.color) parts.push(`<color rgb="${hexToArgb(fmt.color)}"/>`);
      parts.push(`<name val="${xmlEscape(fmt.fontFamily || "Arial")}"/>`);
      fontId = this.fonts.length;
      this.fonts.push(`<font>${parts.join("")}</font>`);
    }

    let fillId = 0;
    if (fmt.bg) {
      fillId = this.fills.length;
      this.fills.push(
        `<fill><patternFill patternType="solid"><fgColor rgb="${hexToArgb(fmt.bg)}"/></patternFill></fill>`
      );
    }

    const bId = fmt.border ? this.borderId(fmt.border) : 0;

    const alignAttrs = `${fmt.align ? ` horizontal="${fmt.align}"` : ""}${fmt.wrap ? ` wrapText="1"` : ""}`;
    const alignXml = alignAttrs ? `<alignment${alignAttrs}/>` : "";
    const applyAttrs =
      `${numFmtId ? ' applyNumberFormat="1"' : ""}${fontId ? ' applyFont="1"' : ""}` +
      `${fillId ? ' applyFill="1"' : ""}${bId ? ' applyBorder="1"' : ""}${alignXml ? ' applyAlignment="1"' : ""}`;
    const idx = this.xfs.length;
    this.xfs.push(
      `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${bId}" xfId="0"${applyAttrs}>${alignXml}</xf>`
    );
    this.cache.set(k, idx);
    return idx;
  }

  toXml(): string {
    const numFmtsXml = this.numFmts.length
      ? `<numFmts count="${this.numFmts.length}">${this.numFmts.join("")}</numFmts>`
      : "";
    const dxfsXml = this.dxfs.length ? `<dxfs count="${this.dxfs.length}">${this.dxfs.join("")}</dxfs>` : "";
    return (
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
      numFmtsXml +
      `<fonts count="${this.fonts.length}">${this.fonts.join("")}</fonts>` +
      `<fills count="${this.fills.length}">${this.fills.join("")}</fills>` +
      `<borders count="${this.borders.length}">${this.borders.join("")}</borders>` +
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
      `<cellXfs count="${this.xfs.length}">${this.xfs.join("")}</cellXfs>` +
      `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
      dxfsXml +
      `</styleSheet>`
    );
  }
}

function hexToArgb(hex: string): string {
  const h = hex.replace("#", "");
  return ("FF" + h).toUpperCase();
}

function parseKeyLocal(key: string): { row: number; col: number } {
  const i = key.indexOf(",");
  return { row: parseInt(key.slice(0, i), 10), col: parseInt(key.slice(i + 1), 10) };
}

// --- static package parts ---

function contentTypesXml(sheetCount: number, commentSheets: boolean[]): string {
  let overrides = "";
  for (let i = 1; i <= sheetCount; i++) {
    overrides += `<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
    if (commentSheets[i - 1]) {
      overrides += `<Override PartName="/xl/comments${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`;
    }
  }
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
    `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
    `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>` +
    overrides +
    `</Types>`
  );
}

function rootRelsXml(): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
    `</Relationships>`
  );
}

function workbookXml(wb: Workbook): string {
  let entries = "";
  wb.sheets.forEach((s, i) => {
    entries += `<sheet name="${xmlEscape(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`;
  });
  let names = "";
  if (wb.names.length) {
    const items = wb.names
      .map((n) => {
        const sheet = wb.sheets.find((s) => s.id === n.sheetId);
        if (!sheet) return "";
        const ref =
          formatSheetName(sheet.name) +
          "!" +
          formatA1(n.r0, n.c0, true, true) +
          (n.r0 === n.r1 && n.c0 === n.c1 ? "" : ":" + formatA1(n.r1, n.c1, true, true));
        return `<definedName name="${xmlEscape(n.name)}">${xmlEscape(ref)}</definedName>`;
      })
      .join("");
    names = `<definedNames>${items}</definedNames>`;
  }
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<sheets>${entries}</sheets>${names}</workbook>`
  );
}

function workbookRelsXml(sheetCount: number): string {
  let rels = "";
  for (let i = 1; i <= sheetCount; i++) {
    rels += `<Relationship Id="rId${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i}.xml"/>`;
  }
  const styleId = sheetCount + 1;
  const sharedId = sheetCount + 2;
  rels += `<Relationship Id="rId${styleId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;
  rels += `<Relationship Id="rId${sharedId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`;
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`
  );
}

function sharedStringsXml(strings: string[]): string {
  const items = strings
    .map((s) => `<si><t xml:space="preserve">${xmlEscape(s)}</t></si>`)
    .join("");
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">` +
    items +
    `</sst>`
  );
}

// CSV/TSV helpers reused by clipboard paste of external Excel data.
export function sheetToTsv(sheet: Sheet, r0: number, c0: number, r1: number, c1: number): string {
  const lines: string[] = [];
  for (let r = r0; r <= r1; r++) {
    const cells: string[] = [];
    for (let c = c0; c <= c1; c++) {
      const cell = sheet.getCell(r, c);
      cells.push(cell ? formatValue(cell.value, cell.format?.numFmt) : "");
    }
    lines.push(cells.join("\t"));
  }
  return lines.join("\n");
}

// exported helper used by callers needing the cell key utility
export { cellKey };
