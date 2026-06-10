// Reader robustness corpus: hand-crafted .xlsx packages that mimic the
// structural variations real Excel / LibreOffice / exporters produce (which our
// own writer does NOT emit), to make sure the reader handles them gracefully.

import { DOMParser } from "@xmldom/xmldom";
(globalThis as unknown as { DOMParser: unknown }).DOMParser = DOMParser;

import { zipSync, strToU8 } from "fflate";
import { readXlsx } from "../src/io/xlsx.js";
import { CalcEngine } from "../src/engine/calc.js";
import { CellError } from "../src/model/types.js";

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

const CT = (sheets = 1) =>
  `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
  Array.from({ length: sheets }, (_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("") +
  `</Types>`;
const ROOT_RELS = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;

// Build + read a single-sheet workbook with the given sheetData inner XML.
function read(opts: {
  cells: string;
  shared?: string[];
  styles?: string;
  sheetPath?: string;
  workbook?: string;
  wbRels?: string;
}) {
  const sheetPath = opts.sheetPath || "xl/worksheets/sheet1.xml";
  const files: Record<string, Uint8Array> = {};
  files["[Content_Types].xml"] = strToU8(CT(1));
  files["_rels/.rels"] = strToU8(ROOT_RELS);
  files["xl/workbook.xml"] = strToU8(
    opts.workbook ||
      `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`
  );
  files["xl/_rels/workbook.xml.rels"] = strToU8(
    opts.wbRels ||
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`
  );
  if (opts.shared) {
    files["xl/sharedStrings.xml"] = strToU8(
      `<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${opts.shared.map((s) => `<si><t>${s}</t></si>`).join("")}</sst>`
    );
  }
  if (opts.styles) files["xl/styles.xml"] = strToU8(opts.styles);
  files[sheetPath] = strToU8(
    `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${opts.cells}</sheetData></worksheet>`
  );
  const bytes = zipSync(files);
  const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  const wb = readXlsx(ab);
  new CalcEngine(wb).rebuild();
  return wb.active;
}
const v = (s: ReturnType<typeof read>, a1: string) => {
  const m = /^([A-Z]+)(\d+)$/.exec(a1)!;
  const col = m[1].split("").reduce((n, c) => n * 26 + (c.charCodeAt(0) - 64), 0) - 1;
  return s.getCell(parseInt(m[2]) - 1, col)?.value ?? null;
};

// 1. rich inline string (multiple runs)
ok(v(read({ cells: `<row r="1"><c r="A1" t="inlineStr"><is><r><t>Hel</t></r><r><t>lo</t></r></is></c></row>` }), "A1") === "Hello", "rich inlineStr joins runs");

// 2. plain inline string
ok(v(read({ cells: `<row r="1"><c r="A1" t="inlineStr"><is><t>Hi</t></is></c></row>` }), "A1") === "Hi", "plain inlineStr");

// 3. shared-formula dependent keeps cached value (no broken formula)
{
  const s = read({ cells: `<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>2</v></c></row>` });
  ok(v(s, "A2") === 2, "shared-formula dependent keeps cached value");
  ok(v(s, "A1") === 2, "shared-formula master recomputes");
}

// 4. shared string reference
ok(v(read({ cells: `<row r="1"><c r="A1" t="s"><v>1</v></c></row>`, shared: ["zero", "one"] }), "A1") === "one", "shared string index");

// 5. boolean / error
ok(v(read({ cells: `<row r="1"><c r="A1" t="b"><v>1</v></c></row>` }), "A1") === true, "boolean cell");
{
  const e = v(read({ cells: `<row r="1"><c r="A1" t="e"><v>#DIV/0!</v></c></row>` }), "A1");
  ok(e instanceof CellError && e.kind === "#DIV/0!", "error cell");
}

// 6. whitespace + scientific + negative numbers
ok(v(read({ cells: `<row r="1"><c r="A1"><v> 42 </v></c></row>` }), "A1") === 42, "whitespace in <v>");
ok(v(read({ cells: `<row r="1"><c r="A1"><v>1.5e3</v></c></row>` }), "A1") === 1500, "scientific notation");
ok(v(read({ cells: `<row r="1"><c r="A1"><v>-12.5</v></c></row>` }), "A1") === -12.5, "negative number");

// 7. self-closing empty cell + sparse row
{
  const s = read({ cells: `<row r="1"><c r="A1"><v>5</v></c><c r="C1"/></row>` });
  ok(v(s, "A1") === 5 && v(s, "C1") === null, "self-closing empty cell");
}

// 8. date serial with builtin numFmt 14
{
  const styles = `<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs></styleSheet>`;
  const s = read({ cells: `<row r="1"><c r="A1" s="1"><v>44927</v></c></row>`, styles });
  ok(v(s, "A1") === 44927, "date serial value");
  ok(s.getCell(0, 0)?.format?.numFmt === "mm/dd/yyyy", "date builtin numFmt applied");
}

// 9. themed/indexed colors must not crash, value still reads
{
  const styles = `<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font/><font><color theme="1"/></font></fonts><fills count="3"><fill/><fill/><fill><patternFill patternType="solid"><fgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellXfs count="2"><xf numFmtId="0"/><xf fontId="1" fillId="2" applyFont="1" applyFill="1"/></cellXfs></styleSheet>`;
  ok(v(read({ cells: `<row r="1"><c r="A1" s="1"><v>7</v></c></row>`, styles }), "A1") === 7, "themed/indexed colors don't crash");
}

// 10. t="s" but NO sharedStrings part → graceful empty (null or "")
{
  const got = v(read({ cells: `<row r="1"><c r="A1" t="s"><v>0</v></c></row>` }), "A1");
  ok(got === "" || got === null, "missing sharedStrings degrades to empty without crashing");
}

// 11. rel-based, non-standard sheet filename
{
  const wbRels = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/data.xml"/></Relationships>`;
  const s = read({ cells: `<row r="1"><c r="A1"><v>99</v></c></row>`, sheetPath: "xl/worksheets/data.xml", wbRels });
  ok(v(s, "A1") === 99, "rel-based non-standard sheet filename");
}

// 12. out-of-order rows / cols
{
  const s = read({ cells: `<row r="3"><c r="B3"><v>3</v></c></row><row r="1"><c r="A1"><v>1</v></c></row>` });
  ok(v(s, "A1") === 1 && v(s, "B3") === 3, "out-of-order rows");
}

console.log(`\n${count - failures}/${count} passed`);
if (failures) process.exit(1);
