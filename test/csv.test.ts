// Headless smoke test for the CSV I/O module: RFC 4180 parsing and
// serialization. Bundled with esbuild and run under Node. Exits non-zero on
// failure. Mirrors the harness style in engine.test.ts.

import { parseCsv, toCsv } from "../src/io/csv.js";

let failures = 0;
let count = 0;

function eq(actual: unknown, expected: unknown, label: string) {
  count++;
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a !== e) {
    failures++;
    console.error(`✗ ${label}: expected ${e}, got ${a}`);
  } else {
    console.log(`✓ ${label}`);
  }
}

// --- simple rows ---
eq(parseCsv("a,b,c"), [["a", "b", "c"]], "single row");
eq(parseCsv("a,b\n1,2"), [["a", "b"], ["1", "2"]], "two rows LF");

// --- empty fields ---
eq(parseCsv("a,,c"), [["a", "", "c"]], "empty middle field");
eq(parseCsv(",,"), [["", "", ""]], "all empty fields");

// --- quoted fields with commas ---
eq(parseCsv('"a,b",c'), [["a,b", "c"]], "quoted comma");
eq(parseCsv('x,"y,z,w"'), [["x", "y,z,w"]], "quoted trailing comma field");

// --- doubled quotes ---
eq(parseCsv('"he said ""hi"""'), [['he said "hi"']], "escaped doubled quote");
eq(parseCsv('"""",a'), [['"', "a"]], "field of single quote char");

// --- embedded newlines in quotes ---
eq(parseCsv('"line1\nline2",b'), [["line1\nline2", "b"]], "embedded LF in quotes");
eq(parseCsv('"a\r\nb"'), [["a\r\nb"]], "embedded CRLF in quotes");
eq(parseCsv('"a\rb"'), [["a\rb"]], "embedded CR in quotes");

// --- CRLF vs LF vs CR row separators ---
eq(parseCsv("a\r\nb"), [["a"], ["b"]], "CRLF row separator");
eq(parseCsv("a\rb"), [["a"], ["b"]], "CR row separator");

// --- trailing newline handling ---
eq(parseCsv("a,b\n"), [["a", "b"]], "single trailing LF no extra row");
eq(parseCsv("a,b\r\n"), [["a", "b"]], "single trailing CRLF no extra row");
eq(parseCsv("a\n\n"), [["a"], [""]], "two trailing LFs => one empty row");

// --- toCsv serialization ---
eq(toCsv([["a", "b"], ["c", "d"]]), "a,b\r\nc,d", "basic serialize CRLF join");
eq(toCsv([["a,b", 'c"d']]), '"a,b","c""d"', "quote comma and double quotes");
eq(toCsv([[1, true, false, null]]), "1,TRUE,FALSE,", "number/boolean/null encoding");
eq(toCsv([["a\nb", "x\r\ny"]]), '"a\nb","x\r\ny"', "quote embedded newlines");

// --- round-trip parseCsv -> toCsv -> parseCsv ---
const original = 'name,note\r\n"Smith, J","said ""hi""\nbye"\r\nplain,';
const parsed1 = parseCsv(original);
const reser = toCsv(parsed1);
const parsed2 = parseCsv(reser);
eq(parsed2, parsed1, "round-trip parse->toCsv->parse stable");
eq(parsed1, [["name", "note"], ["Smith, J", 'said "hi"\nbye'], ["plain", ""]], "round-trip content");

console.log(`\n${count - failures}/${count} passed`);
if (failures) {
  console.error(`${failures} FAILED`);
  process.exit(1);
}
