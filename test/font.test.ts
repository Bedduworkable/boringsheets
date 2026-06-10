import { DOMParser } from "@xmldom/xmldom";
(globalThis as unknown as { DOMParser: unknown }).DOMParser = DOMParser;
import { Workbook } from "../src/model/workbook.js";
import { CalcEngine } from "../src/engine/calc.js";
import { writeXlsx, readXlsx } from "../src/io/xlsx.js";
let failures=0,count=0;
function eq(a:unknown,e:unknown,l:string){count++; if(a!==e){failures++;console.error(`✗ ${l}: got ${JSON.stringify(a)} want ${JSON.stringify(e)}`);}else console.log(`✓ ${l}`);}
const wb=new Workbook(); const s=wb.active; const e=new CalcEngine(wb);
e.setCellRaw(0,0,"Hello"); s.ensureCell(0,0).format={fontFamily:"Georgia", fontSize:20};
e.setCellRaw(0,1,"plain");
const bytes=writeXlsx(wb); const ab=bytes.buffer.slice(bytes.byteOffset,bytes.byteOffset+bytes.byteLength) as ArrayBuffer;
const wb2=readXlsx(ab);
const f=wb2.sheets[0].getCell(0,0)?.format;
eq(f?.fontFamily, "Georgia", "font family round-trips");
eq(f?.fontSize, 20, "font size (px) round-trips via points");
eq(wb2.sheets[0].getCell(0,1)?.format?.fontSize, undefined, "plain cell has no font size");
console.log(`\n${count-failures}/${count} passed`); if(failures)process.exit(1);
