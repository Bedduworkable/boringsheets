// Whole-column/row references + reference functions (ROW/COLUMN/ADDRESS/OFFSET/
// INDIRECT) + SUBTOTAL/AGGREGATE.
import { Workbook } from "../src/model/workbook.js";
import { CalcEngine } from "../src/engine/calc.js";
import { CellError } from "../src/model/types.js";
let failures = 0, count = 0;
function eq(a: unknown, e: unknown, l: string){count++; const x=a instanceof CellError?a.kind:a; const ok=x===e||(typeof x==="number"&&typeof e==="number"&&Math.abs(x-e)<1e-9); if(!ok){failures++;console.error(`✗ ${l}: got ${JSON.stringify(x)} want ${JSON.stringify(e)}`);}else console.log(`✓ ${l}`);}
const wb=new Workbook(); const s=wb.active; const e=new CalcEngine(wb);
const rc=(a1:string)=>{const m=/^([A-Z]+)(\d+)$/.exec(a1)!;return{row:parseInt(m[2])-1,col:m[1].split("").reduce((n,c)=>n*26+(c.charCodeAt(0)-64),0)-1};};
const set=(a1:string,raw:string)=>{const{row,col}=rc(a1);e.setCellRaw(row,col,raw);};
const val=(a1:string)=>{const{row,col}=rc(a1);return s.getCell(row,col)?.value??null;};
set("A1","10");set("A2","20");set("A3","30");set("B1","5");set("B2","15");set("B3","25");

set("D1","=SUM(A:A)"); eq(val("D1"),60,"SUM whole column A:A");
set("D2","=SUM(A:B)"); eq(val("D2"),105,"SUM two whole columns A:B");
set("H20","1");set("I20","2");set("J20","3");
set("K21","=SUM(20:20)"); eq(val("K21"),6,"SUM whole row 20:20");
set("D4","=COUNT(A:A)"); eq(val("D4"),3,"COUNT whole column");
set("D5","=ROW(A5)"); eq(val("D5"),5,"ROW(ref)");
set("D6","=COLUMN(C1)"); eq(val("D6"),3,"COLUMN(ref)");
set("D7","=ADDRESS(2,3)"); eq(val("D7"),"$C$2","ADDRESS absolute");
set("D8","=ADDRESS(2,3,4)"); eq(val("D8"),"C2","ADDRESS relative");
set("D9","=OFFSET(A1,1,0)"); eq(val("D9"),20,"OFFSET single cell");
set("D10","=SUM(OFFSET(A1,0,0,3,1))"); eq(val("D10"),60,"OFFSET range into SUM");
set("D11","=INDIRECT(\"A2\")"); eq(val("D11"),20,"INDIRECT single cell");
set("D12","=SUM(INDIRECT(\"A1:A3\"))"); eq(val("D12"),60,"INDIRECT range");
set("D13","=SUBTOTAL(9,A1:A3)"); eq(val("D13"),60,"SUBTOTAL sum");
set("D14","=SUBTOTAL(1,A1:A3)"); eq(val("D14"),20,"SUBTOTAL average");
set("D15","=AGGREGATE(9,0,A1:A3)"); eq(val("D15"),60,"AGGREGATE sum");
set("D16","=AGGREGATE(14,0,A1:A3,1)"); eq(val("D16"),30,"AGGREGATE large");
// dependency: whole-column SUM recalculates when a member changes
set("A2","200"); eq(val("D1"),240,"whole-column SUM recomputes on member change");
// ROW() no-arg uses the formula's own cell
set("F7","=ROW()"); eq(val("F7"),7,"ROW() no-arg uses current cell");
set("F8","=COLUMN()"); eq(val("F8"),6,"COLUMN() no-arg uses current cell");
console.log(`\n${count-failures}/${count} passed`); if(failures)process.exit(1);
