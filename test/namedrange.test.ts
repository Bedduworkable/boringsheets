import { Workbook } from "../src/model/workbook.js";
import { CalcEngine } from "../src/engine/calc.js";
import { CellError } from "../src/model/types.js";
let failures = 0, count = 0;
function eq(a: unknown, e: unknown, l: string){count++; const x = a instanceof CellError ? a.kind : a; if(x!==e){failures++;console.error(`✗ ${l}: got ${JSON.stringify(x)} want ${JSON.stringify(e)}`);}else console.log(`✓ ${l}`);}
const rc=(a1:string)=>{const m=/^([A-Z]+)(\d+)$/.exec(a1)!;return{row:parseInt(m[2])-1,col:m[1].split("").reduce((n,c)=>n*26+(c.charCodeAt(0)-64),0)-1};};
const wb = new Workbook(); const sheet = wb.active; const e = new CalcEngine(wb);
const set=(a1:string,raw:string)=>{const{row,col}=rc(a1);e.setCellRaw(row,col,raw);};
const val=(a1:string)=>{const{row,col}=rc(a1);return sheet.getCell(row,col)?.value??null;};

set("A1","10"); set("A2","20"); set("A3","30");
// define names BEFORE formulas reference them
wb.names.push({name:"Data", sheetId:sheet.id, r0:0,c0:0,r1:2,c1:0});
wb.names.push({name:"Rate", sheetId:sheet.id, r0:0,c0:0,r1:0,c1:0});

set("B1","=SUM(Data)");        eq(val("B1"),60,"SUM over named range");
set("B2","=AVERAGE(Data)");    eq(val("B2"),20,"AVERAGE over named range");
set("B3","=Rate*2");           eq(val("B3"),20,"single-cell name in arithmetic");
set("B4","=nope");             eq(val("B4"),"#NAME?","unknown name → #NAME?");

// dependency through the name: change a member, dependents recompute
set("A1","100");
eq(val("B1"),150,"named-range SUM recomputes after member change");
eq(val("B3"),200,"single-cell name recomputes");

console.log(`\n${count-failures}/${count} passed`); if(failures) process.exit(1);
