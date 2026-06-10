// Filter recompute logic: hide a row if ANY column's filter excludes its value.
import { Workbook } from "../src/model/workbook.js";
import { CalcEngine } from "../src/engine/calc.js";
import { formatValue } from "../src/engine/format.js";
let failures=0,count=0;
function eq(a:unknown,e:unknown,l:string){count++; if(a!==e){failures++;console.error(`✗ ${l}: got ${JSON.stringify(a)} want ${JSON.stringify(e)}`);}else console.log(`✓ ${l}`);}
const wb=new Workbook(); const s=wb.active; const e=new CalcEngine(wb);
const set=(r:number,c:number,v:string)=>e.setCellRaw(r,c,v);
// header
set(0,0,"Fruit"); set(0,1,"Qty");
set(1,0,"Apple"); set(1,1,"10");
set(2,0,"Banana"); set(2,1,"20");
set(3,0,"Apple"); set(3,1,"30");
set(4,0,"Cherry"); set(4,1,"10");
s.filter = { range:{r0:0,c0:0,r1:4,c1:1}, cols:{} };

// replicate App.recomputeFilter (private)
function recompute(){ const f=s.filter!; const cols=Object.keys(f.cols).map(Number).filter(c=>f.cols[c].length);
  for(let r=f.range.r0+1;r<=f.range.r1;r++){ let hide=false; for(const c of cols){ const cell=s.getCell(r,c); const d=formatValue(cell?.value??null,cell?.format?.numFmt); if(f.cols[c].includes(d)){hide=true;break;} } if(hide)s.hiddenRows.add(r); else s.hiddenRows.delete(r);} }

// filter out Banana in col 0
s.filter.cols[0]=["Banana"]; recompute();
eq(s.hiddenRows.has(2), true, "Banana row hidden");
eq(s.hiddenRows.has(1), false, "Apple row visible");
eq(s.hiddenRows.has(4), false, "Cherry row visible");

// add col 1 filter hiding Qty=10 → AND: hide Banana OR Qty 10
s.filter.cols[1]=["10"]; recompute();
eq(s.hiddenRows.has(1), true, "Apple/10 hidden by Qty filter");
eq(s.hiddenRows.has(2), true, "Banana still hidden");
eq(s.hiddenRows.has(3), false, "Apple/30 visible");
eq(s.hiddenRows.has(4), true, "Cherry/10 hidden");

// clear col 0 → only Qty filter remains
delete s.filter.cols[0]; recompute();
eq(s.hiddenRows.has(2), false, "Banana visible after clearing fruit filter");
eq(s.hiddenRows.has(1), true, "Qty 10 still hidden");
console.log(`\n${count-failures}/${count} passed`); if(failures)process.exit(1);
