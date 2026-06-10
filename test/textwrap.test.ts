import { wrapLines } from "../src/engine/textwrap.js";
let failures = 0, count = 0;
function eq(a: unknown, e: unknown, l: string){count++; if(JSON.stringify(a)!==JSON.stringify(e)){failures++;console.error(`✗ ${l}: got ${JSON.stringify(a)} want ${JSON.stringify(e)}`);}else console.log(`✓ ${l}`);}
// measure = 1px per char
const m = (s: string) => s.length;
eq(wrapLines(m, "hello world foo", 11), ["hello world", "foo"], "wrap on spaces");
eq(wrapLines(m, "hello world foo", 100), ["hello world foo"], "no wrap when fits");
eq(wrapLines(m, "abcdefghij", 4), ["abcd","efgh","ij"], "hard-break long token");
eq(wrapLines(m, "UPIFIN20260401", 6), ["UPIFIN","202604","01"], "hard-break reference number");
eq(wrapLines(m, "a\nb", 50), ["a","b"], "honor explicit newline");
eq(wrapLines(m, "one two", 3).length, 2, "two short words wrap");
console.log(`\n${count-failures}/${count} passed`); if(failures) process.exit(1);
