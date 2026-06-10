// Serializes an AST back to a formula string, and rewrites cell references —
// used when inserting/deleting rows or columns so formulas keep pointing at the
// right cells (and become #REF! when their target is deleted). Sheet-qualified
// references (Sheet2!A1) are preserved and only transformed when they target
// the sheet being edited.

import { Node } from "./ast.js";
import { CellRef, formatA1, formatSheetName, colToLetter } from "./references.js";
import { parseFormula as parse } from "./parser.js";

// Returns a new {row,col} to move a ref, null to delete it (→ #REF!), or
// undefined to leave it unchanged. `sheet` is the ref's sheet qualifier (or
// undefined for an unqualified, same-sheet reference).
export type RefTransform = (
  row: number,
  col: number,
  sheet: string | undefined
) => { row: number; col: number } | null | undefined;

export function astToFormula(node: Node): string {
  return "=" + print(node);
}

function sheetPrefix(sheet?: string): string {
  return sheet ? formatSheetName(sheet) + "!" : "";
}

function print(node: Node): string {
  switch (node.kind) {
    case "num":
      return String(node.value);
    case "str":
      return '"' + node.value.replace(/"/g, '""') + '"';
    case "bool":
      return node.value ? "TRUE" : "FALSE";
    case "ref":
      return sheetPrefix(node.sheet) + refStr(node.ref);
    case "range":
      if (node.fullCol) return sheetPrefix(node.sheet) + colToLetter(node.start.col) + ":" + colToLetter(node.end.col);
      if (node.fullRow) return sheetPrefix(node.sheet) + (node.start.row + 1) + ":" + (node.end.row + 1);
      return sheetPrefix(node.sheet) + refStr(node.start) + ":" + refStr(node.end);
    case "name":
      return node.name;
    case "unary":
      return node.op === "%" ? print(node.operand) + "%" : node.op + print(node.operand);
    case "binary":
      return print(node.left) + node.op + print(node.right);
    case "call":
      return node.name + "(" + node.args.map(print).join(",") + ")";
  }
}

function refStr(ref: CellRef): string {
  return formatA1(ref.row, ref.col, ref.absRow, ref.absCol);
}

// Rewrite a raw formula string ("=...") by applying a transform to every cell
// reference.
export function rewriteFormulaRefs(raw: string, transform: RefTransform): string {
  let ast: Node;
  try {
    ast = parse(raw.slice(1));
  } catch {
    return raw; // leave unparseable formulas untouched
  }
  const out = mapRefs(ast, transform);
  return "=" + printP(out);
}

type PNode = Node | { kind: "referr" };

function mapRefs(node: Node, t: RefTransform): PNode {
  switch (node.kind) {
    case "ref": {
      const r = t(node.ref.row, node.ref.col, node.sheet);
      if (r === undefined) return node;
      if (r === null) return { kind: "referr" };
      return { kind: "ref", sheet: node.sheet, ref: { ...node.ref, row: r.row, col: r.col } };
    }
    case "range": {
      const s = t(node.start.row, node.start.col, node.sheet);
      const e = t(node.end.row, node.end.col, node.sheet);
      if (s === undefined && e === undefined) return node;
      if (s === null && e === null) return { kind: "referr" };
      // clamp a vanished endpoint to the surviving one (Excel shrinks the range)
      const ns = s === null ? e! : s ?? { row: node.start.row, col: node.start.col };
      const ne = e === null ? s! : e ?? { row: node.end.row, col: node.end.col };
      return {
        kind: "range",
        sheet: node.sheet,
        fullCol: node.fullCol,
        fullRow: node.fullRow,
        start: { ...node.start, row: ns.row, col: ns.col },
        end: { ...node.end, row: ne.row, col: ne.col },
      };
    }
    case "unary":
      return { kind: "unary", op: node.op, operand: mapRefs(node.operand, t) as Node };
    case "binary":
      return {
        kind: "binary",
        op: node.op,
        left: mapRefs(node.left, t) as Node,
        right: mapRefs(node.right, t) as Node,
      };
    case "call":
      return { kind: "call", name: node.name, args: node.args.map((a) => mapRefs(a, t) as Node) };
    default:
      return node;
  }
}

function printP(node: PNode): string {
  if (node.kind === "referr") return "#REF!";
  switch (node.kind) {
    case "num":
      return String(node.value);
    case "str":
      return '"' + node.value.replace(/"/g, '""') + '"';
    case "bool":
      return node.value ? "TRUE" : "FALSE";
    case "ref":
      return sheetPrefix(node.sheet) + formatA1(node.ref.row, node.ref.col, node.ref.absRow, node.ref.absCol);
    case "range":
      if (node.fullCol) return sheetPrefix(node.sheet) + colToLetter(node.start.col) + ":" + colToLetter(node.end.col);
      if (node.fullRow) return sheetPrefix(node.sheet) + (node.start.row + 1) + ":" + (node.end.row + 1);
      return (
        sheetPrefix(node.sheet) +
        formatA1(node.start.row, node.start.col, node.start.absRow, node.start.absCol) +
        ":" +
        formatA1(node.end.row, node.end.col, node.end.absRow, node.end.absCol)
      );
    case "name":
      return node.name;
    case "unary":
      return node.op === "%" ? printP(node.operand) + "%" : node.op + printP(node.operand);
    case "binary":
      return printP(node.left) + node.op + printP(node.right);
    case "call":
      return node.name + "(" + node.args.map(printP).join(",") + ")";
  }
}

// Offset every relative reference by (dRow, dCol) — used by the fill handle and
// copy/paste. Absolute ($) parts and the sheet qualifier are preserved.
export function offsetFormula(raw: string, dRow: number, dCol: number): string {
  let ast: Node;
  try {
    ast = parse(raw.slice(1));
  } catch {
    return raw;
  }
  const shift = (n: Node): Node => {
    switch (n.kind) {
      case "ref":
        return {
          kind: "ref",
          sheet: n.sheet,
          ref: {
            row: n.ref.absRow ? n.ref.row : n.ref.row + dRow,
            col: n.ref.absCol ? n.ref.col : n.ref.col + dCol,
            absRow: n.ref.absRow,
            absCol: n.ref.absCol,
          },
        };
      case "range":
        return {
          kind: "range",
          sheet: n.sheet,
          fullCol: n.fullCol,
          fullRow: n.fullRow,
          start: {
            row: n.start.absRow ? n.start.row : n.start.row + dRow,
            col: n.start.absCol ? n.start.col : n.start.col + dCol,
            absRow: n.start.absRow,
            absCol: n.start.absCol,
          },
          end: {
            row: n.end.absRow ? n.end.row : n.end.row + dRow,
            col: n.end.absCol ? n.end.col : n.end.col + dCol,
            absRow: n.end.absRow,
            absCol: n.end.absCol,
          },
        };
      case "unary":
        return { kind: "unary", op: n.op, operand: shift(n.operand) };
      case "binary":
        return { kind: "binary", op: n.op, left: shift(n.left), right: shift(n.right) };
      case "call":
        return { kind: "call", name: n.name, args: n.args.map(shift) };
      default:
        return n;
    }
  };
  return astToFormula(shift(ast));
}

// Rename a sheet inside formula references (used when a tab is renamed).
export function renameSheetRefs(raw: string, oldName: string, newName: string): string {
  let ast: Node;
  try {
    ast = parse(raw.slice(1));
  } catch {
    return raw;
  }
  const eq = (a?: string) => a !== undefined && a.toLowerCase() === oldName.toLowerCase();
  const walk = (n: Node): Node => {
    switch (n.kind) {
      case "ref":
        return eq(n.sheet) ? { ...n, sheet: newName } : n;
      case "range":
        return eq(n.sheet) ? { ...n, sheet: newName } : n;
      case "unary":
        return { kind: "unary", op: n.op, operand: walk(n.operand) };
      case "binary":
        return { kind: "binary", op: n.op, left: walk(n.left), right: walk(n.right) };
      case "call":
        return { kind: "call", name: n.name, args: n.args.map(walk) };
      default:
        return n;
    }
  };
  return astToFormula(walk(ast));
}
