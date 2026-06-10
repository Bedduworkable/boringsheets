import { CellRef } from "./references.js";

export type Node =
  | { kind: "num"; value: number }
  | { kind: "str"; value: string }
  | { kind: "bool"; value: boolean }
  | { kind: "ref"; ref: CellRef; sheet?: string }
  | { kind: "range"; start: CellRef; end: CellRef; sheet?: string; fullCol?: boolean; fullRow?: boolean }
  | { kind: "name"; name: string }
  | { kind: "unary"; op: "-" | "+" | "%"; operand: Node }
  | { kind: "binary"; op: string; left: Node; right: Node }
  | { kind: "call"; name: string; args: Node[] };
