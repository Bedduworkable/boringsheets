// Core value + formatting types shared across the engine and the UI.

export type ErrorKind =
  | "#DIV/0!"
  | "#VALUE!"
  | "#REF!"
  | "#NAME?"
  | "#NUM!"
  | "#N/A"
  | "#NULL!"
  | "#CIRCULAR!";

export class CellError {
  constructor(public readonly kind: ErrorKind, public readonly detail?: string) {}
  toString() {
    return this.kind;
  }
}

// A computed cell value. `null` means an empty cell.
export type CellValue = number | string | boolean | CellError | null;

export type HorizontalAlign = "left" | "center" | "right";

// Which edges of a cell have a border, plus an optional shared color.
export interface BorderSet {
  top?: boolean;
  bottom?: boolean;
  left?: boolean;
  right?: boolean;
  color?: string;
}

export interface CellFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string; // text color, hex
  bg?: string; // fill color, hex
  align?: HorizontalAlign;
  wrap?: boolean; // wrap long text onto multiple lines (auto-grows row height)
  fontSize?: number;
  fontFamily?: string;
  // Number format code, Excel-style (e.g. "0.00", "$#,##0.00", "0%", "yyyy-mm-dd").
  numFmt?: string;
  border?: BorderSet;
}

export interface Cell {
  // The raw text the user typed: "42", "hello", "=SUM(A1:A3)".
  raw: string;
  // Cached computed value (recomputed when dependencies change).
  value: CellValue;
  // Optional per-cell formatting.
  format?: CellFormat;
  // Optional cell note/comment.
  note?: string;
}
