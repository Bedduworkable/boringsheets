// A1-notation helpers. Columns are 0-based internally (A=0, B=1, ... Z=25,
// AA=26, ...). Rows are 0-based internally but displayed 1-based.

export function colToLetter(col: number): string {
  let s = "";
  let n = col;
  while (n >= 0) {
    s = String.fromCharCode((n % 26) + 65) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

export function letterToCol(letters: string): number {
  let n = 0;
  for (let i = 0; i < letters.length; i++) {
    n = n * 26 + (letters.charCodeAt(i) - 64);
  }
  return n - 1;
}

export interface CellRef {
  row: number;
  col: number;
  absRow: boolean;
  absCol: boolean;
}

const A1_RE = /^(\$?)([A-Z]+)(\$?)(\d+)$/;

export function parseA1(a1: string): CellRef | null {
  const m = A1_RE.exec(a1.toUpperCase());
  if (!m) return null;
  const [, absC, letters, absR, digits] = m;
  return {
    col: letterToCol(letters),
    row: parseInt(digits, 10) - 1,
    absCol: absC === "$",
    absRow: absR === "$",
  };
}

export function formatA1(row: number, col: number, absRow = false, absCol = false): string {
  return `${absCol ? "$" : ""}${colToLetter(col)}${absRow ? "$" : ""}${row + 1}`;
}

// A reference possibly qualified by a sheet name, e.g. Sheet2!A1 or 'My Sheet'!A1.
export interface QualifiedRef {
  sheet?: string;
  ref: CellRef;
}

export function parseRef(text: string): QualifiedRef | null {
  let sheet: string | undefined;
  let cellPart = text;
  const bang = text.lastIndexOf("!");
  if (bang !== -1) {
    let s = text.slice(0, bang);
    cellPart = text.slice(bang + 1);
    if (s.startsWith("'") && s.endsWith("'")) s = s.slice(1, -1).replace(/''/g, "'");
    sheet = s;
  }
  const ref = parseA1(cellPart);
  if (!ref) return null;
  return { sheet, ref };
}

// Quote a sheet name for a formula if it needs it (spaces or punctuation).
export function formatSheetName(name: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) return name;
  return "'" + name.replace(/'/g, "''") + "'";
}

export function cellKey(row: number, col: number): string {
  return `${row},${col}`;
}

export function parseKey(key: string): { row: number; col: number } {
  const i = key.indexOf(",");
  return { row: parseInt(key.slice(0, i), 10), col: parseInt(key.slice(i + 1), 10) };
}
