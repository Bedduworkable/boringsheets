import { Sheet } from "./sheet.js";

// A workbook-level named range: `name` resolves to a rectangle on a sheet.
export interface NamedRange {
  name: string;
  sheetId: number;
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

// A workbook is an ordered set of sheets plus a pointer to the active one.
export class Workbook {
  sheets: Sheet[] = [];
  names: NamedRange[] = [];
  activeIndex = 0;
  filePath: string | null = null;
  dirty = false;

  constructor() {
    this.sheets.push(new Sheet("Sheet1"));
  }

  get active(): Sheet {
    return this.sheets[this.activeIndex];
  }

  addSheet(name?: string): Sheet {
    const n = name ?? this.uniqueName();
    const s = new Sheet(n);
    this.sheets.push(s);
    this.activeIndex = this.sheets.length - 1;
    return s;
  }

  private uniqueName(): string {
    let i = this.sheets.length + 1;
    const names = new Set(this.sheets.map((s) => s.name));
    while (names.has(`Sheet${i}`)) i++;
    return `Sheet${i}`;
  }

  setActive(i: number) {
    if (i >= 0 && i < this.sheets.length) this.activeIndex = i;
  }
}
