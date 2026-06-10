// Formula function autocomplete. Attaches to a text <input> that holds a cell
// formula; when the user types a function name after "=", "(", "," or an
// operator, a dropdown of matching functions appears (Excel-style).

import { FUNCTION_NAMES } from "../engine/functions.js";

// Short syntax hints for the most common functions (others just show the name).
const SIGNATURES: Record<string, string> = {
  SUM: "(number1, …)", AVERAGE: "(number1, …)", COUNT: "(value1, …)", COUNTA: "(value1, …)",
  COUNTIF: "(range, criteria)", COUNTIFS: "(range1, crit1, …)", SUMIF: "(range, criteria, [sum_range])",
  SUMIFS: "(sum_range, range1, crit1, …)", AVERAGEIF: "(range, criteria, [avg_range])",
  MAX: "(number1, …)", MIN: "(number1, …)", PRODUCT: "(number1, …)", ROUND: "(number, digits)",
  ROUNDUP: "(number, digits)", ROUNDDOWN: "(number, digits)", INT: "(number)", ABS: "(number)",
  MOD: "(number, divisor)", POWER: "(number, power)", SQRT: "(number)", SUBTOTAL: "(func_num, ref1, …)",
  AGGREGATE: "(func_num, options, ref1, …)", SUMPRODUCT: "(array1, array2, …)",
  IF: "(logical_test, value_if_true, [value_if_false])", IFS: "(test1, val1, …)",
  IFERROR: "(value, value_if_error)", IFNA: "(value, value_if_na)", AND: "(logical1, …)",
  OR: "(logical1, …)", NOT: "(logical)", SWITCH: "(expr, val1, res1, …, [default])",
  VLOOKUP: "(lookup, table, col_index, [exact])", HLOOKUP: "(lookup, table, row_index, [exact])",
  XLOOKUP: "(lookup, lookup_array, return_array, [if_not_found])", INDEX: "(array, row_num, [col_num])",
  MATCH: "(lookup, array, [match_type])", CHOOSE: "(index, val1, …)", LOOKUP: "(lookup, vector, [result])",
  OFFSET: "(ref, rows, cols, [height], [width])", INDIRECT: "(ref_text)", ADDRESS: "(row, col, [abs])",
  ROW: "([ref])", COLUMN: "([ref])", ROWS: "(array)", COLUMNS: "(array)",
  CONCAT: "(text1, …)", CONCATENATE: "(text1, …)", TEXTJOIN: "(delim, ignore_empty, text1, …)",
  LEFT: "(text, [num])", RIGHT: "(text, [num])", MID: "(text, start, num)", LEN: "(text)",
  TRIM: "(text)", UPPER: "(text)", LOWER: "(text)", PROPER: "(text)", REPLACE: "(text, start, len, new)",
  SUBSTITUTE: "(text, old, new, [which])", FIND: "(find, within, [start])", SEARCH: "(find, within, [start])",
  TEXT: "(value, format)", VALUE: "(text)",
  TODAY: "()", NOW: "()", DATE: "(year, month, day)", YEAR: "(date)", MONTH: "(date)", DAY: "(date)",
  EDATE: "(date, months)", EOMONTH: "(date, months)", DATEDIF: "(start, end, unit)", DAYS: "(end, start)",
  NETWORKDAYS: "(start, end)", WEEKDAY: "(date)", WEEKNUM: "(date)",
  MEDIAN: "(number1, …)", STDEV: "(number1, …)", VAR: "(number1, …)", RANK: "(number, ref, [order])",
  PERCENTILE: "(array, k)", LARGE: "(array, k)", SMALL: "(array, k)",
  ISBLANK: "(value)", ISNUMBER: "(value)", ISTEXT: "(value)", ISERROR: "(value)",
  PMT: "(rate, nper, pv, [fv], [type])", FV: "(rate, nper, pmt, [pv], [type])", PV: "(rate, nper, pmt, [fv], [type])",
  NPV: "(rate, value1, …)",
};

export class FormulaAutocomplete {
  private popup: HTMLElement | null = null;
  private items: string[] = [];
  private sel = 0;
  private tokenStart = 0;

  constructor(private input: HTMLInputElement) {}

  isOpen() {
    return this.popup !== null;
  }

  // Recompute suggestions from the input's current value + cursor.
  onInput() {
    const value = this.input.value;
    const cursor = this.input.selectionStart ?? value.length;
    const tok = currentToken(value, cursor);
    if (!tok) return this.close();
    const upper = tok.token.toUpperCase();
    const matches = FUNCTION_NAMES.filter((n) => n.startsWith(upper));
    if (matches.length === 0 || (matches.length === 1 && matches[0] === upper)) return this.close();
    this.items = matches.slice(0, 12);
    this.tokenStart = tok.start;
    this.sel = 0;
    this.render();
  }

  // Returns true if the key was consumed by the dropdown.
  handleKeydown(e: KeyboardEvent): boolean {
    if (!this.popup) return false;
    if (e.key === "ArrowDown") {
      e.preventDefault();
      this.sel = (this.sel + 1) % this.items.length;
      this.render();
      return true;
    }
    if (e.key === "ArrowUp") {
      e.preventDefault();
      this.sel = (this.sel - 1 + this.items.length) % this.items.length;
      this.render();
      return true;
    }
    if (e.key === "Enter" || e.key === "Tab") {
      e.preventDefault();
      this.accept(this.items[this.sel]);
      return true;
    }
    if (e.key === "Escape") {
      e.preventDefault();
      this.close();
      return true;
    }
    return false;
  }

  private accept(name: string) {
    const value = this.input.value;
    const cursor = this.input.selectionStart ?? value.length;
    const next = value.slice(0, this.tokenStart) + name + "(" + value.slice(cursor);
    this.input.value = next;
    const caret = this.tokenStart + name.length + 1;
    this.input.setSelectionRange(caret, caret);
    this.close();
    // notify any 'input' listeners (e.g. formula bar sync)
    this.input.dispatchEvent(new Event("input", { bubbles: true }));
  }

  private render() {
    if (!this.popup) {
      this.popup = document.createElement("div");
      this.popup.className = "fn-autocomplete";
      document.body.appendChild(this.popup);
    }
    this.popup.innerHTML = "";
    this.items.forEach((name, i) => {
      const row = document.createElement("div");
      row.className = "fn-item" + (i === this.sel ? " active" : "");
      row.innerHTML = `<span class="fn-name">${name}</span><span class="fn-sig">${SIGNATURES[name] ?? "()"}</span>`;
      row.addEventListener("mousedown", (ev) => {
        ev.preventDefault();
        this.accept(name);
      });
      this.popup!.appendChild(row);
    });
    const rect = this.input.getBoundingClientRect();
    const top = rect.bottom + 1;
    this.popup.style.left = `${Math.min(rect.left, window.innerWidth - 280)}px`;
    this.popup.style.top = `${top}px`;
  }

  close() {
    if (this.popup) {
      this.popup.remove();
      this.popup = null;
    }
  }
}

// Extract the function-name token being typed at the cursor, if any.
function currentToken(value: string, cursor: number): { token: string; start: number } | null {
  if (!value.startsWith("=")) return null;
  let start = cursor;
  while (start > 0 && /[A-Za-z0-9.]/.test(value[start - 1])) start--;
  const token = value.slice(start, cursor);
  if (token === "" || !/^[A-Za-z]/.test(token)) return null;
  const prev = start > 0 ? value[start - 1] : "=";
  // a function name follows "=", "(", ",", an operator, or whitespace
  if (!/[=(,+\-*/^&<>%: ]/.test(prev)) return null;
  return { token, start };
}
