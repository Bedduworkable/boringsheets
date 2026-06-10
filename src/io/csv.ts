// RFC 4180 CSV parsing and serialization. Pure functions, no I/O.

// Parse CSV text into a 2D array of string cells. Handles: fields quoted with
// double quotes; embedded commas, CR, LF, and CRLF inside quoted fields;
// doubled "" as an escaped quote; rows separated by \n, \r\n, or \r. A single
// trailing line terminator does not produce an extra empty row.
export function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let inQuotes = false;
  // Tracks whether the current row has had any content/structure pushed, so we
  // can distinguish a genuine trailing terminator from an empty final row.
  let started = false;

  const endField = () => {
    row.push(field);
    field = "";
  };
  const endRow = () => {
    row.push(field);
    field = "";
    rows.push(row);
    row = [];
    started = false;
  };

  const n = text.length;
  let i = 0;
  while (i < n) {
    const ch = text[i];

    if (inQuotes) {
      if (ch === '"') {
        if (i + 1 < n && text[i + 1] === '"') {
          // Escaped quote.
          field += '"';
          i += 2;
        } else {
          // Closing quote.
          inQuotes = false;
          i += 1;
        }
      } else {
        field += ch;
        i += 1;
      }
      continue;
    }

    if (ch === '"') {
      inQuotes = true;
      started = true;
      i += 1;
    } else if (ch === ",") {
      started = true;
      endField();
      i += 1;
    } else if (ch === "\n") {
      started = true;
      endRow();
      i += 1;
    } else if (ch === "\r") {
      started = true;
      endRow();
      // Consume a following \n as part of a CRLF terminator.
      if (i + 1 < n && text[i + 1] === "\n") i += 2;
      else i += 1;
    } else {
      started = true;
      field += ch;
      i += 1;
    }
  }

  // Flush the final field/row. Only emit a trailing row if it carried content
  // (a field separator, a quote, or any character) since the last terminator.
  if (started || field.length > 0) {
    endField();
    rows.push(row);
  }

  return rows;
}

// Serialize rows to CSV. A field is quoted iff it contains a comma, double
// quote, CR, or LF; embedded quotes are doubled. number -> String(n);
// boolean -> "TRUE"/"FALSE"; null/undefined -> "". Rows joined with "\r\n".
export function toCsv(rows: (string | number | boolean | null)[][]): string {
  const encodeField = (cell: string | number | boolean | null | undefined): string => {
    let s: string;
    if (cell === null || cell === undefined) s = "";
    else if (typeof cell === "number") s = String(cell);
    else if (typeof cell === "boolean") s = cell ? "TRUE" : "FALSE";
    else s = cell;

    if (s.includes(",") || s.includes('"') || s.includes("\r") || s.includes("\n")) {
      return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  };

  return rows.map((r) => r.map(encodeField).join(",")).join("\r\n");
}
