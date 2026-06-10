// Formula lexer. Turns "=SUM(A1:A3)*2" (without the leading '=') into a flat
// token stream the parser consumes.

export type TokenType =
  | "number"
  | "string"
  | "ref" // a cell like A1 or $B$2
  | "ident" // a name: function name, TRUE/FALSE, or named range
  | "op"
  | "lparen"
  | "rparen"
  | "comma"
  | "colon"
  | "eof";

export interface Token {
  type: TokenType;
  value: string;
  pos: number;
}

const OPS = new Set(["+", "-", "*", "/", "^", "&", "=", "<", ">", "<=", ">=", "<>", "%"]);

export function tokenize(input: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  const n = input.length;

  const isDigit = (c: string) => c >= "0" && c <= "9";
  const isAlpha = (c: string) => (c >= "A" && c <= "Z") || (c >= "a" && c <= "z") || c === "_";

  // Read a plain cell ref ($?letters$?digits) starting at p; return its end
  // index, or -1 if there isn't one.
  const readCellRef = (p: number): number => {
    let q = p;
    if (input[q] === "$") q++;
    let hadAlpha = false;
    while (q < n && isAlpha(input[q])) { q++; hadAlpha = true; }
    if (!hadAlpha) return -1;
    if (input[q] === "$") q++;
    let hadDigit = false;
    while (q < n && isDigit(input[q])) { q++; hadDigit = true; }
    return hadDigit ? q : -1;
  };

  while (i < n) {
    const c = input[i];

    // whitespace
    if (c === " " || c === "\t" || c === "\n" || c === "\r") {
      i++;
      continue;
    }

    // quoted sheet name: 'My Sheet'!A1  ('' is an escaped quote)
    if (c === "'") {
      let q = i + 1;
      while (q < n) {
        if (input[q] === "'") {
          if (input[q + 1] === "'") { q += 2; continue; }
          q++;
          break;
        }
        q++;
      }
      if (input[q] === "!") {
        const end = readCellRef(q + 1);
        if (end !== -1) {
          tokens.push({ type: "ref", value: input.slice(i, end), pos: i });
          i = end;
          continue;
        }
      }
      throw new Error(`Expected a sheet reference after '${input.slice(i, q)}'`);
    }

    // string literal: "double quotes", with "" as an escaped quote
    if (c === '"') {
      let s = "";
      i++;
      while (i < n) {
        if (input[i] === '"') {
          if (input[i + 1] === '"') {
            s += '"';
            i += 2;
            continue;
          }
          i++;
          break;
        }
        s += input[i++];
      }
      tokens.push({ type: "string", value: s, pos: i });
      continue;
    }

    // number: 123, 1.5, .5, 1e3, 1.2E-4
    if (isDigit(c) || (c === "." && isDigit(input[i + 1]))) {
      let j = i;
      while (j < n && isDigit(input[j])) j++;
      if (input[j] === ".") {
        j++;
        while (j < n && isDigit(input[j])) j++;
      }
      if (input[j] === "e" || input[j] === "E") {
        let k = j + 1;
        if (input[k] === "+" || input[k] === "-") k++;
        if (isDigit(input[k])) {
          k++;
          while (k < n && isDigit(input[k])) k++;
          j = k;
        }
      }
      tokens.push({ type: "number", value: input.slice(i, j), pos: i });
      i = j;
      continue;
    }

    // identifier or cell reference. Letters (optionally with $) then maybe digits.
    if (isAlpha(c) || (c === "$" && (isAlpha(input[i + 1]) || isDigit(input[i + 1])))) {
      let j = i;
      while (j < n && (isAlpha(input[j]) || isDigit(input[j]) || input[j] === "$" || input[j] === ".")) {
        j++;
      }
      const word = input.slice(i, j);
      // Sheet-qualified reference: an identifier followed by "!" and a cell ref.
      if (input[j] === "!") {
        const end = readCellRef(j + 1);
        if (end !== -1) {
          tokens.push({ type: "ref", value: input.slice(i, end), pos: i });
          i = end;
          continue;
        }
      }
      // A word like "A1" or "$B$2" is a cell ref — UNLESS it's immediately
      // followed by "(", in which case it's a function whose name ends in a
      // digit (ATAN2, LOG10, etc.).
      if (/^\$?[A-Za-z]+\$?\d+$/.test(word) && input[j] !== "(") {
        tokens.push({ type: "ref", value: word, pos: i });
      } else {
        tokens.push({ type: "ident", value: word, pos: i });
      }
      i = j;
      continue;
    }

    // two-char operators first
    const two = input.slice(i, i + 2);
    if (two === "<=" || two === ">=" || two === "<>") {
      tokens.push({ type: "op", value: two, pos: i });
      i += 2;
      continue;
    }

    if (c === "(") {
      tokens.push({ type: "lparen", value: c, pos: i++ });
      continue;
    }
    if (c === ")") {
      tokens.push({ type: "rparen", value: c, pos: i++ });
      continue;
    }
    if (c === ",") {
      tokens.push({ type: "comma", value: c, pos: i++ });
      continue;
    }
    if (c === ":") {
      tokens.push({ type: "colon", value: c, pos: i++ });
      continue;
    }
    if (OPS.has(c)) {
      tokens.push({ type: "op", value: c, pos: i++ });
      continue;
    }

    throw new Error(`Unexpected character '${c}' at ${i}`);
  }

  tokens.push({ type: "eof", value: "", pos: n });
  return tokens;
}
