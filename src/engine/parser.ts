// Recursive-descent parser producing an AST. Implements Excel operator
// precedence: comparison < concat < add/sub < mul/div < exponent < unary minus
// < percent(postfix) < range(:).

import { Token, tokenize } from "./tokenizer.js";
import { Node } from "./ast.js";
import { parseRef, letterToCol } from "./references.js";

export class ParseError extends Error {}

class Parser {
  private toks: Token[];
  private i = 0;

  constructor(input: string) {
    this.toks = tokenize(input);
  }

  private peek(): Token {
    return this.toks[this.i];
  }
  private next(): Token {
    return this.toks[this.i++];
  }
  private expect(type: Token["type"]): Token {
    const t = this.peek();
    if (t.type !== type) throw new ParseError(`Expected ${type} but got '${t.value || t.type}'`);
    return this.next();
  }

  parse(): Node {
    const node = this.parseComparison();
    if (this.peek().type !== "eof") {
      throw new ParseError(`Unexpected '${this.peek().value}'`);
    }
    return node;
  }

  // = <> < > <= >=
  private parseComparison(): Node {
    let left = this.parseConcat();
    while (this.peek().type === "op" && ["=", "<>", "<", ">", "<=", ">="].includes(this.peek().value)) {
      const op = this.next().value;
      const right = this.parseConcat();
      left = { kind: "binary", op, left, right };
    }
    return left;
  }

  // &
  private parseConcat(): Node {
    let left = this.parseAddSub();
    while (this.peek().type === "op" && this.peek().value === "&") {
      this.next();
      const right = this.parseAddSub();
      left = { kind: "binary", op: "&", left, right };
    }
    return left;
  }

  // + -
  private parseAddSub(): Node {
    let left = this.parseMulDiv();
    while (this.peek().type === "op" && (this.peek().value === "+" || this.peek().value === "-")) {
      const op = this.next().value;
      const right = this.parseMulDiv();
      left = { kind: "binary", op, left, right };
    }
    return left;
  }

  // * /
  private parseMulDiv(): Node {
    let left = this.parseExponent();
    while (this.peek().type === "op" && (this.peek().value === "*" || this.peek().value === "/")) {
      const op = this.next().value;
      const right = this.parseExponent();
      left = { kind: "binary", op, left, right };
    }
    return left;
  }

  // ^  (right-associative)
  private parseExponent(): Node {
    const left = this.parseUnary();
    if (this.peek().type === "op" && this.peek().value === "^") {
      this.next();
      const right = this.parseExponent();
      return { kind: "binary", op: "^", left, right };
    }
    return left;
  }

  // unary + / -
  private parseUnary(): Node {
    if (this.peek().type === "op" && (this.peek().value === "-" || this.peek().value === "+")) {
      const op = this.next().value as "-" | "+";
      const operand = this.parseUnary();
      return { kind: "unary", op, operand };
    }
    return this.parsePercent();
  }

  // postfix %
  private parsePercent(): Node {
    let node = this.parsePrimary();
    while (this.peek().type === "op" && this.peek().value === "%") {
      this.next();
      node = { kind: "unary", op: "%", operand: node };
    }
    return node;
  }

  private parsePrimary(): Node {
    const t = this.peek();

    if (t.type === "number") {
      // whole-row range: 5:8
      if (/^\d+$/.test(t.value) && this.toks[this.i + 1]?.type === "colon" && /^\d+$/.test(this.toks[this.i + 2]?.value || "")) {
        const r0 = parseInt(t.value, 10) - 1;
        this.next();
        this.next();
        const r1 = parseInt(this.next().value, 10) - 1;
        const lo = Math.min(r0, r1);
        const hi = Math.max(r0, r1);
        return {
          kind: "range",
          fullRow: true,
          start: { row: lo, col: 0, absRow: false, absCol: false },
          end: { row: hi, col: 0, absRow: false, absCol: false },
        };
      }
      this.next();
      return { kind: "num", value: parseFloat(t.value) };
    }
    if (t.type === "string") {
      this.next();
      return { kind: "str", value: t.value };
    }
    if (t.type === "lparen") {
      this.next();
      const inner = this.parseComparison();
      this.expect("rparen");
      return inner;
    }
    if (t.type === "ref") {
      return this.parseRefOrRange();
    }
    if (t.type === "ident") {
      const name = t.value.toUpperCase();
      // whole-column range: A:A or B:D
      if (/^[A-Za-z]+$/.test(t.value) && this.toks[this.i + 1]?.type === "colon" && /^[A-Za-z]+$/.test(this.toks[this.i + 2]?.value || "")) {
        const c0 = letterToCol(t.value.toUpperCase());
        this.next();
        this.next();
        const c1 = letterToCol(this.next().value.toUpperCase());
        const lo = Math.min(c0, c1);
        const hi = Math.max(c0, c1);
        return {
          kind: "range",
          fullCol: true,
          start: { row: 0, col: lo, absRow: false, absCol: false },
          end: { row: 0, col: hi, absRow: false, absCol: false },
        };
      }
      // boolean literals
      if (name === "TRUE" && this.toks[this.i + 1]?.type !== "lparen") {
        this.next();
        return { kind: "bool", value: true };
      }
      if (name === "FALSE" && this.toks[this.i + 1]?.type !== "lparen") {
        this.next();
        return { kind: "bool", value: false };
      }
      // A bare identifier not followed by "(" is a named-range reference.
      if (this.toks[this.i + 1]?.type !== "lparen") {
        this.next();
        return { kind: "name", name: t.value };
      }
      // function call
      this.next();
      this.expect("lparen");
      const args: Node[] = [];
      if (this.peek().type !== "rparen") {
        args.push(this.parseComparison());
        while (this.peek().type === "comma") {
          this.next();
          args.push(this.parseComparison());
        }
      }
      this.expect("rparen");
      return { kind: "call", name, args };
    }

    throw new ParseError(`Unexpected '${t.value || t.type}'`);
  }

  private parseRefOrRange(): Node {
    const startTok = this.expect("ref");
    const start = parseRef(startTok.value);
    if (!start) throw new ParseError(`Bad reference '${startTok.value}'`);
    if (this.peek().type === "colon") {
      this.next();
      const endTok = this.expect("ref");
      const end = parseRef(endTok.value);
      if (!end) throw new ParseError(`Bad reference '${endTok.value}'`);
      // A sheet qualifier on either endpoint applies to the whole range.
      return { kind: "range", start: start.ref, end: end.ref, sheet: start.sheet ?? end.sheet };
    }
    return { kind: "ref", ref: start.ref, sheet: start.sheet };
  }
}

export function parseFormula(input: string): Node {
  return new Parser(input).parse();
}
