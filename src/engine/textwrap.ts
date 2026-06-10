// Word-wrap a string to a pixel width. `measure` returns the pixel width of a
// string in the current font. Breaks on whitespace, and hard-breaks tokens that
// are themselves wider than the available width (e.g. long reference numbers).

export function wrapLines(measure: (s: string) => number, text: string, maxWidth: number): string[] {
  if (maxWidth <= 0) return [text];
  const out: string[] = [];

  const hardBreak = (token: string, seed: string): string => {
    let chunk = seed;
    for (const ch of token) {
      if (chunk !== "" && measure(chunk + ch) > maxWidth) {
        out.push(chunk);
        chunk = ch;
      } else {
        chunk += ch;
      }
    }
    return chunk;
  };

  for (const rawLine of text.split("\n")) {
    let line = "";
    // split keeping the whitespace runs as their own tokens
    const tokens = rawLine.split(/(\s+)/);
    for (const token of tokens) {
      if (token === "") continue;
      const candidate = line + token;
      if (measure(candidate) <= maxWidth) {
        line = candidate;
      } else if (line === "") {
        // token alone overflows → hard-break it
        line = hardBreak(token, "");
      } else {
        out.push(line.replace(/\s+$/, ""));
        const t = token.replace(/^\s+/, "");
        line = measure(t) > maxWidth ? hardBreak(t, "") : t;
      }
    }
    out.push(line.replace(/\s+$/, ""));
  }
  return out.length ? out : [""];
}
