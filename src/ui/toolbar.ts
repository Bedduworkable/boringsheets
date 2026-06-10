// Google-Sheets-style toolbar: Material-style SVG icon buttons, grouped with
// dividers. Functionally drives the same App methods; the look is the upgrade.

import type { App } from "../app.js";
import { icon, IconName } from "./icons.js";

const NUM_FORMATS: { label: string; code: string }[] = [
  { label: "Automatic", code: "General" },
  { label: "Number", code: "0.00" },
  { label: "Currency", code: '"$"#,##0.00' },
  { label: "Accounting", code: "#,##0.00" },
  { label: "Percent", code: "0%" },
  { label: "Comma", code: "#,##0" },
  { label: "Date", code: "yyyy-mm-dd" },
  { label: "Time", code: "h:mm:ss" },
  { label: "Plain text", code: "@" },
];

export function buildToolbar(el: HTMLElement, app: App) {
  el.innerHTML = "";

  const mk = (name: IconName, title: string, onClick: () => void) => {
    const b = document.createElement("button");
    b.className = "tb-btn";
    b.innerHTML = icon(name);
    b.title = title;
    b.tabIndex = -1;
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", onClick);
    el.appendChild(b);
    return b;
  };
  const sep = () => {
    const s = document.createElement("div");
    s.className = "tb-sep";
    el.appendChild(s);
  };

  // A color button: icon + colored underline bar, opens the native color picker.
  const mkColor = (name: IconName, title: string, initial: string, onPick: (c: string) => void) => {
    const wrap = document.createElement("label");
    wrap.className = "tb-btn tb-colorbtn";
    wrap.title = title;
    wrap.innerHTML = icon(name);
    wrap.addEventListener("mousedown", (e) => e.preventDefault());
    const bar = document.createElement("span");
    bar.className = "tb-colorbar";
    bar.style.background = initial;
    wrap.appendChild(bar);
    const input = document.createElement("input");
    input.type = "color";
    input.value = initial;
    input.className = "tb-colorinput";
    input.tabIndex = -1;
    input.addEventListener("input", () => {
      bar.style.background = input.value;
      onPick(input.value);
    });
    wrap.appendChild(input);
    el.appendChild(wrap);
    return { wrap, input, bar };
  };

  const sel = app.grid.selRange.bind(app.grid);

  // 1. History + paint format
  mk("undo", "Undo (⌘Z)", () => app.undo());
  mk("redo", "Redo (⌘⇧Z)", () => app.redo());
  mk("paint", "Paint format (then select target)", () => app.startFormatPainter());
  sep();

  // 2. Zoom
  const zoomOut = mk("caret", "Zoom out", () => app.zoomOut());
  zoomOut.innerHTML = "&minus;";
  zoomOut.classList.add("tb-text");
  const zoomLabel = document.createElement("button");
  zoomLabel.className = "tb-btn tb-zoom";
  zoomLabel.textContent = "100%";
  zoomLabel.title = "Reset zoom";
  zoomLabel.tabIndex = -1;
  zoomLabel.addEventListener("mousedown", (e) => e.preventDefault());
  zoomLabel.addEventListener("click", () => app.zoomReset());
  el.appendChild(zoomLabel);
  const zoomIn = mk("caret", "Zoom in", () => app.zoomIn());
  zoomIn.innerHTML = "+";
  zoomIn.classList.add("tb-text");
  sep();

  // 3. Number formats
  mk("currency", "Format as currency", () => app.applyFormat({ numFmt: '"$"#,##0.00' }));
  mk("percent", "Format as percent", () => app.applyFormat({ numFmt: "0%" }));
  mk("decDec", "Decrease decimal places", () => app.changeDecimals(-1));
  mk("incDec", "Increase decimal places", () => app.changeDecimals(1));
  const numSel = document.createElement("select");
  numSel.className = "tb-select";
  numSel.title = "More formats";
  for (const f of NUM_FORMATS) {
    const o = document.createElement("option");
    o.value = f.code;
    o.textContent = f.label;
    numSel.appendChild(o);
  }
  numSel.addEventListener("change", () => app.applyFormat({ numFmt: numSel.value === "General" ? undefined : numSel.value }));
  el.appendChild(numSel);
  sep();

  // 3b. Font family + size
  const FONTS: [string, string][] = [
    ["", "Default"], ["Arial", "Arial"], ["Georgia", "Georgia"], ["Times New Roman", "Times New Roman"],
    ["Courier New", "Courier New"], ["Verdana", "Verdana"], ["Trebuchet MS", "Trebuchet MS"], ["Comic Sans MS", "Comic Sans MS"],
  ];
  const fontSel = document.createElement("select");
  fontSel.className = "tb-select tb-font";
  fontSel.title = "Font";
  for (const [val, label] of FONTS) {
    const o = document.createElement("option");
    o.value = val;
    o.textContent = label;
    if (val) o.style.fontFamily = val;
    fontSel.appendChild(o);
  }
  fontSel.addEventListener("change", () => app.setFontFamily(fontSel.value));
  el.appendChild(fontSel);

  // Font size: − [10] + with a custom suggestions dropdown (Google Sheets style)
  const SIZES = [6, 7, 8, 9, 10, 11, 12, 14, 18, 24, 36];
  const sizeBox = document.createElement("div");
  sizeBox.className = "tb-sizebox";
  const minus = document.createElement("button");
  minus.className = "tb-step";
  minus.innerHTML = "&minus;";
  minus.tabIndex = -1;
  minus.title = "Decrease font size";
  const sizeInput = document.createElement("input");
  sizeInput.type = "text";
  sizeInput.inputMode = "numeric";
  sizeInput.className = "tb-size";
  sizeInput.title = "Font size";
  sizeInput.value = "10";
  const plus = document.createElement("button");
  plus.className = "tb-step";
  plus.textContent = "+";
  plus.tabIndex = -1;
  plus.title = "Increase font size";
  sizeBox.append(minus, sizeInput, plus);
  el.appendChild(sizeBox);

  const curSize = () => {
    const n = parseInt(sizeInput.value, 10);
    return Number.isNaN(n) ? app.activeFontPt() : n;
  };
  const setSize = (n: number) => {
    if (n < 1) return;
    sizeInput.value = String(n);
    app.setFontSize(n);
  };

  let sizeDrop: HTMLElement | null = null;
  const closeDrop = () => {
    if (sizeDrop) {
      sizeDrop.remove();
      sizeDrop = null;
    }
  };
  const openDrop = () => {
    closeDrop();
    const d = document.createElement("div");
    d.className = "size-dropdown";
    const cur = curSize();
    for (const s of SIZES) {
      const item = document.createElement("div");
      item.className = "size-item" + (s === cur ? " active" : "");
      item.textContent = String(s);
      item.addEventListener("mousedown", (ev) => {
        ev.preventDefault();
        setSize(s);
        closeDrop();
        sizeInput.blur();
      });
      d.appendChild(item);
    }
    document.body.appendChild(d);
    const r = sizeInput.getBoundingClientRect();
    d.style.left = `${r.left}px`;
    d.style.top = `${r.bottom + 3}px`;
    d.style.minWidth = `${r.width}px`;
    sizeDrop = d;
  };

  // Step from the SELECTED cell's actual size (not the possibly-stale input),
  // so a click instantly resizes the current selection.
  minus.addEventListener("mousedown", (e) => e.preventDefault());
  minus.addEventListener("click", () => setSize(app.activeFontPt() - 1));
  plus.addEventListener("mousedown", (e) => e.preventDefault());
  plus.addEventListener("click", () => setSize(app.activeFontPt() + 1));
  sizeInput.addEventListener("focus", () => {
    sizeInput.select();
    openDrop();
  });
  sizeInput.addEventListener("blur", () => setTimeout(closeDrop, 100));
  sizeInput.addEventListener("change", () => {
    const n = parseInt(sizeInput.value, 10);
    if (n >= 1) app.setFontSize(n);
  });
  sizeInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      const n = parseInt(sizeInput.value, 10);
      if (n >= 1) app.setFontSize(n);
      sizeInput.blur();
    } else if (e.key === "Escape") {
      closeDrop();
      sizeInput.blur();
    }
  });
  sep();

  // 4. Text style
  const bBtn = mk("bold", "Bold (⌘B)", () => app.toggleFormat("bold"));
  const iBtn = mk("italic", "Italic (⌘I)", () => app.toggleFormat("italic"));
  const uBtn = mk("underline", "Underline (⌘U)", () => app.toggleFormat("underline"));
  const sBtn = mk("strike", "Strikethrough", () => app.toggleFormat("strike"));
  const textColor = mkColor("textColor", "Text color", "#1b1b1b", (c) => app.applyFormat({ color: c }));
  mkColor("fillColor", "Fill color", "#ffff00", (c) => app.applyFormat({ bg: c }));
  sep();

  // 5. Borders + merge
  const borderSel = document.createElement("select");
  borderSel.className = "tb-select tb-iconselect";
  borderSel.title = "Borders";
  for (const [val, label] of [
    ["", "Borders"], ["all", "All borders"], ["outer", "Outer"], ["top", "Top"],
    ["bottom", "Bottom"], ["left", "Left"], ["right", "Right"], ["none", "Clear"],
  ] as [string, string][]) {
    const o = document.createElement("option");
    o.value = val;
    o.textContent = label;
    borderSel.appendChild(o);
  }
  borderSel.addEventListener("change", () => {
    if (borderSel.value) app.applyBorder(borderSel.value as Parameters<typeof app.applyBorder>[0]);
    borderSel.value = "";
  });
  el.appendChild(borderSel);
  mk("merge", "Merge cells", () => {
    const s = sel();
    if (app.wb.active.mergeAt(s.r0, s.c0)) app.unmergeCells();
    else app.mergeCells();
  });
  sep();

  // 6. Alignment + wrap
  const alignL = mk("alignLeft", "Align left", () => app.applyFormat({ align: "left" }));
  const alignC = mk("alignCenter", "Align center", () => app.applyFormat({ align: "center" }));
  const alignR = mk("alignRight", "Align right", () => app.applyFormat({ align: "right" }));
  const wrapBtn = mk("wrap", "Wrap text", () => app.toggleWrap());
  sep();

  // 7. Structure + data
  mk("freeze", "Freeze panes at selection", () => app.toggleFreeze());
  mk("autofit", "Auto-fit all columns", () => app.autofitAllColumns());
  mk("insRow", "Insert row above", () => { const s = sel(); app.insertRows(s.r0, 1); });
  mk("delRow", "Delete row(s)", () => { const s = sel(); app.deleteRows(s.r0, s.r1 - s.r0 + 1); });
  mk("sortAsc", "Sort ascending", () => app.sortRange(true));
  mk("sortAsc", "Sort descending", () => app.sortRange(false)).classList.add("tb-flip");
  mk("filter", "Create / remove filter", () => app.toggleFilter());
  sep();

  // 8. Insert / tools
  mk("func", "AutoSum (Alt+=)", () => app.autoSum());
  mk("chart", "Insert chart", () => app.openChartDialog());
  mk("palette", "Conditional formatting", () => app.openConditionalDialog());
  mk("validation", "Data validation", () => app.openValidationDialog());
  mk("note", "Insert note", () => app.editNote());
  mk("search", "Find & replace (⌘F)", () => app.openFindReplace());

  // Reflect the active cell's formatting on every selection change.
  el.addEventListener("refresh", () => {
    zoomLabel.textContent = `${app.getZoomPct()}%`;
    const fmt = app.activeFormat();
    bBtn.classList.toggle("active", !!fmt?.bold);
    iBtn.classList.toggle("active", !!fmt?.italic);
    uBtn.classList.toggle("active", !!fmt?.underline);
    sBtn.classList.toggle("active", !!fmt?.strike);
    alignL.classList.toggle("active", fmt?.align === "left");
    alignC.classList.toggle("active", fmt?.align === "center");
    alignR.classList.toggle("active", fmt?.align === "right");
    wrapBtn.classList.toggle("active", !!fmt?.wrap);
    if (fmt?.color) textColor.input.value = fmt.color;
    textColor.bar.style.background = fmt?.color || "#1b1b1b";
    const match = NUM_FORMATS.find((f) => f.code === (fmt?.numFmt || "General"));
    numSel.value = match ? match.code : "General";
    const fam = app.activeFontFamily();
    fontSel.value = FONTS.some(([v]) => v === fam) ? fam : "";
    if (document.activeElement !== sizeInput) sizeInput.value = String(app.activeFontPt());
  });
}
