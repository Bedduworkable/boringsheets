// Conditional-formatting dialog: pick a rule type + parameters and apply it to
// the current selection.

import type { App } from "../app.js";
import type { ConditionalRule, CondType } from "../engine/conditional.js";

let dialog: HTMLElement | null = null;
let seq = 0;

export function showConditionalDialog(app: App) {
  close();
  const el = document.createElement("div");
  el.className = "cf-dialog";
  el.innerHTML = `
    <div class="fr-row"><span class="fr-title">Conditional formatting</span><span class="fr-close">&#x2715;</span></div>
    <div class="fr-row"><label>Rule</label>
      <select class="cf-type">
        <option value="greaterThan">Cell value &gt;</option>
        <option value="lessThan">Cell value &lt;</option>
        <option value="between">Between</option>
        <option value="equalTo">Equal to</option>
        <option value="textContains">Text contains</option>
        <option value="duplicate">Duplicate values</option>
        <option value="top">Top N</option>
        <option value="bottom">Bottom N</option>
        <option value="colorScale">Color scale</option>
        <option value="dataBar">Data bar</option>
      </select>
    </div>
    <div class="fr-row cf-p1"><label class="cf-l1">Value</label><input class="cf-v1" type="text" /></div>
    <div class="fr-row cf-p2" style="display:none"><label>and</label><input class="cf-v2" type="text" /></div>
    <div class="fr-row cf-fmt"><label>Format</label>
      <span class="cf-swatch">Fill</span><input class="cf-bg" type="color" value="#ffeb9c" />
      <span class="cf-swatch">Text</span><input class="cf-color" type="color" value="#9c5700" />
      <label class="cf-bold"><input type="checkbox" class="cf-b" /> Bold</label>
    </div>
    <div class="fr-row fr-buttons"><button class="cf-clear">Clear in selection</button><button class="cf-apply">Apply</button></div>
  `;
  document.body.appendChild(el);
  dialog = el;

  const type = el.querySelector(".cf-type") as HTMLSelectElement;
  const v1 = el.querySelector(".cf-v1") as HTMLInputElement;
  const v2 = el.querySelector(".cf-v2") as HTMLInputElement;
  const p1 = el.querySelector(".cf-p1") as HTMLElement;
  const p2 = el.querySelector(".cf-p2") as HTMLElement;
  const l1 = el.querySelector(".cf-l1") as HTMLElement;
  const fmtRow = el.querySelector(".cf-fmt") as HTMLElement;
  const bg = el.querySelector(".cf-bg") as HTMLInputElement;
  const color = el.querySelector(".cf-color") as HTMLInputElement;
  const bold = el.querySelector(".cf-b") as HTMLInputElement;

  const sync = () => {
    const t = type.value as CondType;
    const needsV2 = t === "between";
    const needsV1 = ["greaterThan", "lessThan", "between", "equalTo", "textContains", "top", "bottom"].includes(t);
    const isVisual = t === "colorScale" || t === "dataBar";
    p1.style.display = needsV1 ? "" : "none";
    p2.style.display = needsV2 ? "" : "none";
    fmtRow.style.display = isVisual ? "none" : "";
    l1.textContent = t === "textContains" ? "Text" : t === "top" || t === "bottom" ? "N" : "Value";
  };
  type.addEventListener("change", sync);
  sync();

  (el.querySelector(".fr-close") as HTMLElement).addEventListener("click", close);
  (el.querySelector(".cf-clear") as HTMLButtonElement).addEventListener("click", () => {
    app.clearConditional();
    close();
  });
  (el.querySelector(".cf-apply") as HTMLButtonElement).addEventListener("click", () => {
    const t = type.value as CondType;
    const rect = app.selectionRect();
    const rule: ConditionalRule = { id: `cf-${++seq}`, range: { ...rect }, type: t };
    const format = { bold: bold.checked || undefined, bg: bg.value, color: color.value };
    if (t === "between") {
      rule.value1 = parseFloat(v1.value);
      rule.value2 = parseFloat(v2.value);
      rule.format = format;
    } else if (t === "greaterThan" || t === "lessThan") {
      rule.value1 = parseFloat(v1.value);
      rule.format = format;
    } else if (t === "equalTo" || t === "textContains") {
      rule.value1 = v1.value;
      rule.format = format;
    } else if (t === "top" || t === "bottom") {
      rule.n = parseInt(v1.value, 10) || 10;
      rule.format = format;
    } else if (t === "duplicate") {
      rule.format = format;
    } else if (t === "colorScale") {
      rule.minColor = "#f8696b";
      rule.midColor = "#ffeb84";
      rule.maxColor = "#63be7b";
    } else if (t === "dataBar") {
      rule.color = "#638ec6";
    }
    app.addConditionalRule(rule);
    close();
  });

  el.addEventListener("keydown", (e) => {
    if (e.key === "Escape") close();
  });
  v1.focus();
}

function close() {
  if (dialog) {
    dialog.remove();
    dialog = null;
  }
}
