// Data-validation dialog: apply a list (dropdown) or numeric rule to the
// current selection.

import type { App } from "../app.js";
import type { DataValidation } from "../model/sheet.js";

let dialog: HTMLElement | null = null;

export function showValidationDialog(app: App) {
  close();
  const el = document.createElement("div");
  el.className = "cf-dialog";
  el.innerHTML = `
    <div class="fr-row"><span class="fr-title">Data validation</span><span class="fr-close">&#x2715;</span></div>
    <div class="fr-row"><label>Allow</label>
      <select class="dv-type">
        <option value="list">List (dropdown)</option>
        <option value="number">Number</option>
        <option value="textLength">Text length</option>
      </select>
    </div>
    <div class="fr-row dv-list"><label>Values</label><input class="dv-source" type="text" placeholder="Low, Medium, High" /></div>
    <div class="fr-row dv-num" style="display:none"><label>Operator</label>
      <select class="dv-op">
        <option value="between">between</option>
        <option value="notBetween">not between</option>
        <option value="gt">greater than</option>
        <option value="lt">less than</option>
        <option value="gte">≥</option>
        <option value="lte">≤</option>
        <option value="eq">equal to</option>
        <option value="ne">not equal to</option>
      </select>
    </div>
    <div class="fr-row dv-num" style="display:none"><label>Min / value</label><input class="dv-min" type="number" /></div>
    <div class="fr-row dv-num dv-max" style="display:none"><label>Max</label><input class="dv-max-in" type="number" /></div>
    <div class="fr-row fr-buttons"><button class="dv-clear">Clear in selection</button><button class="dv-apply">Apply</button></div>
  `;
  document.body.appendChild(el);
  dialog = el;

  const type = el.querySelector(".dv-type") as HTMLSelectElement;
  const listRow = el.querySelector(".dv-list") as HTMLElement;
  const numRows = Array.from(el.querySelectorAll(".dv-num")) as HTMLElement[];
  const maxRow = el.querySelector(".dv-max") as HTMLElement;
  const op = el.querySelector(".dv-op") as HTMLSelectElement;

  const sync = () => {
    const isList = type.value === "list";
    listRow.style.display = isList ? "" : "none";
    numRows.forEach((r) => (r.style.display = isList ? "none" : ""));
    if (!isList) maxRow.style.display = op.value === "between" || op.value === "notBetween" ? "" : "none";
  };
  type.addEventListener("change", sync);
  op.addEventListener("change", sync);
  sync();

  (el.querySelector(".fr-close") as HTMLElement).addEventListener("click", close);
  (el.querySelector(".dv-clear") as HTMLButtonElement).addEventListener("click", () => {
    app.clearValidation();
    close();
  });
  (el.querySelector(".dv-apply") as HTMLButtonElement).addEventListener("click", () => {
    if (type.value === "list") {
      const vals = (el.querySelector(".dv-source") as HTMLInputElement).value
        .split(",")
        .map((s) => s.trim())
        .filter((s) => s !== "");
      if (vals.length) app.addListValidation(vals);
    } else {
      const min = parseFloat((el.querySelector(".dv-min") as HTMLInputElement).value);
      const max = parseFloat((el.querySelector(".dv-max-in") as HTMLInputElement).value);
      const operator = op.value as DataValidation["operator"];
      if (type.value === "textLength") app.addTextLengthValidation(operator, min, isNaN(max) ? undefined : max);
      else app.addNumberValidation(operator, min, isNaN(max) ? undefined : max);
    }
    close();
  });

  el.addEventListener("keydown", (e) => {
    if (e.key === "Escape") close();
  });
}

function close() {
  if (dialog) {
    dialog.remove();
    dialog = null;
  }
}
