// Chart creation dialog: pick a chart type; the chart is built from the current
// selection (first row = series names, first column = category labels).

import type { App } from "../app.js";
import type { ChartType } from "../charts/render.js";

let dialog: HTMLElement | null = null;

const TYPES: { value: ChartType; label: string }[] = [
  { value: "column", label: "Column" },
  { value: "bar", label: "Bar" },
  { value: "line", label: "Line" },
  { value: "area", label: "Area" },
  { value: "pie", label: "Pie" },
  { value: "scatter", label: "Scatter" },
];

export function showChartDialog(app: App) {
  close();
  const el = document.createElement("div");
  el.className = "cf-dialog";
  el.innerHTML = `
    <div class="fr-row"><span class="fr-title">Insert chart</span><span class="fr-close">&#x2715;</span></div>
    <div class="fr-row"><label>Type</label><select class="ch-type"></select></div>
    <div class="fr-row" style="color:var(--muted);font-size:12px">Uses the selected range. First row = series names, first column = category labels.</div>
    <div class="fr-row fr-buttons"><button class="ch-create">Create</button></div>
  `;
  document.body.appendChild(el);
  dialog = el;

  const sel = el.querySelector(".ch-type") as HTMLSelectElement;
  for (const t of TYPES) {
    const o = document.createElement("option");
    o.value = t.value;
    o.textContent = t.label;
    sel.appendChild(o);
  }

  (el.querySelector(".fr-close") as HTMLElement).addEventListener("click", close);
  (el.querySelector(".ch-create") as HTMLButtonElement).addEventListener("click", () => {
    app.addChart(sel.value as ChartType);
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
