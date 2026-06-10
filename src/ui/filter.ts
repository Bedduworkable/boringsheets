// Per-column filter dropdown (Google Sheets style): Sort A→Z / Z→A, a search
// box, and a checklist of the column's distinct values. Opened from the funnel
// icon in a filtered column's header.

import type { App } from "../app.js";
import { colToLetter } from "../engine/references.js";

let dialog: HTMLElement | null = null;

export function closeColumnFilter() {
  if (dialog) {
    dialog.remove();
    dialog = null;
  }
}

export function showColumnFilter(app: App, col: number, x: number, y: number) {
  closeColumnFilter();
  const values = app.filterColumnValues(col);
  const hidden = new Set(app.filterColumnHidden(col)); // currently unchecked

  const el = document.createElement("div");
  el.className = "filter-pop";
  el.innerHTML = `
    <div class="flt-sort">
      <div class="flt-sortrow flt-asc">▲&nbsp; Sort A → Z</div>
      <div class="flt-sortrow flt-desc">▼&nbsp; Sort Z → A</div>
    </div>
    <div class="flt-divider"></div>
    <div class="flt-head">Filter by values</div>
    <div class="flt-actions"><a class="flt-all">Select all</a> · <a class="flt-none">Clear</a></div>
    <input class="flt-search" type="text" placeholder="Search…" spellcheck="false" />
    <div class="flt-list"></div>
    <div class="flt-buttons">
      <button class="flt-remove">Remove filter</button>
      <span style="flex:1"></span>
      <button class="flt-cancel">Cancel</button>
      <button class="flt-ok">OK</button>
    </div>`;
  document.body.appendChild(el);
  dialog = el;

  const list = el.querySelector(".flt-list") as HTMLElement;
  const search = el.querySelector(".flt-search") as HTMLInputElement;
  const checks = new Map<string, HTMLInputElement>();

  const render = (q: string) => {
    list.innerHTML = "";
    checks.clear();
    for (const v of values) {
      if (q && !v.toLowerCase().includes(q.toLowerCase())) continue;
      const row = document.createElement("label");
      row.className = "flt-item";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = !hidden.has(v);
      cb.addEventListener("change", () => {
        if (cb.checked) hidden.delete(v);
        else hidden.add(v);
      });
      checks.set(v, cb);
      row.appendChild(cb);
      const span = document.createElement("span");
      span.textContent = v === "" ? "(Blanks)" : v;
      row.appendChild(span);
      list.appendChild(row);
    }
  };
  render("");

  search.addEventListener("input", () => render(search.value));
  (el.querySelector(".flt-all") as HTMLElement).addEventListener("click", () => {
    values.forEach((v) => hidden.delete(v));
    checks.forEach((cb) => (cb.checked = true));
  });
  (el.querySelector(".flt-none") as HTMLElement).addEventListener("click", () => {
    values.forEach((v) => hidden.add(v));
    checks.forEach((cb) => (cb.checked = false));
  });
  (el.querySelector(".flt-asc") as HTMLElement).addEventListener("click", () => {
    app.sortFilterColumn(col, true);
    closeColumnFilter();
  });
  (el.querySelector(".flt-desc") as HTMLElement).addEventListener("click", () => {
    app.sortFilterColumn(col, false);
    closeColumnFilter();
  });
  (el.querySelector(".flt-ok") as HTMLButtonElement).addEventListener("click", () => {
    app.applyColumnFilter(col, [...hidden]);
    closeColumnFilter();
  });
  (el.querySelector(".flt-cancel") as HTMLButtonElement).addEventListener("click", closeColumnFilter);
  (el.querySelector(".flt-remove") as HTMLButtonElement).addEventListener("click", () => {
    app.removeFilter();
    closeColumnFilter();
  });

  // position near the funnel, clamped to the viewport
  const rect = el.getBoundingClientRect();
  el.style.left = `${Math.max(4, Math.min(x, window.innerWidth - rect.width - 6))}px`;
  el.style.top = `${Math.max(4, Math.min(y, window.innerHeight - rect.height - 6))}px`;
  void colToLetter;

  setTimeout(() => {
    const onDoc = (ev: MouseEvent) => {
      if (dialog && !dialog.contains(ev.target as Node)) {
        closeColumnFilter();
        window.removeEventListener("mousedown", onDoc);
      }
    };
    window.addEventListener("mousedown", onDoc);
  }, 0);
}
