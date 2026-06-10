// Find & Replace dialog. Operates on the active sheet via App methods.

import type { App } from "../app.js";

let dialog: HTMLElement | null = null;

export function showFindReplace(app: App) {
  if (dialog) {
    (dialog.querySelector(".fr-find") as HTMLInputElement)?.focus();
    return;
  }

  const el = document.createElement("div");
  el.className = "find-replace";
  el.innerHTML = `
    <div class="fr-row"><span class="fr-title">Find &amp; Replace</span><span class="fr-close">&#x2715;</span></div>
    <div class="fr-row"><label>Find</label><input class="fr-find" type="text" spellcheck="false" /></div>
    <div class="fr-row"><label>Replace</label><input class="fr-replace" type="text" spellcheck="false" /></div>
    <div class="fr-row"><label class="fr-check"><input type="checkbox" class="fr-case" /> Match case</label></div>
    <div class="fr-row fr-buttons">
      <button class="fr-next">Find Next</button>
      <button class="fr-rep">Replace</button>
      <button class="fr-repall">Replace All</button>
    </div>
    <div class="fr-status"></div>
  `;
  document.body.appendChild(el);
  dialog = el;

  const findIn = el.querySelector(".fr-find") as HTMLInputElement;
  const replIn = el.querySelector(".fr-replace") as HTMLInputElement;
  const caseIn = el.querySelector(".fr-case") as HTMLInputElement;
  const status = el.querySelector(".fr-status") as HTMLElement;

  let matches: { row: number; col: number }[] = [];
  let idx = -1;

  const recompute = () => {
    matches = app.findMatches(findIn.value, caseIn.checked);
    idx = -1;
    status.textContent = findIn.value ? `${matches.length} match${matches.length === 1 ? "" : "es"}` : "";
  };

  const findNext = () => {
    if (!matches.length) recompute();
    if (!matches.length) {
      status.textContent = "No matches";
      return;
    }
    idx = (idx + 1) % matches.length;
    const m = matches[idx];
    app.gotoCell(m.row, m.col);
    status.textContent = `Match ${idx + 1} of ${matches.length}`;
  };

  findIn.addEventListener("input", recompute);
  caseIn.addEventListener("change", recompute);

  (el.querySelector(".fr-next") as HTMLButtonElement).addEventListener("click", findNext);
  (el.querySelector(".fr-rep") as HTMLButtonElement).addEventListener("click", () => {
    if (idx >= 0 && idx < matches.length) {
      const m = matches[idx];
      app.replaceOne(m.row, m.col, findIn.value, replIn.value, caseIn.checked);
      recompute();
      findNext();
    } else {
      findNext();
    }
  });
  (el.querySelector(".fr-repall") as HTMLButtonElement).addEventListener("click", () => {
    const n = app.replaceAll(findIn.value, replIn.value, caseIn.checked);
    status.textContent = `Replaced ${n}`;
    recompute();
  });

  const close = () => {
    el.remove();
    dialog = null;
  };
  (el.querySelector(".fr-close") as HTMLElement).addEventListener("click", close);
  el.addEventListener("keydown", (e) => {
    if (e.key === "Escape") close();
    else if (e.key === "Enter") findNext();
  });

  findIn.focus();
}
