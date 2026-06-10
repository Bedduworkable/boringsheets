import "./style.css";
import { App } from "./app.js";

function $(id: string): HTMLElement {
  const el = document.getElementById(id);
  if (!el) throw new Error(`Missing element #${id}`);
  return el;
}

const app = new App({
  canvas: $("grid-canvas") as HTMLCanvasElement,
  container: $("grid-container"),
  editor: $("cell-editor") as HTMLInputElement,
  nameBox: $("name-box") as HTMLInputElement,
  formulaInput: $("formula-input") as HTMLInputElement,
  tabs: $("sheet-tabs"),
  toolbar: $("toolbar"),
});

// Keyboard shortcuts that work everywhere (Electron menu also fires some of
// these on its own, but we handle them here for the browser fallback too).
window.addEventListener("keydown", (e) => {
  // Alt+= : AutoSum (no Cmd/Ctrl).
  if (e.altKey && (e.key === "=" || e.key === "+") && !app.grid.isEditing()) {
    e.preventDefault();
    app.autoSum();
    return;
  }

  const mod = e.metaKey || e.ctrlKey;
  if (!mod) return;
  const k = e.key.toLowerCase();

  // Cmd/Ctrl shortcuts that only apply when NOT editing a cell.
  if (!app.grid.isEditing()) {
    if (k === "d") { e.preventDefault(); app.fillDown(); return; }
    if (k === "r") { e.preventDefault(); app.fillRight(); return; }
    if (e.key === ";" || e.key === ":") {
      e.preventDefault();
      e.shiftKey ? app.insertTime() : app.insertDate();
      return;
    }
  }

  if (k === "b") {
    e.preventDefault();
    app.toggleFormat("bold");
  } else if (k === "i") {
    e.preventDefault();
    app.toggleFormat("italic");
  } else if (k === "u") {
    e.preventDefault();
    app.toggleFormat("underline");
  } else if (k === "z" && !e.shiftKey) {
    e.preventDefault();
    app.undo();
  } else if ((k === "z" && e.shiftKey) || k === "y") {
    e.preventDefault();
    app.redo();
  } else if (k === "s") {
    e.preventDefault();
    e.shiftKey ? app.saveAs() : app.save();
  } else if (k === "o") {
    e.preventDefault();
    app.openFile();
  } else if (k === "n") {
    e.preventDefault();
    app.newWorkbook();
  } else if (k === "f" || k === "h") {
    e.preventDefault();
    app.openFindReplace();
  } else if (k === "a") {
    // Select all — but let inputs (formula bar, cell editor, dialogs) keep their
    // own text select-all.
    const el = document.activeElement;
    if (el && (el.tagName === "INPUT" || el.tagName === "TEXTAREA")) return;
    e.preventDefault();
    app.grid.selectAll();
  }
});

// expose for debugging in the devtools console
(window as any).app = app;
