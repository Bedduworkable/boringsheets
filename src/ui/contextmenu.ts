// A lightweight right-click context menu. Renders an absolutely-positioned
// popup and tears itself down on the next click / Escape / scroll.

export interface MenuItem {
  label?: string;
  action?: () => void;
  separator?: boolean;
  disabled?: boolean;
}

let current: HTMLElement | null = null;

export function showContextMenu(x: number, y: number, items: MenuItem[]) {
  closeContextMenu();

  const menu = document.createElement("div");
  menu.className = "context-menu";
  for (const item of items) {
    if (item.separator) {
      const sep = document.createElement("div");
      sep.className = "ctx-sep";
      menu.appendChild(sep);
      continue;
    }
    const el = document.createElement("div");
    el.className = "ctx-item" + (item.disabled ? " disabled" : "");
    el.textContent = item.label ?? "";
    if (!item.disabled && item.action) {
      el.addEventListener("mousedown", (e) => {
        e.preventDefault();
        closeContextMenu();
        item.action!();
      });
    }
    menu.appendChild(el);
  }

  document.body.appendChild(menu);
  // keep it on-screen
  const rect = menu.getBoundingClientRect();
  const left = Math.min(x, window.innerWidth - rect.width - 4);
  const top = Math.min(y, window.innerHeight - rect.height - 4);
  menu.style.left = `${Math.max(0, left)}px`;
  menu.style.top = `${Math.max(0, top)}px`;
  current = menu;

  setTimeout(() => {
    window.addEventListener("mousedown", onDocClick, { once: true });
    window.addEventListener("keydown", onKey);
    window.addEventListener("wheel", closeContextMenu, { once: true });
  }, 0);
}

function onDocClick() {
  closeContextMenu();
}
function onKey(e: KeyboardEvent) {
  if (e.key === "Escape") closeContextMenu();
}

export function closeContextMenu() {
  if (current) {
    current.remove();
    current = null;
    window.removeEventListener("keydown", onKey);
  }
}
