// A small modal text-input dialog. Electron's renderer does not implement
// window.prompt(), so we provide our own promise-based replacement.

export function showPrompt(message: string, defaultValue = ""): Promise<string | null> {
  return new Promise((resolve) => {
    const overlay = document.createElement("div");
    overlay.className = "prompt-overlay";
    overlay.innerHTML = `
      <div class="prompt-box">
        <div class="prompt-msg"></div>
        <input class="prompt-input" type="text" spellcheck="false" />
        <div class="prompt-buttons">
          <button class="prompt-cancel">Cancel</button>
          <button class="prompt-ok">OK</button>
        </div>
      </div>`;
    (overlay.querySelector(".prompt-msg") as HTMLElement).textContent = message;
    const input = overlay.querySelector(".prompt-input") as HTMLInputElement;
    input.value = defaultValue;
    document.body.appendChild(overlay);
    input.focus();
    input.select();

    let done = false;
    const finish = (val: string | null) => {
      if (done) return;
      done = true;
      overlay.remove();
      resolve(val);
    };

    (overlay.querySelector(".prompt-ok") as HTMLButtonElement).addEventListener("click", () => finish(input.value));
    (overlay.querySelector(".prompt-cancel") as HTMLButtonElement).addEventListener("click", () => finish(null));
    overlay.addEventListener("mousedown", (e) => {
      if (e.target === overlay) finish(null);
    });
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        finish(input.value);
      } else if (e.key === "Escape") {
        e.preventDefault();
        finish(null);
      }
    });
  });
}
