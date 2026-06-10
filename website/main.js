/* ============================================================
   BoringSheets site — GSAP interactions
   1) Hamburger → X nav overlay (timeline .clear() + rebuild)
   2) SplitText char-flip on hover (.flip)
   3) SplitText line-reveal wipe on scroll (.reveal)
   ============================================================ */

(function () {
  "use strict";

  const reduce = window.matchMedia("(prefers-reduced-motion: reduce)").matches;

  // Footer year
  const yearEl = document.getElementById("year");
  if (yearEl) yearEl.textContent = String(new Date().getFullYear());

  if (typeof gsap === "undefined") return; // CDN blocked — site still works statically
  if (typeof SplitText !== "undefined" && typeof ScrollTrigger !== "undefined") {
    gsap.registerPlugin(SplitText, ScrollTrigger);
  }

  /* ---------- 1) Nav overlay ---------- */
  const toggleBtn = document.getElementById("menuToggle");
  let isOpen = false;
  let tl = gsap.timeline();

  function openMenu() {
    tl.set("#nav", { visibility: "visible", pointerEvents: "auto" })
      .fromTo(".nav-bg", { opacity: 0 }, { opacity: 1, duration: 0.4, ease: "power2.out" }, 0)
      .fromTo(
        ".nav-panel",
        { x: "101%", y: 0, rotation: 0 },
        { x: "0%", duration: 0.6, ease: "back.out(1.2)", stagger: 0.2 },
        0
      )
      .fromTo(
        ".nav-item",
        { opacity: 0, x: -20 },
        { opacity: 1, x: 0, duration: 1.2, ease: "expo.out", stagger: 0.04 },
        0.1
      )
      .to(".bar-mid", { opacity: 0, duration: 0.2 }, 0.04)
      .fromTo(
        ".bar-top",
        { stroke: "#fffce1", attr: { x1: 3, y1: 7, x2: 17, y2: 7 } },
        { stroke: "#0e100f", attr: { x1: 5, y1: 5, x2: 15, y2: 15 }, duration: 0.35, ease: "back.out(1.4)" },
        0.06
      )
      .fromTo(
        ".bar-bot",
        { stroke: "#fffce1", attr: { x1: 3, y1: 13, x2: 17, y2: 13 } },
        { stroke: "#0e100f", attr: { x1: 15, y1: 5, x2: 5, y2: 15 }, duration: 0.35, ease: "back.out(1.4)" },
        0.06
      )
      .fromTo(
        ".nav-login",
        { opacity: 0, y: 8 },
        { opacity: 1, y: 0, duration: 0.3, ease: "power3.out" },
        0.4
      );
  }

  function closeMenu() {
    tl.to(".bar", { stroke: "#fffce1", duration: 0.2 })
      .to(".bar-mid", { opacity: 1, duration: 0.2 }, "<")
      .to(".bar-top", { attr: { x1: 3, y1: 7, x2: 17, y2: 7 }, duration: 0.2, ease: "power3.in" }, "<")
      .to(".bar-bot", { attr: { x1: 3, y1: 13, x2: 17, y2: 13 }, duration: 0.2, ease: "power3.in" }, "<")
      .to(
        ".nav-panel",
        {
          y: "160vh",
          rotation: "random(-15, 15)",
          duration: 1,
          ease: "power3.in",
          stagger: { from: "end", each: 0.02 },
        },
        "<"
      )
      .to(".nav-bg", { opacity: 0, duration: 0.3, ease: "power2.in" }, "<0.1")
      .set("#nav", { visibility: "hidden", pointerEvents: "none" });
  }

  function toggle() {
    isOpen = !isOpen;
    toggleBtn.setAttribute("aria-expanded", String(isOpen));
    toggleBtn.setAttribute("aria-label", isOpen ? "Close menu" : "Open menu");
    const tabindex = isOpen ? "0" : "-1";
    document.querySelectorAll(".nav-link, .nav-socials a").forEach((l) => l.setAttribute("tabindex", tabindex));

    tl.clear();
    if (isOpen) openMenu();
    else closeMenu();
  }

  if (toggleBtn) {
    toggleBtn.addEventListener("click", toggle);
    document.querySelector(".nav-bg").addEventListener("click", () => { if (isOpen) toggle(); });
    document.querySelectorAll(".nav-link").forEach((l) =>
      l.addEventListener("click", () => { if (isOpen) toggle(); })
    );
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape" && isOpen) { toggle(); toggleBtn.focus(); }
    });
  }

  /* ---------- 2) SplitText char-flip on hover ---------- */
  if (typeof SplitText !== "undefined" && !reduce) {
    document.querySelectorAll(".flip").forEach((el) => {
      SplitText.create(el, {
        type: "words,chars",
        wordsClass: "word",
        charsClass: "char",
        mask: "words",
        autoSplit: true,
        onSplit: (self) => {
          self.words.forEach((word) => {
            const content = word.innerHTML;
            word.innerHTML = "";
            const before = document.createElement("span");
            before.className = "before";
            before.innerHTML = content;
            const after = document.createElement("span");
            after.className = "after";
            after.setAttribute("aria-hidden", "true");
            after.innerHTML = content;
            word.append(before, after);
          });
        },
      });
    });
  }

  /* ---------- 3) Line-reveal wipe on scroll ---------- */
  if (typeof SplitText !== "undefined" && typeof ScrollTrigger !== "undefined" && !reduce) {
    document.querySelectorAll(".reveal").forEach((el) => {
      SplitText.create(el, {
        type: "lines",
        linesClass: "line",
        autoSplit: true,
        onSplit: (self) => {
          if (el.dataset.played === "true") return;

          self.lines.forEach((line) => {
            const content = line.innerHTML;
            line.innerHTML = "";
            const wrap = document.createElement("span");
            wrap.className = "lwrap";
            wrap.innerHTML = content;
            const overlay = document.createElement("span");
            overlay.className = "loverlay";
            line.append(wrap, overlay);
          });

          const build = () => {
            const t = gsap.timeline({ onComplete: () => { el.dataset.played = "true"; } });
            self.lines.forEach((line, i) => {
              const wrap = line.querySelector(".lwrap");
              const overlay = line.querySelector(".loverlay");
              const start = i * 0.12;
              t.to(overlay, { width: "100%", ease: "power1.in", duration: 0.34 }, start)
                .set(wrap, { opacity: 1 }, start + 0.34)
                .to(overlay, { scaleX: 0, transformOrigin: "right center", ease: "power3.out", duration: 0.4 }, start + 0.34);
            });
            return t;
          };

          ScrollTrigger.create({
            trigger: el,
            start: "top 88%",
            once: true,
            onEnter: build,
          });
        },
      });
    });
  } else {
    // reduced motion / no plugins: ensure reveal text is visible
    document.querySelectorAll(".reveal").forEach((el) => (el.style.opacity = "1"));
  }
})();
