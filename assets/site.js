// Shared behaviour for every Flash Analyst page: theme, sticky nav, mobile menu, banner, FAQ.
(() => {
  const reduced = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  const root = document.documentElement;

  // Light / dark theme: follows the device until the visitor picks one
  const themeBtn = document.getElementById("theme-btn");
  if (themeBtn) {
    const systemDark = () => window.matchMedia("(prefers-color-scheme: dark)").matches;
    const isDark = () => root.dataset.theme ? root.dataset.theme === "dark" : systemDark();
    const label = () => themeBtn.setAttribute("aria-label", isDark() ? "Switch to light mode" : "Switch to dark mode");
    themeBtn.addEventListener("click", () => {
      root.dataset.theme = isDark() ? "light" : "dark";
      try { localStorage.setItem("fa-theme", root.dataset.theme); } catch (e) {}
      label();
    });
    label();
  }

  const nav = document.querySelector(".nav");
  if (nav) {
    const onScroll = () => nav.classList.toggle("scrolled", window.scrollY > 8);
    window.addEventListener("scroll", onScroll, { passive: true });
    onScroll();
  }

  const menuBtn = document.getElementById("menu-btn");
  const menu = document.getElementById("mobile-menu");
  if (menuBtn && menu) {
    const setMenu = open => {
      menu.hidden = !open;
      menuBtn.setAttribute("aria-expanded", String(open));
      menuBtn.setAttribute("aria-label", open ? "Close menu" : "Open menu");
    };
    menuBtn.addEventListener("click", () => setMenu(menu.hidden));
    menu.addEventListener("click", e => { if (e.target.closest("a")) setMenu(false); });
    document.addEventListener("keydown", e => { if (e.key === "Escape" && !menu.hidden) { setMenu(false); menuBtn.focus(); } });
    window.addEventListener("resize", () => { if (window.innerWidth > 1020) setMenu(false); });
  }

  // Launch banner can be hidden; remembered for this visitor
  const announce = document.getElementById("announce");
  if (announce) {
    try { if (localStorage.getItem("fa-hide-announce") === "1") announce.hidden = true; } catch (e) {}
    document.getElementById("announce-close").addEventListener("click", () => {
      announce.hidden = true;
      try { localStorage.setItem("fa-hide-announce", "1"); } catch (e) {}
    });
  }

  // FAQ: smooth open and close
  document.querySelectorAll(".faq details").forEach(d => {
    const summary = d.querySelector("summary");
    const answer = d.querySelector(".answer");
    summary.addEventListener("click", e => {
      if (reduced) return;
      e.preventDefault();
      if (d.open) {
        answer.style.height = answer.scrollHeight + "px";
        requestAnimationFrame(() => { answer.style.height = "0px"; });
        answer.addEventListener("transitionend", () => { d.open = false; answer.style.height = ""; }, { once: true });
      } else {
        d.open = true;
        const h = answer.scrollHeight;
        answer.style.height = "0px";
        requestAnimationFrame(() => { answer.style.height = h + "px"; });
        answer.addEventListener("transitionend", () => { answer.style.height = ""; }, { once: true });
      }
    });
  });
})();
