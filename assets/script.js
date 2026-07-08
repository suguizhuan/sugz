// PARP7 资讯监控网站 - 交互脚本
(function () {
  const root = document.documentElement;
  const stored = localStorage.getItem("parp7-theme");
  if (stored) root.setAttribute("data-theme", stored);

  document.addEventListener("DOMContentLoaded", () => {
    // Theme toggle
    const btn = document.querySelector(".theme-toggle");
    if (btn) {
      const setLabel = () => {
        const cur = root.getAttribute("data-theme") || "dark";
        btn.textContent = cur === "light" ? "🌙 深色" : "☀️ 浅色";
      };
      setLabel();
      btn.addEventListener("click", () => {
        const cur = root.getAttribute("data-theme") || "dark";
        const next = cur === "light" ? "dark" : "light";
        root.setAttribute("data-theme", next);
        localStorage.setItem("parp7-theme", next);
        setLabel();
      });
    }

    // Mobile menu
    const menuBtn = document.querySelector(".menu-btn");
    const links = document.querySelector(".nav-links");
    if (menuBtn && links) {
      menuBtn.addEventListener("click", () => links.classList.toggle("open"));
    }

    // Active link highlighting
    const path = window.location.pathname.split("/").pop() || "index.html";
    document.querySelectorAll(".nav-links a").forEach(a => {
      const href = a.getAttribute("href");
      if (href === path || (path === "" && href === "index.html")) {
        a.classList.add("active");
      }
    });
  });
})();
