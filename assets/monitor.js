// 康百达在研项目信息监控 —— 数据驱动渲染（总览 hub + 单靶点详情，哈希路由）
(function () {
  "use strict";

  // 主题与移动端菜单（自包含，不依赖其它脚本）
  (function initChrome() {
    const root = document.documentElement;
    const KEY = "kbd-monitor-theme";
    const stored = localStorage.getItem(KEY);
    if (stored) root.setAttribute("data-theme", stored);
    document.addEventListener("DOMContentLoaded", () => {
      const btn = document.querySelector(".theme-toggle");
      if (btn) {
        const label = () => {
          const cur = root.getAttribute("data-theme") || "dark";
          btn.textContent = cur === "light" ? "🌙 深色" : "☀️ 浅色";
        };
        label();
        btn.addEventListener("click", () => {
          const cur = root.getAttribute("data-theme") || "dark";
          const next = cur === "light" ? "dark" : "light";
          root.setAttribute("data-theme", next);
          localStorage.setItem(KEY, next);
          label();
        });
      }
      const menuBtn = document.querySelector(".menu-btn");
      const links = document.querySelector(".nav-links");
      if (menuBtn && links) menuBtn.addEventListener("click", () => links.classList.toggle("open"));
    });
  })();

  const MODULES = [
    { key: "competitive", label: "竞争格局", tab: "竞争格局全览", icon: "🏁" },
    { key: "clinical", label: "临床更新", tab: "临床进展", icon: "🧪" },
    { key: "patent", label: "专利更新", tab: "专利更新", icon: "📄" },
    { key: "article", label: "文章更新", tab: "文章发布", icon: "📚" },
    { key: "other", label: "其他更新", tab: "投融资 / 其他", icon: "📰" },
  ];

  const esc = (s) =>
    String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");

  const state = { data: null, query: "", tab: "competitive", detailId: null };

  function fmtDate(iso) {
    if (!iso) return "";
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso);
    return m ? `${m[1]}.${m[2]}.${m[3]}` : iso;
  }

  // 合并策展 + 自动抓取的条目（自动条目带 auto:true）
  function targetUpdates(t) {
    const auto = (state.data && state.data.auto) || {};
    const curated = (t.updates || []).map((u) => Object.assign({ auto: false }, u));
    const autoItems = (auto[t.id] || []).map((u) => Object.assign({ auto: true }, u));
    return curated.concat(autoItems);
  }

  function moduleUpdates(all, key) {
    return all
      .filter((u) => u.module === key)
      .sort((a, b) => String(b.date || "").localeCompare(String(a.date || "")));
  }

  function latestOf(all) {
    return all.slice().sort((a, b) => String(b.date || "").localeCompare(String(a.date || "")))[0] || null;
  }

  function tableHtml(tbl, cls) {
    if (!tbl || !tbl.columns) return "";
    const head = tbl.columns.map((c) => `<th>${esc(c)}</th>`).join("");
    const rows = (tbl.rows || [])
      .map((r) => `<tr>${r.map((c) => `<td>${esc(c)}</td>`).join("")}</tr>`)
      .join("");
    return `<div class="table-scroll"><table class="${cls}"><thead><tr>${head}</tr></thead><tbody>${rows}</tbody></table></div>`;
  }

  function sourcesHtml(sources) {
    if (!sources || !sources.length) return "";
    const items = sources
      .map((s) =>
        s.url
          ? `<li><a href="${esc(s.url)}" target="_blank" rel="noopener">${esc(s.label)}</a></li>`
          : `<li class="no-url">${esc(s.label)}</li>`
      )
      .join("");
    return `<ul class="sources">${items}</ul>`;
  }

  function updateHtml(u, opts) {
    opts = opts || {};
    const tags = [];
    if (opts.latest) tags.push(`<span class="u-latest">最新</span>`);
    tags.push(`<span class="u-date">${esc(fmtDate(u.date) || u.dateText || "")}</span>`);
    if (u.tag) tags.push(`<span class="u-tag">${esc(u.tag)}</span>`);
    if (u.auto) tags.push(`<span class="u-auto">自动抓取</span>`);
    const body = (u.body || []).map((p) => `<p>${esc(p)}</p>`).join("");
    const table = u.table ? tableHtml(u.table, "mini") : "";
    const insight = u.insight
      ? `<div class="insight"><b>战略启示：</b>${esc(u.insight)}</div>`
      : "";
    return (
      `<article class="update${opts.latest ? " is-latest" : ""}">` +
      `<div class="u-top">${tags.join("")}</div>` +
      `<h4>${esc(u.title)}</h4>` +
      body + table + insight + sourcesHtml(u.sources) +
      `</article>`
    );
  }

  /* ---------------- 总览 hub ---------------- */

  function targetCardHtml(t) {
    const all = targetUpdates(t);
    const latest = latestOf(all);
    const chips = MODULES.map((m) => {
      const n = moduleUpdates(all, m.key).length;
      return `<span class="mchip${n ? "" : " zero"}" title="${esc(m.label)}">${m.icon}<b>${n}</b></span>`;
    }).join("");
    return (
      `<a class="target-card" href="#/t/${encodeURIComponent(t.id)}">` +
      `<div class="tc-top"><span class="code">${esc(t.code)}</span><span class="badge">${esc(t.target)}</span></div>` +
      `<h3>${esc(t.drug)}</h3>` +
      `<p class="tc-ind">${esc(t.indication)}</p>` +
      `<p class="tc-sum">${esc(t.summary)}</p>` +
      `<div class="tc-mods">${chips}</div>` +
      `<div class="tc-foot"><span class="tc-latest">最新动态 ${latest ? esc(fmtDate(latest.date)) : "—"}</span>` +
      `<span class="go">进入详情 →</span></div>` +
      `</a>`
    );
  }

  function renderCards() {
    const host = document.getElementById("cards");
    if (!host) return;
    const q = state.query.toLowerCase();
    const list = state.data.targets.filter((t) => {
      if (!q) return true;
      const hay = [t.code, t.target, t.drug, t.company, t.indication, t.summary]
        .concat(targetUpdates(t).map((u) => [u.title, (u.body || []).join(" ")].join(" ")))
        .join(" ").toLowerCase();
      return hay.indexOf(q) !== -1;
    });
    host.innerHTML = list.length
      ? list.map(targetCardHtml).join("")
      : `<p class="empty-note">没有匹配的靶点。</p>`;
  }

  function renderFeed() {
    const el = document.getElementById("feed");
    if (!el) return;
    const all = [];
    state.data.targets.forEach((t) => targetUpdates(t).forEach((u) => all.push({ u, t })));
    all.sort((a, b) => String(b.u.date || "").localeCompare(String(a.u.date || "")));
    el.innerHTML = all.slice(0, 6).map(({ u, t }) => {
      const label = MODULES.find((m) => m.key === u.module);
      return (
        `<a class="feed-item" href="#/t/${encodeURIComponent(t.id)}">` +
        `<div class="fi-top"><span class="u-proj">${esc(t.id)}</span>` +
        `<span class="u-date">${esc(fmtDate(u.date) || u.dateText || "")}</span>` +
        `<span class="fi-mod">${label ? label.icon + " " + esc(label.label) : ""}</span></div>` +
        `<h4>${esc(u.title)}</h4></a>`
      );
    }).join("");
  }

  function renderMatrix() {
    const host = document.getElementById("matrix");
    if (!host) return;
    const periods = state.data.meta.periods || [];
    if (!periods.length) { host.innerHTML = ""; return; }
    if (!state.period) state.period = periods[periods.length - 1].id;
    const picker = periods
      .map((p) => `<button class="chip${state.period === p.id ? " active" : ""}" data-period="${p.id}">${esc(p.label)}</button>`)
      .join("");
    const head = `<tr><th>项目</th>` + MODULES.map((m) => `<th>${m.icon}<br>${esc(m.label)}</th>`).join("") + `</tr>`;
    const rows = state.data.targets.map((t) => {
      const inPeriod = targetUpdates(t).filter((u) => u.period === state.period);
      const cells = MODULES.map((m) => {
        const hit = inPeriod.some((u) => u.module === m.key);
        return `<td>${hit ? '<span class="tick">✓</span>' : '<span class="dash">—</span>'}</td>`;
      }).join("");
      return `<tr data-goto="${esc(t.id)}"><td><span class="proj">${esc(t.id)}</span><span class="mech">${esc(t.target)}</span></td>${cells}</tr>`;
    }).join("");
    host.innerHTML =
      `<div class="period-picker">${picker}</div>` +
      `<div class="table-wrap"><table class="matrix"><thead>${head}</thead><tbody>${rows}</tbody></table></div>`;
    host.querySelectorAll(".chip[data-period]").forEach((b) =>
      b.addEventListener("click", () => { state.period = b.getAttribute("data-period"); renderMatrix(); }));
    host.querySelectorAll("tr[data-goto]").forEach((r) =>
      r.addEventListener("click", () => { location.hash = "#/t/" + encodeURIComponent(r.getAttribute("data-goto")); }));
  }

  function renderStats() {
    let updates = 0;
    state.data.targets.forEach((t) => (updates += targetUpdates(t).length));
    const set = (id, v) => { const e = document.getElementById(id); if (e) e.textContent = v; };
    set("stat-targets", state.data.targets.length);
    set("stat-updates", updates);
    const periods = state.data.meta.periods || [];
    set("stat-period", periods.length ? periods[periods.length - 1].label : "—");
    const autoAt = state.data.meta.auto_generated_at || state.data.meta.generated_at;
    set("stat-updated", autoAt ? fmtDate(autoAt.slice(0, 10)) : "—");
  }

  /* ---------------- 单靶点详情 ---------------- */

  function panelHtml(t, all, key) {
    const items = moduleUpdates(all, key);
    let html = "";
    if (key === "competitive" && t.landscape) {
      html += `<section class="overview-block">` +
        `<div class="ob-head"><span class="ob-badge">竞争格局全览</span>` +
        `<span class="ob-sub">竞品全景（最新一期）</span></div>`;
      if (t.landscape.note) html += `<p class="landscape-note">${esc(t.landscape.note)}</p>`;
      html += tableHtml(t.landscape, "data") + `</section>`;
    }
    if (items.length) {
      html += items.map((u, i) => updateHtml(u, { latest: i === 0 })).join("");
    } else if (!(key === "competitive" && t.landscape)) {
      html += `<p class="empty-note">该模块暂无更新记录。</p>`;
    }
    return `<div class="panel-inner">${html}</div>`;
  }

  function renderDetail(id, keepScroll) {
    const t = state.data.targets.find((x) => x.id === id);
    const host = document.getElementById("detail");
    if (!t) { location.hash = ""; return; }
    state.detailId = id;
    const all = targetUpdates(t);
    const tabs = MODULES.map((m) => {
      const n = moduleUpdates(all, m.key).length;
      const dot = m.key === "competitive" && t.landscape && !n ? "" : `<span class="tcount">${n}</span>`;
      return `<button class="tab${state.tab === m.key ? " active" : ""}" data-tab="${m.key}">` +
        `<span class="ti">${m.icon}</span>${esc(m.tab)}${dot}</button>`;
    }).join("");
    host.innerHTML =
      `<div class="container detail-wrap">` +
      `<a class="back" href="#/">← 返回全部靶点</a>` +
      `<header class="detail-head">` +
      `<span class="code">${esc(t.code)}</span>` +
      `<h1>${esc(t.drug)}</h1>` +
      `<div class="meta"><span class="badge">${esc(t.target)}</span>` +
      `<span class="badge indigo">${esc(t.indication)}</span>` +
      `<span class="badge">${esc(t.company)}</span></div>` +
      `<p class="desc">${esc(t.summary)}</p>` +
      `</header>` +
      `<nav class="tabbar" role="tablist">${tabs}</nav>` +
      `<div class="panel" id="panel">${panelHtml(t, all, state.tab)}</div>` +
      `</div>`;
    host.querySelectorAll(".tab").forEach((b) =>
      b.addEventListener("click", () => {
        state.tab = b.getAttribute("data-tab");
        renderDetail(id, true);
      }));
    if (!keepScroll) window.scrollTo({ top: 0, behavior: "instant" in window ? "instant" : "auto" });
  }

  /* ---------------- 路由 ---------------- */

  function route() {
    const hub = document.getElementById("hub");
    const detail = document.getElementById("detail");
    const m = /^#\/t\/([^/?#]+)/.exec(location.hash || "");
    const id = m ? decodeURIComponent(m[1]) : null;
    if (id && state.data.targets.some((t) => t.id === id)) {
      state.tab = "competitive";
      hub.hidden = true;
      detail.hidden = false;
      renderDetail(id, false);
    } else {
      detail.hidden = true;
      hub.hidden = false;
    }
  }

  function boot() {
    const search = document.getElementById("search");
    if (search) {
      let tmr = null;
      search.addEventListener("input", () => {
        clearTimeout(tmr);
        tmr = setTimeout(() => { state.query = search.value.trim(); renderCards(); }, 150);
      });
    }
    renderStats();
    renderMatrix();
    renderFeed();
    renderCards();
    window.addEventListener("hashchange", route);
    route();
  }

  function start(d) {
    state.data = d;
    if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", boot);
    else boot();
  }

  if (window.__MONITOR_DATA__) {
    start(window.__MONITOR_DATA__);
  } else {
    fetch("data/monitor.json?_=" + Date.now())
      .then((r) => { if (!r.ok) throw new Error("HTTP " + r.status); return r.json(); })
      .then(start)
      .catch((e) => {
        const board = document.getElementById("cards");
        if (board) board.innerHTML = `<p class="empty-note">数据加载失败：${esc(e.message)}。请通过本地静态服务器（如 <code>python3 -m http.server</code>）访问本页。</p>`;
      });
  }
})();
