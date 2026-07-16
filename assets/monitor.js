// 康百达在研项目信息监控 —— 数据驱动渲染
(function () {
  "use strict";

  const MODULES = [
    { key: "competitive", label: "竞争格局", icon: "🏁" },
    { key: "clinical", label: "临床更新", icon: "🧪" },
    { key: "patent", label: "专利更新", icon: "📄" },
    { key: "article", label: "文章更新", icon: "📚" },
    { key: "other", label: "其他更新", icon: "📰" },
  ];

  const esc = (s) =>
    String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");

  const state = { data: null, module: "all", query: "" };

  function fmtDate(iso) {
    if (!iso) return "";
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso);
    return m ? `${m[1]}.${m[2]}.${m[3]}` : iso;
  }

  // 合并策展 + 自动抓取的条目（自动条目带 auto:true）
  function targetUpdates(t, auto) {
    const curated = (t.updates || []).map((u) => Object.assign({ auto: false }, u));
    const autoItems = ((auto && auto[t.id]) || []).map((u) => Object.assign({ auto: true }, u));
    return curated.concat(autoItems);
  }

  function moduleUpdates(all, key) {
    return all
      .filter((u) => u.module === key)
      .sort((a, b) => String(b.date || "").localeCompare(String(a.date || "")));
  }

  function matchesQuery(t, all) {
    if (!state.query) return true;
    const q = state.query.toLowerCase();
    const hay = [t.code, t.target, t.drug, t.company, t.indication, t.summary]
      .concat(all.map((u) => [u.title, (u.body || []).join(" ")].join(" ")))
      .join(" ")
      .toLowerCase();
    return hay.indexOf(q) !== -1;
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
    if (opts.showProject) tags.push(`<span class="u-proj">${esc(opts.showProject)}</span>`);
    tags.push(`<span class="u-date">${esc(fmtDate(u.date) || u.dateText || "")}</span>`);
    if (u.tag) tags.push(`<span class="u-tag">${esc(u.tag)}</span>`);
    if (u.auto) tags.push(`<span class="u-auto">自动抓取</span>`);
    const body = (u.body || []).map((p) => `<p>${esc(p)}</p>`).join("");
    const table = u.table ? tableHtml(u.table, "mini") : "";
    const insight = u.insight
      ? `<div class="insight"><b>战略启示：</b>${esc(u.insight)}</div>`
      : "";
    return (
      `<article class="update">` +
      `<div class="u-top">${tags.join("")}</div>` +
      `<h4>${esc(u.title)}</h4>` +
      body + table + insight + sourcesHtml(u.sources) +
      `</article>`
    );
  }

  function moduleHtml(t, all, mod) {
    const items = moduleUpdates(all, mod.key);
    const hasLandscape = mod.key === "competitive" && t.landscape;
    const count = items.length + (hasLandscape ? 0 : 0);
    const openAttr = items.length || hasLandscape ? " open" : "";
    let inner = "";
    if (hasLandscape) {
      if (t.landscape.note) inner += `<p class="landscape-note">${esc(t.landscape.note)}</p>`;
      inner += tableHtml(t.landscape, "mini");
    }
    if (items.length) {
      inner += items.map((u) => updateHtml(u)).join("");
    } else if (!hasLandscape) {
      inner += `<p class="empty-note">暂无更新</p>`;
    }
    const badgeCls = items.length ? "m-count has" : "m-count";
    return (
      `<details class="module" data-module="${mod.key}"${openAttr}>` +
      `<summary><span class="m-icon">${mod.icon}</span>${esc(mod.label)}` +
      `<span class="${badgeCls}">${items.length}</span><span class="caret">▶</span></summary>` +
      `<div class="module-inner">${inner}</div>` +
      `</details>`
    );
  }

  function columnHtml(t, auto) {
    const all = targetUpdates(t, auto);
    const mods = state.module === "all" ? MODULES : MODULES.filter((m) => m.key === state.module);
    const body = mods.map((m) => moduleHtml(t, all, m)).join("");
    const badges = [
      `<span class="badge">${esc(t.target)}</span>`,
      `<span class="badge indigo">${esc(t.indication)}</span>`,
    ].join("");
    return (
      `<section class="column" data-target="${t.id}">` +
      `<header class="column-head">` +
      `<span class="code">${esc(t.code)}</span>` +
      `<h3>${esc(t.drug)}</h3>` +
      `<div class="meta">${badges}</div>` +
      `<p class="desc">${esc(t.summary)}</p>` +
      `</header>` +
      `<div class="column-body">${body}</div>` +
      `</section>`
    );
  }

  function renderBoard() {
    const board = document.getElementById("board");
    const { targets } = state.data;
    const auto = state.data.auto || {};
    const visible = targets.filter((t) => matchesQuery(t, targetUpdates(t, auto)));
    board.innerHTML = visible.length
      ? visible.map((t) => columnHtml(t, auto)).join("")
      : `<p class="empty-note">没有匹配的靶点。</p>`;
  }

  function renderFilters() {
    const wrap = document.getElementById("moduleFilters");
    const auto = state.data.auto || {};
    const counts = { all: 0 };
    MODULES.forEach((m) => (counts[m.key] = 0));
    state.data.targets.forEach((t) => {
      targetUpdates(t, auto).forEach((u) => {
        counts.all++;
        if (counts[u.module] != null) counts[u.module]++;
      });
    });
    const chip = (key, label, icon) =>
      `<button class="chip${state.module === key ? " active" : ""}" data-module="${key}">` +
      `${icon ? icon + " " : ""}${esc(label)}<span class="count">${counts[key] || 0}</span></button>`;
    wrap.innerHTML =
      chip("all", "全部模块", "") +
      MODULES.map((m) => chip(m.key, m.label, m.icon)).join("");
    wrap.querySelectorAll(".chip").forEach((b) =>
      b.addEventListener("click", () => {
        state.module = b.getAttribute("data-module");
        renderFilters();
        renderBoard();
      })
    );
  }

  function renderFeed() {
    const el = document.getElementById("feed");
    if (!el) return;
    const auto = state.data.auto || {};
    const all = [];
    state.data.targets.forEach((t) => {
      targetUpdates(t, auto).forEach((u) => all.push({ u, t }));
    });
    all.sort((a, b) => String(b.u.date || "").localeCompare(String(a.u.date || "")));
    const top = all.slice(0, 8);
    el.innerHTML = top
      .map(({ u, t }) => updateHtml(u, { showProject: t.id })).join("");
  }

  function renderMatrix() {
    const host = document.getElementById("matrix");
    if (!host) return;
    const periods = state.data.meta.periods || [];
    if (!periods.length) { host.innerHTML = ""; return; }
    if (!state.period) state.period = periods[periods.length - 1].id;

    const picker = periods
      .map((p) =>
        `<button class="chip${state.period === p.id ? " active" : ""}" data-period="${p.id}">${esc(p.label)}</button>`
      )
      .join("");

    const auto = state.data.auto || {};
    const head =
      `<tr><th>项目</th>` +
      MODULES.map((m) => `<th>${m.icon}<br>${esc(m.label)}</th>`).join("") +
      `</tr>`;
    const rows = state.data.targets
      .map((t) => {
        const inPeriod = targetUpdates(t, auto).filter((u) => u.period === state.period);
        const cells = MODULES.map((m) => {
          const hit = inPeriod.some((u) => u.module === m.key);
          return `<td>${hit ? '<span class="tick">✓</span>' : '<span class="dash">—</span>'}</td>`;
        }).join("");
        return `<tr><td><span class="proj">${esc(t.id)}</span><span class="mech">${esc(t.target)}</span></td>${cells}</tr>`;
      })
      .join("");

    host.innerHTML =
      `<div class="period-picker">${picker}</div>` +
      `<div class="table-wrap"><table class="matrix"><thead>${head}</thead><tbody>${rows}</tbody></table></div>`;

    host.querySelectorAll(".chip[data-period]").forEach((b) =>
      b.addEventListener("click", () => {
        state.period = b.getAttribute("data-period");
        renderMatrix();
      })
    );
  }

  function renderStats() {
    const auto = state.data.auto || {};
    let updates = 0;
    state.data.targets.forEach((t) => (updates += targetUpdates(t, auto).length));
    const set = (id, v) => { const e = document.getElementById(id); if (e) e.textContent = v; };
    set("stat-targets", state.data.targets.length);
    set("stat-updates", updates);
    const periods = state.data.meta.periods || [];
    set("stat-period", periods.length ? periods[periods.length - 1].label : "—");
    const autoAt = state.data.meta.auto_generated_at || state.data.meta.generated_at;
    set("stat-updated", autoAt ? fmtDate(autoAt.slice(0, 10)) : "—");
  }

  function boot() {
    const search = document.getElementById("search");
    if (search) {
      let tmr = null;
      search.addEventListener("input", () => {
        clearTimeout(tmr);
        tmr = setTimeout(() => { state.query = search.value.trim(); renderBoard(); }, 150);
      });
    }
    renderStats();
    renderMatrix();
    renderFilters();
    renderFeed();
    renderBoard();
  }

  fetch("data/monitor.json?_=" + Date.now())
    .then((r) => {
      if (!r.ok) throw new Error("HTTP " + r.status);
      return r.json();
    })
    .then((d) => { state.data = d; document.addEventListener("DOMContentLoaded", boot); if (document.readyState !== "loading") boot(); })
    .catch((e) => {
      const board = document.getElementById("board");
      if (board) board.innerHTML = `<p class="empty-note">数据加载失败：${esc(e.message)}。请通过本地静态服务器（如 <code>python3 -m http.server</code>）访问本页。</p>`;
    });
})();
