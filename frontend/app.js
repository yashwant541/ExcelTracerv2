/* Excel Formula Tracer - front end (vanilla JS, no build step). */
(function () {
  "use strict";

  function apiBase() {
    try { if (typeof getWebAppBackendUrl === "function") return getWebAppBackendUrl(""); }
    catch (e) {}
    return "";
  }
  const API = apiBase().replace(/\/$/, "");

  async function call(path, opts) {
    opts = opts || {};
    const headers = opts.headers || {};
    if (opts.json !== undefined) { headers["Content-Type"] = "application/json"; opts.body = JSON.stringify(opts.json); }
    const res = await fetch(API + path, { method: opts.method || "GET", headers, body: opts.body });
    let data = null;
    try { data = await res.json(); } catch (e) {}
    if (!res.ok || (data && data.success === false)) {
      let msg = (data && (data.error || data.message)) || ("HTTP " + res.status);
      if (data && data.details) msg += " — " + (typeof data.details === "string" ? data.details : JSON.stringify(data.details));
      throw new Error(msg);
    }
    return data;
  }

  const $ = (s, r) => (r || document).querySelector(s);
  const $$ = (s, r) => Array.prototype.slice.call((r || document).querySelectorAll(s));
  const SVGNS = "http://www.w3.org/2000/svg";
  const el = (tag, attrs, kids) => {
    const svg = ["svg", "path", "rect", "g", "text", "line", "defs", "marker", "polygon", "circle"].indexOf(tag) >= 0;
    const n = svg ? document.createElementNS(SVGNS, tag) : document.createElement(tag);
    Object.entries(attrs || {}).forEach(([k, v]) => {
      if (v == null) return;
      if (k === "class") n.setAttribute("class", v);
      else if (k === "text") n.textContent = v;
      else if (k === "html") n.innerHTML = v;
      else if (k.slice(0, 2) === "on") n.addEventListener(k.slice(2), v);
      else n.setAttribute(k, v);
    });
    (kids || []).forEach((c) => c != null && n.appendChild(typeof c === "string" ? document.createTextNode(c) : c));
    return n;
  };
  function toast(msg, kind) {
    const t = el("div", { class: "toast " + (kind || ""), text: msg });
    $("#toast-wrap").appendChild(t);
    setTimeout(() => t.remove(), 4200);
  }
  const fmt = (v) => (v === null || v === undefined) ? "∅" : (typeof v === "number" ? trimNum(v) : String(v));
  function trimNum(n) {
    if (typeof n !== "number") return String(n);
    if (!isFinite(n)) return n > 0 ? "∞" : "-∞";
    return Number.isInteger(n) ? String(n) : String(Math.round(n * 1e10) / 1e10);
  }
  function show(view) {
    $$(".view").forEach((v) => v.classList.add("hidden"));
    $("#view-" + view).classList.remove("hidden");
    window.scrollTo(0, 0);
  }

  let S = { sid: null, filename: null, summary: null };

  // ================================================================ OPEN
  async function boot() {
    $("#nav-new").addEventListener("click", () => { S = { sid: null }; show("open"); $("#nav-new").classList.add("hidden"); });
    $("#folder-list").addEventListener("click", listFiles);
    $("#folder-open").addEventListener("click", openFromFolder);
    $("#browse-btn").addEventListener("click", () => $("#file-input").click());
    $("#file-input").addEventListener("change", (e) => uploadFile(e.target.files[0]));
    const dz = $("#dropzone");
    ["dragover", "dragenter"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add("drag"); }));
    ["dragleave", "drop"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove("drag"); }));
    dz.addEventListener("drop", (e) => uploadFile(e.dataTransfer.files[0]));

    $("#tr-go").addEventListener("click", loadRowJourney);
    $("#tr-cell-go").addEventListener("click", () => { const c = $("#tr-cell").value.trim(); if (c) explainCell(c); });
    $("#tr-scan").addEventListener("click", scanMismatch);
    $("#tr-fx-filter").addEventListener("input", debounce(loadFormulas, 250));
    $("#d-close").addEventListener("click", closeDrawer);
    $("#drawer-scrim").addEventListener("click", closeDrawer);
    $("#d-flow").addEventListener("click", () => { if (DRAWER) { closeDrawer(); showFlowchart(DRAWER.cell); } });
    $("#d-steps").addEventListener("click", () => { if (DRAWER) { closeDrawer(); showSteps(DRAWER.cell); } });
    $$("#d-tabs button").forEach((b) => b.addEventListener("click", () => showTab(b.dataset.tab)));
    $("#flow-back").addEventListener("click", () => show("trace"));
    $("#steps-back").addEventListener("click", () => show("trace"));
    $("#dbg-first").addEventListener("click", () => dbgSet(0));
    $("#dbg-prev").addEventListener("click", () => dbgSet(STEPS.i - 1));
    $("#dbg-next").addEventListener("click", () => dbgSet(STEPS.i + 1));
    $("#dbg-last").addEventListener("click", () => dbgSet(STEPS.data.steps.length - 1));
    $("#dbg-play").addEventListener("click", dbgPlay);
    $("#steps-copy").addEventListener("click", () => { navigator.clipboard.writeText($("#steps-why").textContent); toast("Copied", "ok"); });
    $("#steps-json").addEventListener("click", dbgDownload);
    $("#flow-zin").addEventListener("click", () => FLOW && FLOW._zoom && FLOW._zoom(0.8));
    $("#flow-zout").addEventListener("click", () => FLOW && FLOW._zoom && FLOW._zoom(1.25));
    $("#flow-fit").addEventListener("click", () => FLOW && FLOW._fit && FLOW._fit());
    $("#flow-hide-const").addEventListener("change", buildAndRenderFlow);
    $("#flow-only-taken").addEventListener("change", buildAndRenderFlow);
    document.addEventListener("keydown", (e) => { if (e.key === "Escape") closeDrawer(); });

    try {
      const h = await call("/api/health");
      if (!h.dataiku) {
        $("#folder-hint").textContent = "Not running inside Dataiku — use Upload.";
        $("#folder-select").disabled = $("#folder-list").disabled = $("#folder-open").disabled = true;
      } else {
        const r = await call("/api/folders");
        const sel = $("#folder-select"); sel.innerHTML = "<option value=''>— choose —</option>";
        (r.folders || []).forEach((f) => sel.appendChild(el("option", { value: f.id, text: `${f.name}  (${f.id})` })));
      }
    } catch (e) { /* offline is fine, upload still works */ }
  }
  function debounce(fn, ms) { let t; return () => { clearTimeout(t); t = setTimeout(fn, ms); }; }
  const folderId = () => $("#folder-id").value.trim() || $("#folder-select").value;

  async function listFiles() {
    $("#folder-err").textContent = "";
    const fid = folderId();
    if (!fid) return ($("#folder-err").textContent = "Pick a folder or paste an id.");
    try {
      const r = await call(`/api/folder-files?folder_id=${encodeURIComponent(fid)}`);
      const sel = $("#folder-file"); sel.innerHTML = "";
      (r.files || []).forEach((p) => sel.appendChild(el("option", { value: p, text: p })));
      if (!(r.files || []).length) sel.appendChild(el("option", { value: "", text: "no .xlsx / .xlsm files" }));
    } catch (e) { $("#folder-err").textContent = e.message; }
  }

  async function openFromFolder() {
    $("#folder-err").textContent = "";
    const fid = folderId();
    const path = $("#folder-path").value.trim() || $("#folder-file").value;
    if (!fid || !path) return ($("#folder-err").textContent = "Folder and file are required.");
    try {
      const r = await call("/api/open", { method: "POST", json: { folder_id: fid, path } });
      enterTrace(r);
    } catch (e) { $("#folder-err").textContent = e.message; }
  }

  async function uploadFile(file) {
    $("#upload-err").textContent = "";
    if (!file) return;
    const fd = new FormData(); fd.append("file", file);
    try {
      const r = await call("/api/upload", { method: "POST", body: fd });
      enterTrace(r);
    } catch (e) { $("#upload-err").textContent = e.message; }
  }

  // ================================================================ TRACE
  function enterTrace(r) {
    S = { sid: r.session_id, filename: r.filename, summary: r.summary };
    show("trace");
    $("#nav-new").classList.remove("hidden");
    $("#trace-title").textContent = r.filename;
    const s = r.summary;
    if (!s.engine_build) {
      $("#trace-summary").innerHTML = "<span class='errbox'>An old <code>excel_deep_trace.py</code> is on the project libraries — replace it and restart the backend.</span>";
    } else {
      $("#trace-summary").textContent =
        `${s.total_formula_cells} formula cells · ${s.sheet_names.length} sheet(s) · `
        + `${s.conditional_cells} conditional · ${s.lookup_cells} lookup · ${s.findings.total} finding(s)`
        + (s.unparsed_cells ? ` · ${s.unparsed_cells} unparsed` : "");
    }
    const sheetSel = $("#tr-sheet"); sheetSel.innerHTML = "";
    (s.sheet_names || []).forEach((n) => sheetSel.appendChild(el("option", { value: n, text: n })));
    const busiest = (s.sheets || []).slice().sort((a, b) => b.formula_cells - a.formula_cells)[0];
    if (busiest) sheetSel.value = busiest.name;
    $("#tr-body").innerHTML = "<p class='muted'>Pick a row or a cell on the left to start.</p>";
    $("#tr-mismatch-list").innerHTML = "";
    loadFindings();
    loadFormulas();
  }

  async function loadFindings() {
    const box = $("#tr-findings");
    try {
      const r = await call(`/api/findings?session_id=${S.sid}`);
      box.innerHTML = "";
      if (!(r.findings || []).length) { box.innerHTML = "<p class='muted small'>No findings.</p>"; return; }
      r.findings.slice(0, 60).forEach((f) => {
        const sev = { HIGH: "err", MEDIUM: "warn", LOW: "neutral", INFO: "neutral" }[f.severity] || "neutral";
        box.appendChild(el("div", { class: "rj-narr", style: "margin-bottom:6px;cursor:pointer", onclick: () => explainCell(f.cell) }, [
          el("span", { class: "pill " + sev, text: f.severity }), " ",
          el("span", { class: "mono small", text: f.cell }),
          el("p", { class: "muted small", text: f.message }),
        ]));
      });
    } catch (e) { box.innerHTML = "<p class='errbox'>" + e.message + "</p>"; }
  }

  async function loadFormulas() {
    const box = $("#tr-formulas");
    const q = $("#tr-fx-filter").value.trim().toLowerCase();
    try {
      const r = await call(`/api/formulas?session_id=${S.sid}`);
      const rows = (r.formulas || []).filter((f) =>
        !q || f.key.toLowerCase().includes(q) || (f.functions || []).join(",").toLowerCase().includes(q) || (f.formula || "").toLowerCase().includes(q));
      box.innerHTML = "";
      box.appendChild(el("p", { class: "muted small", text: `${rows.length} / ${r.total}` }));
      rows.slice(0, 200).forEach((f) => box.appendChild(el("div", { class: "rj-narr", style: "cursor:pointer;margin-bottom:4px", onclick: () => explainCell(f.key) }, [
        el("span", { class: "mono small", text: f.cell }), " ",
        el("span", { class: "muted small mono", text: (f.formula || "").slice(0, 46) }),
      ])));
    } catch (e) { box.innerHTML = "<p class='errbox'>" + e.message + "</p>"; }
  }

  async function scanMismatch() {
    const box = $("#tr-mismatch-list");
    const left = $("#tr-left").value.trim(), right = $("#tr-right").value.trim();
    if (!left || !right) return toast("Enter two columns", "err");
    box.innerHTML = "<p class='muted small'>Scanning…</p>";
    try {
      const r = await call("/api/mismatch", { method: "POST", json: { session_id: S.sid, sheet: $("#tr-sheet").value, left, right } });
      box.innerHTML = `<p class='muted small'>${r.total} mismatch row(s)</p>`;
      (r.rows || []).slice(0, 80).forEach((row) => box.appendChild(el("div", { class: "rj-narr", style: "cursor:pointer", onclick: () => { $("#tr-row").value = row.row; loadRowJourney(); } }, [
        el("span", { class: "mono small", text: "row " + row.row }), " ",
        el("span", { class: "muted small", text: `${fmt(row.left)} vs ${fmt(row.right)}` }),
      ])));
    } catch (e) { box.innerHTML = "<p class='errbox'>" + e.message + "</p>"; }
  }

  async function loadRowJourney() {
    const sheet = $("#tr-sheet").value, row = parseInt($("#tr-row").value, 10);
    if (!row) return toast("Enter a row number", "err");
    const box = $("#tr-body"); box.innerHTML = "<p class='muted'>Building row journey…</p>";
    try {
      const r = await call("/api/journey", { method: "POST", json: { session_id: S.sid, sheet, row } });
      const j = r.journey;
      box.innerHTML = "";
      box.appendChild(el("div", { class: "view-head" }, [
        el("h1", { style: "font-size:17px", text: `${sheet} · row ${row}` }),
        el("p", { class: "muted small", text: `${j.formula_cells} computed cell(s), ${j.input_cells} input(s)` }),
      ]));
      if ((j.inputs || []).length) box.appendChild(el("div", { class: "chip-list", style: "margin-bottom:14px" },
        j.inputs.map((c) => el("span", { class: "chip", text: `${c.header || c.cell}: ${fmt(c.value)}` }))));
      (j.cells || []).forEach((c) => box.appendChild(el("div", { class: "rj-cell" + (c.tracer_reproduced_excel === false ? " mismatch" : "") }, [
        el("div", { class: "rj-head" }, [
          el("span", { class: "rj-name", text: c.header || c.cell }),
          el("span", { class: "rj-key", text: c.key }),
          el("span", { class: "rj-val", text: c.value_repr }),
        ]),
        el("div", { class: "rj-formula", text: c.formula }),
        el("div", { class: "rj-narr" }, (c.narrative || []).map((n) => el("p", { text: n }))),
        el("div", { class: "rj-actions" }, [
          el("button", { class: "btn ghost small", text: "Full trace →", onclick: () => explainCell(c.key) }),
          el("button", { class: "btn ghost small", text: "Steps →", onclick: () => showSteps(c.key) }),
          el("button", { class: "btn ghost small", text: "Flowchart →", onclick: () => showFlowchart(c.key) }),
        ]),
      ])));
      if (!(j.cells || []).length) box.appendChild(el("p", { class: "muted", text: "No formula cells on this row." }));
    } catch (e) { box.innerHTML = "<div class='errbox'>" + e.message + "</div>"; }
  }

  // ================================================================ DRAWER
  const DSTACK = [];
  let DRAWER = null;
  async function explainCell(cell, push) {
    try {
      const [treeR, explainR] = await Promise.all([
        call("/api/tree", { method: "POST", json: { session_id: S.sid, cell } }),
        call("/api/explain", { method: "POST", json: { session_id: S.sid, cell } }).catch(() => null),
      ]);
      if (push) DSTACK.push(cell); else { DSTACK.length = 0; DSTACK.push(cell); }
      DRAWER = { cell, tree: treeR.tree, explain: explainR };
      $("#drawer").classList.remove("hidden"); $("#drawer-scrim").classList.remove("hidden");
      renderDrawer(); showTab("tree");
    } catch (e) { toast(e.message, "err"); }
  }
  function closeDrawer() { $("#drawer").classList.add("hidden"); $("#drawer-scrim").classList.add("hidden"); }

  function renderDrawer() {
    const t = DRAWER.tree;
    const cr = $("#d-crumbs"); cr.innerHTML = "";
    DSTACK.forEach((k, i) => {
      if (i) cr.appendChild(el("span", { class: "sep", text: " ▸ " }));
      if (i === DSTACK.length - 1) cr.appendChild(el("span", { class: "cur", text: k }));
      else cr.appendChild(el("button", { text: k, onclick: () => { DSTACK.length = i + 1; explainCell(k); } }));
    });
    $("#d-col").textContent = t.key;
    $("#d-val").textContent = t.value_repr;
    $("#d-dtype").textContent = (t.root && t.root.dtype) || "";
    $("#d-sub").textContent = t.formula || "";
    const flags = $("#d-flags"); flags.innerHTML = "";
    if (t.is_formula && t.no_excel_value) {
      flags.appendChild(el("div", { class: "flag info" }, [el("span", { text: "•" }),
        el("span", { text: "This workbook stored no computed value for this cell — the value shown is the tracer's own evaluation of the formula." })]));
    } else if (t.is_formula && t.tracer_reproduced_excel === false) {
      flags.appendChild(el("div", { class: "flag warn" }, [el("span", { text: "⚠" }),
        el("span", { html: `The tracer couldn't fully reproduce Excel's value (unsupported or volatile function). Excel's saved value is <b>${t.excel_value_repr}</b>; the breakdown below is partial.` })]));
    }
    (t.notes || []).forEach((n) => flags.appendChild(el("div", { class: "flag info" }, [el("span", { text: "•" }), el("span", { text: n })])));
    renderTree(); renderFlow(); renderExplain(); renderGraph(); renderSource();
  }
  function showTab(name) {
    $$("#d-tabs button").forEach((b) => b.classList.toggle("active", b.dataset.tab === name));
    $$("#drawer .tabpane").forEach((p) => p.classList.add("hidden"));
    $("#d-tab-" + name).classList.remove("hidden");
  }

  const OPLABEL = { if: "IF", op: "op", compare: "cmp", concat: "&", logical: "bool", func: "fn", lookup: "lookup",
    agg: "agg", cell: "cell", name: "name", number: "num", text: "txt", bool: "bool", empty: "∅", error: "err",
    range: "range", unary: "±", skipped: "—", unsupported: "?" };

  function renderTree() {
    const box = $("#d-tab-tree"); box.innerHTML = "";
    if (!DRAWER.tree.root) { box.innerHTML = "<p class='muted'>" + (DRAWER.tree.parse_error ? "Formula could not be parsed." : "No breakdown.") + "</p>"; return; }
    box.appendChild(nodeEl(DRAWER.tree.root, true));
  }
  function nodeEl(n, open) {
    const kids = n.children || [];
    const wrap = el("div", { class: "tnode" + (n.taken === true ? " branch-taken" : n.taken === false ? " branch-skip" : "") + (n.kind === "error" ? " is-error" : "") });
    const has = kids.length > 0;
    const tog = el("span", { class: "tn-toggle", text: has ? (open ? "▾" : "▸") : "·" });
    const kb = el("div", { class: "tn-kids" });
    if (has) {
      kids.forEach((c) => kb.appendChild(nodeEl(c, c.evaluated !== false)));
      if (!open) kb.classList.add("hidden");
      tog.style.cursor = "pointer";
      tog.addEventListener("click", () => { const h = kb.classList.toggle("hidden"); tog.textContent = h ? "▸" : "▾"; });
    }
    const opc = n.kind === "if" ? "op-if" : n.kind === "lookup" ? "op-lookup" : n.kind === "cell" ? "op-cell" : n.kind === "error" ? "op-error" : "";
    const expr = el("span", { class: "tn-expr" });
    if (n.kind === "cell" && n.detail && n.detail.is_formula && n.detail.key) {
      expr.appendChild(el("span", { class: "cell-link", text: n.excel, title: "trace " + n.detail.key, onclick: () => explainCell(n.detail.key, true) }));
    } else expr.textContent = n.excel;
    const row = el("div", { class: "tn-row" }, [tog, el("span", { class: "tn-op " + opc, text: OPLABEL[n.kind] || n.kind }), expr]);
    if (n.evaluated !== false) {
      row.appendChild(el("span", { class: "tn-arrow", text: "=" }));
      row.appendChild(el("span", { class: "tn-val" + (n.value === null ? " empty" : ""), text: n.value_repr }));
    } else row.appendChild(el("span", { class: "tn-tag skip", text: "not taken" }));
    if (n.taken === true && n.role !== "root") row.appendChild(el("span", { class: "tn-tag taken", text: "taken" }));
    if (n.role && ["cond", "then", "else", "key", "value", "fallback"].indexOf(n.role) >= 0) row.appendChild(el("span", { class: "tn-role", text: n.role }));
    wrap.appendChild(row);
    const det = detailEl(n);
    if (det) wrap.appendChild(det);
    if (has) wrap.appendChild(kb);
    return wrap;
  }
  function detailEl(n) {
    const d = n.detail || {};
    if (n.kind === "lookup" && d.table_preview) {
      const tbl = el("table", { class: "lk-table" });
      d.table_preview.forEach((r) => tbl.appendChild(el("tr", { class: r.matched ? "match" : "" }, (r.cells || []).map((c) => el("td", { text: fmt(c) })))));
      return el("div", { class: "tn-detail" }, [
        el("div", { class: "muted small", text: `${d.function} key ${d.lookup_repr} · ${d.match_found ? "matched row " + (d.matched_row_index + 1) : "no match"}` }), tbl,
      ]);
    }
    if (n.kind === "error") return el("div", { class: "tn-detail muted small", text: d.error || "" });
    if (d.short_circuit_at != null) return el("div", { class: "tn-detail muted small", text: `${d.operator} short-circuited at argument ${d.short_circuit_at + 1}` });
    if (n.kind === "if" && d.cond_empty) return el("div", { class: "tn-detail muted small", text: "condition empty → treated as TRUE" });
    return null;
  }

  function renderFlow() {
    const box = $("#d-tab-flow"); box.innerHTML = "";
    if (!DRAWER.tree.root) return;
    const steps = [];
    (function walk(n) {
      (n.children || []).forEach((c) => { if (c.evaluated !== false) walk(c); });
      if (n.evaluated === false) return;
      if (["number", "text", "bool", "empty"].indexOf(n.kind) >= 0) return;
      steps.push(n);
    })(DRAWER.tree.root);
    box.appendChild(el("p", { class: "muted small", text: "Every evaluated sub-expression, innermost first." }));
    const list = el("div", { class: "flow" });
    steps.forEach((n, i) => list.appendChild(el("div", { class: "flow-step" + (i === steps.length - 1 ? " final" : "") }, [
      el("span", { class: "fi", text: i + 1 }), el("span", { class: "fe", text: n.excel }), el("span", { class: "fv", text: n.value_repr }),
    ])));
    box.appendChild(list);
  }

  function renderExplain() {
    const box = $("#d-tab-explain"); box.innerHTML = "";
    (DRAWER.tree.narrative || []).forEach((line) => box.appendChild(el("p", { style: "margin:0 0 10px;font-size:13.5px", text: line })));
    if ((DRAWER.tree.notes || []).length) {
      box.appendChild(el("p", { class: "muted small", text: "Notes:" }));
      const ul = el("ul", {});
      DRAWER.tree.notes.forEach((n) => ul.appendChild(el("li", { class: "muted small", text: n })));
      box.appendChild(ul);
    }
  }

  function renderGraph() {
    const box = $("#d-tab-graph"); box.innerHTML = "";
    const g = DRAWER.explain && DRAWER.explain.precedent_graph;
    if (!g || !g.nodes || !g.nodes.length) { box.innerHTML = "<p class='muted'>No precedent graph.</p>"; return; }
    const NW = 150, NH = 26, GX = 40, GY = 12, PAD = 12;
    const byDepth = {};
    g.nodes.forEach((n) => { (byDepth[n.depth] = byDepth[n.depth] || []).push(n); });
    const depths = Object.keys(byDepth).map(Number).sort((a, b) => b - a);
    const pos = {}; let maxRows = 0;
    depths.forEach((d, di) => {
      byDepth[d].forEach((n, ri) => { pos[n.id] = { x: PAD + di * (NW + GX), y: PAD + ri * (NH + GY) }; });
      maxRows = Math.max(maxRows, byDepth[d].length);
    });
    const W = PAD * 2 + depths.length * (NW + GX) - GX, H = PAD * 2 + maxRows * (NH + GY) - GY;
    const svg = el("svg", { viewBox: `0 0 ${W} ${H}`, width: W, height: H });
    (g.edges || []).forEach((e) => {
      const a = pos[e.from], b = pos[e.to]; if (!a || !b) return;
      const x1 = a.x + NW, y1 = a.y + NH / 2, x2 = b.x, y2 = b.y + NH / 2, mx = (x1 + x2) / 2;
      svg.appendChild(el("path", { class: "gedge", d: `M${x1},${y1} C${mx},${y1} ${mx},${y2} ${x2},${y2}` }));
    });
    g.nodes.forEach((n) => {
      const p = pos[n.id];
      const cls = n.id === g.root ? "root" : n.is_formula ? "formula" : "input";
      const grp = el("g", { class: "gnode " + cls, transform: `translate(${p.x},${p.y})`, onclick: () => n.is_formula && explainCell(n.id, true) }, [
        el("rect", { width: NW, height: NH, rx: 7 }),
        el("text", { x: 8, y: NH / 2 + 3, text: `${n.cell}  ${n.value_repr}`.slice(0, 24) }),
      ]);
      grp.appendChild(el("title", { text: n.id + (n.formula ? "\n" + n.formula : "") }));
      svg.appendChild(grp);
    });
    box.appendChild(svg);
  }

  function renderSource() {
    const box = $("#d-tab-source"); box.innerHTML = "";
    const t = DRAWER.tree, x = DRAWER.explain || {};
    box.appendChild(el("h4", { text: "Formula" }));
    box.appendChild(el("pre", { class: "code", text: t.formula || "—" }));
    const meta = x.formula_metadata || {};
    box.appendChild(el("dl", { class: "kv" }, [
      el("dt", { text: "Family" }), el("dd", { text: t.family || meta.formula_family || "–" }),
      el("dt", { text: "Functions" }), el("dd", { text: (t.functions || meta.functions || []).join(", ") || "–" }),
      el("dt", { text: "Dependencies" }), el("dd", { text: meta.dependency_count != null ? meta.dependency_count : "–" }),
      el("dt", { text: "Excel value" }), el("dd", { text: t.excel_value_repr }),
      el("dt", { text: "Tracer value" }), el("dd", { text: t.value_repr }),
    ]));
  }

  // ================================================================ FLOWCHART
  // A backtracking flowchart for ANY cell, generated from the live expression tree.
  // No external library: a tidy-tree layout + SVG + viewBox pan/zoom.
  let FLOW = null;
  const clip = (s, n) => { s = String(s == null ? "" : s); return s.length > n ? s.slice(0, n - 1) + "…" : s; };

  const FLOW_LEGEND = [
    ["fn-formula", "#1565C0", "formula — a cell driven by a formula (click to re-root)"],
    ["fn-lookup", "#7E57C2", "lookup — VLOOKUP / INDEX / MATCH"],
    ["fn-source", "#43A047", "source — input cell / constant / matched value"],
    ["fn-calculation", "#F9A825", "calculation — arithmetic & functions"],
    ["fn-condition", "#FB8C00", "condition — IF / AND / OR / IFERROR"],
    ["fn-final", "#00796B", "the result of this formula"],
    ["fn-error", "#E53935", "error"],
    ["fn-inactive", "#9E9E9E", "branch not taken"],
  ];
  function renderFlowLegend() {
    const box = $("#flow-legend"); box.innerHTML = "";
    FLOW_LEGEND.forEach(([, c, label]) => box.appendChild(el("span", { class: "legend-item" }, [
      el("span", { class: "dot", style: `background:${c}` }), label])));
  }

  function flowKind(n, isRoot) {
    if (n.evaluated === false || n.kind === "skipped") return "inactive";
    if (n.kind === "error") return "error";
    if (isRoot) return "final";
    switch (n.kind) {
      case "if": case "logical": case "iferror": return "condition";
      case "lookup": return "lookup";
      case "cell": return (n.detail && n.detail.is_formula) ? "formula" : "source";
      case "number": case "text": case "bool": case "empty": case "name": case "range": return "source";
      default: return "calculation";
    }
  }

  function buildFlowGraph(treeRoot, opts) {
    let seq = 0;
    function make(n, isRoot) {
      const fn = {
        id: "f" + (seq++), kind: flowKind(n, isRoot), active: n.evaluated !== false && n.kind !== "skipped",
        excel: n.excel || n.kind, value: (n.evaluated === false ? "not evaluated" : n.value_repr),
        role: n.role, raw: n, kids: [], edgeLabel: "",
      };
      (n.children || []).forEach((c) => {
        const skipped = c.evaluated === false || c.kind === "skipped";
        if (opts.onlyTaken && skipped) {
          fn.kids.push({ id: "f" + (seq++), kind: "inactive", active: false,
            excel: clip(c.excel || "not taken", 34), value: "not taken", role: c.role, raw: c, kids: [], edgeLabel: "" });
          return;
        }
        if (opts.hideConst && !skipped && ["number", "text", "bool"].indexOf(c.kind) >= 0) return;
        const child = make(c, false);
        child.edgeLabel = child.active ? child.value : "";
        fn.kids.push(child);
      });
      return fn;
    }
    return make(treeRoot, true);
  }

  const F_NW = 176, F_NH = 54, F_XG = 30, F_YG = 92;
  function layoutFlowGraph(root) {
    let leaf = 0, maxD = 0;
    (function depth(n, d) { n._d = d; maxD = Math.max(maxD, d); n.kids.forEach((k) => depth(k, d + 1)); })(root, 0);
    (function place(n) {
      if (!n.kids.length) { n._gx = leaf++; return; }
      n.kids.forEach(place);
      n._gx = (n.kids[0]._gx + n.kids[n.kids.length - 1]._gx) / 2;
    })(root);
    const all = [];
    (function pos(n) {
      n._x = n._gx * (F_NW + F_XG);
      n._y = (maxD - n._d) * (F_NH + F_YG);   // result at the bottom, inputs at the top
      all.push(n); n.kids.forEach(pos);
    })(root);
    const xs = all.map((n) => n._x), ys = all.map((n) => n._y);
    return { all, minX: Math.min.apply(null, xs), maxX: Math.max.apply(null, xs),
             minY: Math.min.apply(null, ys), maxY: Math.max.apply(null, ys) };
  }

  function renderFlowSvg(root) {
    const L = layoutFlowGraph(root);
    const canvas = $("#flow-canvas"); canvas.innerHTML = "";
    const pad = 46;
    const vbX = L.minX - pad, vbY = L.minY - pad;
    const vbW = (L.maxX - L.minX) + F_NW + pad * 2, vbH = (L.maxY - L.minY) + F_NH + pad * 2;
    const svg = el("svg", { id: "flow-svg", viewBox: `${vbX} ${vbY} ${vbW} ${vbH}`, preserveAspectRatio: "xMidYMid meet" });
    svg.appendChild(el("defs", {}, [
      el("marker", { id: "fa", markerWidth: "9", markerHeight: "9", refX: "7", refY: "4", orient: "auto" },
        [el("path", { d: "M0,0 L8,4 L0,8 z", fill: "#8fa0b2" })]),
    ]));
    (function edges(n) {
      n.kids.forEach((k) => {
        const x1 = k._x + F_NW / 2, y1 = k._y + F_NH, x2 = n._x + F_NW / 2, y2 = n._y;
        const my = (y1 + y2) / 2;
        svg.appendChild(el("path", { class: "fn-edge" + (k.active ? "" : " inactive"),
          d: `M${x1},${y1} C${x1},${my} ${x2},${my} ${x2},${y2}`, "marker-end": "url(#fa)" }));
        if (k.edgeLabel) svg.appendChild(el("text", { class: "fn-edge-label", x: (x1 + x2) / 2 + 5, y: my - 2, text: clip(k.edgeLabel, 14) }));
        edges(k);
      });
    })(root);
    (function nodes(n) {
      const g = el("g", { class: `fn-box fn-${n.kind}` + (n.active ? "" : " inactive"),
        transform: `translate(${n._x},${n._y})`, onclick: (e) => { e.stopPropagation(); selectFlowNode(n, e.currentTarget); } });
      g.appendChild(el("rect", { width: F_NW, height: F_NH, rx: 9 }));
      g.appendChild(el("text", { class: "fn-title", x: 11, y: 17, text: clip(n.excel, 27) }));
      g.appendChild(el("text", { class: "fn-val", x: 11, y: 35, text: clip(n.value, 25) }));
      g.appendChild(el("text", { class: "fn-sub", x: 11, y: 48, text: (n.role && n.role !== "root" && n.role !== "arg") ? n.role : n.kind }));
      svg.appendChild(g);
      n.kids.forEach(nodes);
    })(root);
    canvas.appendChild(svg);
    wireFlowPanZoom(svg, [vbX, vbY, vbW, vbH]);
  }

  function wireFlowPanZoom(svg, base) {
    let vb = { x: base[0], y: base[1], w: base[2], h: base[3] };
    const apply = () => svg.setAttribute("viewBox", `${vb.x} ${vb.y} ${vb.w} ${vb.h}`);
    FLOW._fit = () => { vb = { x: base[0], y: base[1], w: base[2], h: base[3] }; apply(); };
    FLOW._zoom = (f, cx, cy) => {
      const r = svg.getBoundingClientRect();
      const px = vb.x + ((cx == null ? r.left + r.width / 2 : cx) - r.left) / r.width * vb.w;
      const py = vb.y + ((cy == null ? r.top + r.height / 2 : cy) - r.top) / r.height * vb.h;
      vb.w *= f; vb.h *= f; vb.x = px - (px - vb.x) * f; vb.y = py - (py - vb.y) * f; apply();
    };
    svg.addEventListener("wheel", (e) => { e.preventDefault(); FLOW._zoom(e.deltaY > 0 ? 1.12 : 0.9, e.clientX, e.clientY); }, { passive: false });
    let drag = null;
    svg.addEventListener("mousedown", (e) => { drag = { x: e.clientX, y: e.clientY }; $("#flow-canvas").classList.add("grabbing"); });
    window.addEventListener("mouseup", () => { drag = null; $("#flow-canvas").classList.remove("grabbing"); });
    window.addEventListener("mousemove", (e) => {
      if (!drag) return;
      const r = svg.getBoundingClientRect();
      vb.x -= (e.clientX - drag.x) * vb.w / r.width;
      vb.y -= (e.clientY - drag.y) * vb.h / r.height;
      drag = { x: e.clientX, y: e.clientY }; apply();
    });
  }

  function selectFlowNode(n, g) {
    FLOW.sel = n.id;
    $$("#flow-canvas .fn-box").forEach((x) => x.classList.remove("sel"));
    if (g) g.classList.add("sel");
    const raw = n.raw, d = raw.detail || {};
    const box = $("#flow-detail"); box.classList.remove("hidden"); box.innerHTML = "";
    box.appendChild(el("button", { class: "icon-btn fd-close", text: "✕", onclick: () => box.classList.add("hidden") }));
    box.appendChild(el("h4", { text: raw.excel || raw.kind }));
    box.appendChild(el("dl", { class: "kv" }, [
      el("dt", { text: "value" }), el("dd", { text: n.value }),
      el("dt", { text: "type" }), el("dd", { text: raw.dtype || n.kind }),
      raw.py ? el("dt", { text: "expr" }) : null, raw.py ? el("dd", { class: "mono", text: raw.py }) : null,
    ]));
    if (d.function && d.table_preview) {
      box.appendChild(el("p", { class: "muted small", text: `${d.function} key ${d.lookup_repr} · ${d.match_found ? "matched row " + (d.matched_row_index + 1) : "no match"}` }));
      const t = el("table", { class: "lk-table" });
      d.table_preview.forEach((r) => t.appendChild(el("tr", { class: r.matched ? "match" : "" }, (r.cells || []).map((c) => el("td", { text: fmt(c) })))));
      box.appendChild(t);
    }
    if (d.error) box.appendChild(el("pre", { text: d.error }));
    if (d.toggle) box.appendChild(el("p", { class: "muted small", text: `toggle “${d.toggle}” = ${d.toggle_value} — ${d.source || ""}` }));
    if (raw.kind === "cell" && d.is_formula && d.key) {
      box.appendChild(el("button", { class: "btn small", style: "margin-top:8px", text: "Re-root chart on " + d.key, onclick: () => showFlowchart(d.key) }));
    }
  }

  function buildAndRenderFlow(keepSel) {
    if (!FLOW || !FLOW.tree) return;
    if (!keepSel) FLOW.sel = null;
    const canvas = $("#flow-canvas");
    if (!FLOW.tree.root) { canvas.innerHTML = "<p class='muted' style='padding:24px'>This cell has no formula to chart.</p>"; return; }
    const opts = { hideConst: $("#flow-hide-const").checked, onlyTaken: $("#flow-only-taken").checked };
    renderFlowSvg(buildFlowGraph(FLOW.tree.root, opts));
  }

  async function showFlowchart(cell) {
    if (!S.summary || !S.summary.engine_build) { toast("Redeploy excel_deep_trace.py and restart the backend.", "err"); return; }
    try {
      const r = await call("/api/tree", { method: "POST", json: { session_id: S.sid, cell, max_depth: 24 } });
      FLOW = { cell, tree: r.tree, sel: null };
      show("flow");
      $("#flow-title").textContent = "Flowchart — " + r.tree.key;
      $("#flow-sub").textContent = r.tree.is_formula
        ? `${r.tree.key} = ${r.tree.value_repr}` + (r.tree.no_excel_value ? "  (evaluated by the tracer)" : "")
        : `${r.tree.key} is an input cell`;
      $("#flow-formula").textContent = r.tree.formula || "(input cell — nothing to chart)";
      renderFlowLegend();
      buildAndRenderFlow();
      if (FLOW._fit) FLOW._fit();
    } catch (e) { toast(e.message, "err"); }
  }

  // ================================================================ STEPS / DEBUGGER
  let STEPS = { data: null, i: 0, timer: null };

  function engineOutdated() {
    const caps = (S.summary && S.summary.engine_caps) || [];
    if (S.summary && S.summary.engine_build && (!caps.length || caps.indexOf("formula_transformation") >= 0)) return false;
    toast("The excel_deep_trace.py on the project libraries is out of date — redeploy python-lib/ and restart the backend.", "err");
    return true;
  }

  async function showSteps(cell) {
    if (engineOutdated()) return;
    try {
      const [tr, why] = await Promise.all([
        call("/api/transformation", { method: "POST", json: { session_id: S.sid, cell } }),
        call("/api/explain-plain", { method: "POST", json: { session_id: S.sid, cell } }).catch(() => null),
      ]);
      const d = tr.transformation;
      STEPS = { data: d, i: 0, timer: null };
      show("steps");
      $("#steps-title").textContent = "Transformation — " + d.key;
      $("#steps-sub").textContent = d.is_formula
        ? `${d.key} = ${d.final_repr}  ·  ${d.steps.length} steps`
        : `${d.key} is an input cell`;
      $("#steps-formula").textContent = d.formula || "(input cell)";
      $("#steps-why").textContent = (why && why.text) || "—";
      renderSteps();
      dbgSet(0);
    } catch (e) { toast(e.message, "err"); }
  }

  function renderSteps() {
    const box = $("#steps-list"); box.innerHTML = "";
    (STEPS.data.steps || []).forEach((s, idx) => {
      const card = el("div", { class: "step", "data-i": idx, id: "step-" + idx });
      card.appendChild(el("div", { class: "step-head" }, [
        el("span", { class: "step-n", text: s.n }),
        el("span", { class: "step-title", text: s.title }),
        el("span", { class: "step-tag " + s.kind, text: (s.kind || "").replace(/_/g, " ") }),
      ]));
      if (s.substitutions) {
        card.appendChild(el("div", { class: "step-subs" }, s.substitutions.map((u) =>
          el("span", { class: "chip", text: `${u.ref} = ${u.repr}` }))));
      }
      if (s.expr_in && s.result_repr != null) {
        card.appendChild(el("div", { class: "step-expr" }, [s.expr_in, "  →  ",
          el("span", { class: "step-result", text: s.result_repr })]));
      } else if (s.result_repr != null && s.kind !== "final" && s.kind !== "original") {
        card.appendChild(el("div", { class: "step-expr" }, [
          el("span", { class: "step-result", text: s.result_repr })]));
      }
      if (s.kind === "condition" && s.condition) {
        card.appendChild(el("div", { class: "step-expr", text: `${s.condition} is ${s.condition_repr} → ${(s.branch || "").toUpperCase()} branch` }));
      }
      const d = s.detail || {};
      if (d.rows && d.function) {   // VLOOKUP-style lookup table
        const t = el("table", { class: "step-lk" });
        (d.rows).forEach((r) => t.appendChild(el("tr", { class: r.matched ? "match" : "" },
          (r.cells || []).map((c) => el("td", { text: fmt(c) })))));
        card.appendChild(t);
      }
      if (d.matches || d.preview) {   // SUMIFS-style matching records
        const rows = (d.preview || d.matches || []).slice(0, 12);
        const t = el("table", { class: "step-lk" });
        t.appendChild(el("tr", {}, [el("th", { text: "row" }),
          ...(d.criteria_ranges || []).map((c) => el("th", { text: c })),
          d.agg_range ? el("th", { text: d.agg_range }) : null, el("th", { text: "" })]));
        rows.forEach((r) => t.appendChild(el("tr", { class: r.matched ? "match" : "" }, [
          el("td", { text: r.row }),
          ...((r.criteria || []).map((c) => el("td", { text: fmt(c) }))),
          d.agg_range ? el("td", { text: fmt(r.value) }) : null,
          el("td", { text: r.matched ? "✓" : "" }),
        ])));
        card.appendChild(t);
        card.appendChild(el("div", { class: "muted small", text: `${d.matched_rows} of ${d.scanned_rows} rows matched` }));
      }
      if (s.formula) card.appendChild(el("div", { class: "step-formula", text: s.formula }));
      box.appendChild(card);
    });
  }

  function dbgSet(i) {
    const n = STEPS.data.steps.length;
    STEPS.i = Math.max(0, Math.min(n - 1, i));
    $$("#steps-list .step").forEach((c, idx) => {
      c.classList.toggle("done", idx <= STEPS.i);
      c.classList.toggle("current", idx === STEPS.i);
    });
    const cur = $("#step-" + STEPS.i);
    if (cur) cur.scrollIntoView({ block: "center", behavior: "smooth" });
    $("#dbg-status").textContent = `Step ${STEPS.i + 1} of ${n} — ${STEPS.data.steps[STEPS.i].title}`;
    if (STEPS.i >= n - 1 && STEPS.timer) { clearInterval(STEPS.timer); STEPS.timer = null; $("#dbg-play").innerHTML = "&#9654;"; }
  }
  function dbgPlay() {
    if (STEPS.timer) { clearInterval(STEPS.timer); STEPS.timer = null; $("#dbg-play").innerHTML = "&#9654;"; return; }
    $("#dbg-play").innerHTML = "&#9208;";
    STEPS.timer = setInterval(() => dbgSet(STEPS.i + 1), 900);
  }
  function dbgDownload() {
    const blob = new Blob([JSON.stringify(STEPS.data, null, 2)], { type: "application/json" });
    const a = el("a", { href: URL.createObjectURL(blob), download: STEPS.data.key.replace(/[!:]/g, "_") + "_transformation.json" });
    document.body.appendChild(a); a.click(); a.remove();
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", boot);
  else boot();
})();
