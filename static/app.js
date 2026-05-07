/* Day 03. Budget vs Actual Tracker.
   Vanilla JS. CSRF. AbortController. localStorage restore. Click-to-filter pivots. */

(() => {
  "use strict";
  const $  = (sel, ctx = document) => ctx.querySelector(sel);
  const $$ = (sel, ctx = document) => Array.from(ctx.querySelectorAll(sel));

  // --- refs ------------------------------------------------------------
  const apiKeyInput  = $("#api-key");
  const skipAi       = $("#skip-ai");
  const modelSel     = $("#model");
  const analyseBtn   = $("#analyse-btn");
  const resetBtn     = $("#reset");
  const themeBtn     = $("#theme-toggle");
  const radios       = $$('input[name="src"]');
  const dropzone     = $("#upload-zone");
  const fileInput    = $("#file-input");
  const fileListEl   = $("#file-list");
  const bundledZone  = $("#bundled-zone");
  const compareZone  = $("#compare-zone");
  const compareFileA = $("#compare-file-a");
  const compareFileB = $("#compare-file-b");
  const compareNameA = $("#compare-name-a");
  const compareNameB = $("#compare-name-b");
  const compareCard  = $("#compare-card");
  const compareSummary = $("#compare-summary");
  const compareBody    = $("#compare-body");
  const compareCsvLink  = $("#compare-csv-link");
  const compareXlsxLink = $("#compare-xlsx-link");

  const empty   = $("#empty-state");
  const loading = $("#loading-state");
  const errorEl = $("#error-state");
  const errorMsg= $("#error-message");
  const loadMsg = $("#loading-message");
  const resultsEl = $("#results");
  const summaryEl = $("#results-summary");
  const exportEl  = $("#export-strip");
  const csvLink   = $("#csv-link");
  const xlsxLink  = $("#xlsx-link");
  const pptxLink  = $("#pptx-link");

  const commentaryCard = $("#commentary-card");
  const commentaryHeadline = $("#commentary-headline");
  const commentarySummary  = $("#commentary-summary");
  const commentaryAdverse  = $("#commentary-adverse");
  const commentaryFav      = $("#commentary-favourable");
  const commentaryActions  = $("#commentary-actions");
  const commentaryCost     = $("#commentary-cost");

  const pivotRow         = $("#pivot-row");
  const pivotCostCentre  = $("#pivot-cost-centre");
  const pivotCategory    = $("#pivot-category");
  const pivotPeriod      = $("#pivot-period");
  const pivotPeriodCard  = $("#pivot-period-card");
  const pivotParent      = $("#pivot-parent");
  const pivotParentCard  = $("#pivot-parent-card");
  const heatmapCard      = $("#heatmap-card");
  const heatmapBody      = $("#heatmap-body");
  const detailCard       = $("#detail-card");
  const detailBody       = $("#detail-body");
  const detailSearch     = $("#detail-search");
  const detailClear      = $("#detail-clear");
  const filterChip       = $("#filter-chip");

  const warningsCard = $("#warnings-card");
  const warningsList = $("#warnings-list");

  const statRows   = $("#stat-rows");
  const statBudget = $("#stat-budget");
  const statVarPct = $("#stat-var-pct");
  const statCost   = $("#stat-cost");

  const cost30dEl   = $("#cost-30d");
  const costTotalEl = $("#cost-total");
  const costRunsEl  = $("#cost-runs");
  const costClearBtn = $("#cost-clear");

  const historyOpen   = $("#history-open");
  const historyModal  = $("#history-modal");
  const historyBody   = $("#history-body");
  const historyRefresh = $("#history-refresh");
  const historySummaryLine = $("#history-summary-line");
  const shortcutsModal = $("#shortcuts-modal");

  // --- state -----------------------------------------------------------
  let queuedFile = null;
  let inFlight = null;
  let lastData = null;
  let activeFilter = null;            // {dim, key} or null
  let activePills = {                 // multi-select filters set by the pill bar
    rag: new Set(),                   // subset of {green,amber,red,favourable,na}
    line_type: new Set(),             // subset of {cost,revenue}
    direction: new Set(),             // subset of {adverse,favourable,onplan}
  };
  let detailSort = null;              // {key, dir: "asc"|"desc"} or null
  let notes = {};                     // {company: {sigKey: noteText}}
  const STORAGE_KEY = "day03.lastResult.v1";
  const NOTES_KEY = "day03.notes.v1";

  function loadNotes() {
    try {
      const raw = localStorage.getItem(NOTES_KEY);
      return raw ? JSON.parse(raw) : {};
    } catch (_) { return {}; }
  }
  function saveNotes() {
    try { localStorage.setItem(NOTES_KEY, JSON.stringify(notes)); } catch (_) {}
  }
  function noteSig(r) {
    return [r.period, r.cost_centre, r.category, r.line_type].join("||");
  }
  function getNote(company, r) {
    const bucket = notes[company] || {};
    const sig = noteSig(r);
    return bucket[sig] || r.notes || "";
  }
  function setNote(company, r, text) {
    if (!notes[company]) notes[company] = {};
    const sig = noteSig(r);
    if (text && text.trim()) notes[company][sig] = text.trim();
    else delete notes[company][sig];
    saveNotes();
  }
  notes = loadNotes();

  // --- CSRF + theme ----------------------------------------------------
  function readCookie(name) {
    return document.cookie.split(";").map(c => c.trim()).find(c => c.startsWith(name + "="))
      ?.slice(name.length + 1) || "";
  }
  function applyTheme(t) {
    document.documentElement.setAttribute("data-theme", t);
    try { localStorage.setItem("day03.theme", t); } catch (_) {}
  }
  (function initTheme() {
    let t;
    try { t = localStorage.getItem("day03.theme"); } catch (_) {}
    if (!t) t = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches
      ? "dark" : "light";
    applyTheme(t);
  })();
  themeBtn.addEventListener("click", () => {
    const cur = document.documentElement.getAttribute("data-theme") || "light";
    applyTheme(cur === "dark" ? "light" : "dark");
  });

  // --- source toggle ---------------------------------------------------
  function syncSourceUI() {
    const v = (radios.find(r => r.checked) || radios[0]).value;
    dropzone.hidden    = v !== "upload";
    bundledZone.hidden = v !== "bundled";
    if (compareZone) compareZone.hidden = v !== "compare";
  }
  radios.forEach(r => r.addEventListener("change", syncSourceUI));
  syncSourceUI();

  // --- compare file pickers --------------------------------------------
  if (compareFileA && compareNameA) {
    compareFileA.addEventListener("change", () => {
      const f = (compareFileA.files || [])[0];
      compareNameA.textContent = f ? `${f.name} (${humanBytes(f.size)})` : "No file picked";
    });
  }
  if (compareFileB && compareNameB) {
    compareFileB.addEventListener("change", () => {
      const f = (compareFileB.files || [])[0];
      compareNameB.textContent = f ? `${f.name} (${humanBytes(f.size)})` : "No file picked";
    });
  }

  // --- file picker + drag and drop -------------------------------------
  function setQueuedFile(file) {
    queuedFile = file;
    fileListEl.innerHTML = "";
    if (!file) return;
    const li = document.createElement("li");
    li.innerHTML = `<span>${escapeHtml(file.name)} <span style="color:var(--muted)">(${humanBytes(file.size)})</span></span>
                    <span class="remove-x" title="Remove">&times;</span>`;
    fileListEl.appendChild(li);
    li.querySelector(".remove-x").addEventListener("click", (e) => {
      e.stopPropagation();
      setQueuedFile(null);
    });
  }
  dropzone.addEventListener("click", (e) => {
    if (e.target.closest(".remove-x")) return;
    fileInput.click();
  });
  dropzone.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); fileInput.click(); }
  });
  fileInput.addEventListener("change", () => {
    const f = (fileInput.files || [])[0];
    if (f) setQueuedFile(f);
  });
  ["dragenter","dragover"].forEach(ev => dropzone.addEventListener(ev, (e) => {
    e.preventDefault(); dropzone.classList.add("drag-over");
  }));
  ["dragleave","drop"].forEach(ev => dropzone.addEventListener(ev, (e) => {
    e.preventDefault(); dropzone.classList.remove("drag-over");
  }));
  dropzone.addEventListener("drop", (e) => {
    if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]) {
      setQueuedFile(e.dataTransfer.files[0]);
    }
  });
  // Dropping onto the bundled-samples zone also auto-switches to upload mode.
  if (bundledZone) {
    ["dragenter","dragover"].forEach(ev => bundledZone.addEventListener(ev, (e) => {
      if (dragHasFiles(e)) { e.preventDefault(); bundledZone.classList.add("drag-over"); }
    }));
    ["dragleave","drop"].forEach(ev => bundledZone.addEventListener(ev, (e) => {
      e.preventDefault(); bundledZone.classList.remove("drag-over");
    }));
    bundledZone.addEventListener("drop", (e) => {
      const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (!f) return;
      const uploadRadio = radios.find(r => r.value === "upload");
      if (uploadRadio && !uploadRadio.checked) {
        uploadRadio.checked = true;
        syncSourceUI();
      }
      setQueuedFile(f);
    });
  }

  // --- page-wide drag overlay -----------------------------------------
  const pageDropOverlay = $("#page-drop-overlay");
  let dragCounter = 0;
  function dragHasFiles(e) {
    const t = e.dataTransfer;
    if (!t || !t.types) return false;
    for (const ty of t.types) if (ty === "Files" || ty === "application/x-moz-file") return true;
    return false;
  }
  window.addEventListener("dragenter", (e) => {
    if (!dragHasFiles(e)) return;
    e.preventDefault();
    dragCounter += 1;
    if (pageDropOverlay) pageDropOverlay.hidden = false;
  });
  window.addEventListener("dragover", (e) => {
    if (!dragHasFiles(e)) return;
    e.preventDefault();
    if (e.dataTransfer) e.dataTransfer.dropEffect = "copy";
  });
  window.addEventListener("dragleave", (e) => {
    if (!dragHasFiles(e)) return;
    dragCounter = Math.max(0, dragCounter - 1);
    if (dragCounter === 0 && pageDropOverlay) pageDropOverlay.hidden = true;
  });
  window.addEventListener("drop", (e) => {
    if (!dragHasFiles(e)) return;
    e.preventDefault();
    dragCounter = 0;
    if (pageDropOverlay) pageDropOverlay.hidden = true;
    const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (!f) return;
    if (!/\.(xlsx|csv)$/i.test(f.name)) {
      showError(`Only .xlsx or .csv files are accepted. Got: ${f.name}`);
      return;
    }
    const uploadRadio = radios.find(r => r.value === "upload");
    if (uploadRadio && !uploadRadio.checked) {
      uploadRadio.checked = true;
      syncSourceUI();
    }
    setQueuedFile(f);
    dropzone.scrollIntoView({ behavior: "smooth", block: "nearest" });
  });

  // --- states ----------------------------------------------------------
  function showState(name) {
    empty.hidden    = name !== "empty";
    loading.hidden  = name !== "loading";
    errorEl.hidden  = name !== "error";
    resultsEl.hidden= name !== "results";
    if (name === "error") setTimeout(() => errorEl.focus(), 0);
  }
  function showError(msg) { errorMsg.textContent = msg; showState("error"); }

  // --- analyse ---------------------------------------------------------
  analyseBtn.addEventListener("click", () => analyse().catch(e => showError(e.message)));
  resetBtn.addEventListener("click", () => {
    try { localStorage.removeItem(STORAGE_KEY); } catch (_) {}
    setQueuedFile(null);
    document.title = "Day 03. Budget vs Actual Tracker";
    resetBtn.hidden = true;
    activeFilter = null;
    activePills.rag.clear();
    activePills.line_type.clear();
    activePills.direction.clear();
    detailSort = null;
    lastData = null;
    showState("empty");
    setStats(null);
  });
  document.addEventListener("keydown", (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === "Enter") {
      e.preventDefault();
      if (!analyseBtn.disabled) analyseBtn.click();
    }
  });

  async function analyse() {
    const v = (radios.find(r => r.checked) || radios[0]).value;
    if (v === "compare") {
      return compareRun();
    }
    if (inFlight) inFlight.abort();
    const ac = new AbortController();
    inFlight = ac;

    const fd = new FormData();
    fd.append("skip_ai", skipAi.checked ? "true" : "false");
    if (apiKeyInput.value.trim()) fd.append("api_key", apiKeyInput.value.trim());
    if (modelSel.value)            fd.append("model", modelSel.value);

    if (v === "bundled") {
      const picked = $$('input[name="sample-pick"]:checked')[0];
      if (!picked) return showError("Pick a bundled sample first.");
      fd.append("use_samples", "true");
      fd.append("sample_id", picked.value);
    } else {
      if (!queuedFile) return showError("Pick a workbook first.");
      fd.append("use_samples", "false");
      fd.append("file", queuedFile, queuedFile.name);
    }

    analyseBtn.disabled = true;
    loadMsg.textContent = skipAi.checked
      ? "Computing variances (AI skipped)..."
      : `Computing variances and drafting commentary via ${modelSel.value || "Haiku 4.5"}...`;
    showState("loading");

    let resp, text;
    try {
      resp = await fetch("/api/analyse", {
        method: "POST", body: fd, signal: ac.signal,
        headers: { "X-CSRF-Token": readCookie("csrf_token") },
      });
      text = await resp.text();
    } catch (e) {
      analyseBtn.disabled = false;
      if (e.name === "AbortError") return;
      return showError(`Network error: ${e.message}. Is the server running? Run start.bat or 'python server.py'.`);
    } finally {
      inFlight = null;
    }
    analyseBtn.disabled = false;

    let data;
    try { data = JSON.parse(text); }
    catch {
      const ctype = resp.headers.get("content-type") || "unknown";
      return showError(`Server returned ${resp.status} (${ctype}). Restart with start.bat and try again.\nPreview: ${text.slice(0, 240)}`);
    }
    if (!resp.ok || data.error) return showError(data.error || `HTTP ${resp.status}`);

    render(data);
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify({ at: Date.now(), data })); } catch (_) {}
    showState("results");
    resetBtn.hidden = false;
  }

  // --- compare mode ----------------------------------------------------
  async function compareRun() {
    const fa = (compareFileA && compareFileA.files || [])[0];
    const fb = (compareFileB && compareFileB.files || [])[0];
    if (!fa || !fb) return showError("Pick two .xlsx workbooks first.");

    if (inFlight) inFlight.abort();
    const ac = new AbortController();
    inFlight = ac;

    const fd = new FormData();
    fd.append("file_a", fa, fa.name);
    fd.append("file_b", fb, fb.name);

    analyseBtn.disabled = true;
    loadMsg.textContent = "Comparing two workbooks (AI skipped)...";
    showState("loading");

    let resp, text;
    try {
      resp = await fetch("/api/compare", {
        method: "POST", body: fd, signal: ac.signal,
        headers: { "X-CSRF-Token": readCookie("csrf_token") },
      });
      text = await resp.text();
    } catch (e) {
      analyseBtn.disabled = false;
      if (e.name === "AbortError") return;
      return showError(`Network error: ${e.message}.`);
    } finally {
      inFlight = null;
    }
    analyseBtn.disabled = false;

    let data;
    try { data = JSON.parse(text); }
    catch { return showError(`Server returned ${resp.status}. ${text.slice(0, 200)}`); }
    if (!resp.ok || data.error) return showError(data.error || `HTTP ${resp.status}`);

    renderCompare(data);
    showState("results");
    resetBtn.hidden = false;
  }

  function renderCompare(c) {
    // Hide non-compare result panels for a clean view.
    if (compareCard) compareCard.hidden = false;
    [pivotRow, pivotParentCard, $("#commentary-card"), $("#heatmap-card"),
     detailCard, $("#warnings-card"), exportEl].forEach(el => { if (el) el.hidden = true; });
    summaryEl.innerHTML = "";

    if (compareCsvLink && c.csv_filename) {
      compareCsvLink.hidden = false;
      compareCsvLink.href = `/api/download/${encodeURIComponent(c.csv_filename)}`;
      compareCsvLink.textContent = `↓ ${c.csv_filename}`;
    } else if (compareCsvLink) {
      compareCsvLink.hidden = true;
    }
    if (compareXlsxLink && c.xlsx_filename) {
      compareXlsxLink.hidden = false;
      compareXlsxLink.href = `/api/download/${encodeURIComponent(c.xlsx_filename)}`;
      compareXlsxLink.textContent = `↓ ${c.xlsx_filename}`;
    } else if (compareXlsxLink) {
      compareXlsxLink.hidden = true;
    }

    const a = c.headline_a || {};
    const b = c.headline_b || {};
    compareSummary.innerHTML = `
      <div class="compare-kpi"><h4>A: ${escapeHtml(a.company || "-")}</h4>
        <p>${formatGBPSigned(a.total_variance)} <small style="color:var(--muted);font-size:.7rem;">(${escapeHtml(a.period_label || "")})</small></p>
      </div>
      <div class="compare-kpi"><h4>B: ${escapeHtml(b.company || "-")}</h4>
        <p>${formatGBPSigned(b.total_variance)} <small style="color:var(--muted);font-size:.7rem;">(${escapeHtml(b.period_label || "")})</small></p>
      </div>
      <div class="compare-kpi"><h4>Delta total variance</h4>
        <p class="${c.delta_total_variance > 0 ? "v-adverse" : (c.delta_total_variance < 0 ? "v-favourable" : "")}">${formatGBPSigned(c.delta_total_variance)}</p>
      </div>
      <div class="compare-kpi"><h4>Movers</h4><p>${c.n_movers || 0}</p></div>
      <div class="compare-kpi"><h4>Common rows</h4><p>${c.n_common || 0}</p></div>
      <div class="compare-kpi"><h4>Added in B</h4><p>${c.n_added || 0}</p></div>
      <div class="compare-kpi"><h4>Removed from A</h4><p>${c.n_removed || 0}</p></div>
    `;

    const rows = (c.rows || []).slice(0, 200);
    if (!rows.length) {
      compareBody.innerHTML = `<p class="muted">No rows.</p>`;
      return;
    }
    compareBody.innerHTML = `
      <table class="compare-table">
        <thead>
          <tr>
            <th>Status</th><th>Period</th><th>Cost centre</th><th>Category</th><th>Type</th>
            <th>Var A</th><th>Var B</th><th>&Delta;</th><th>RAG A &rarr; B</th>
          </tr>
        </thead>
        <tbody>
          ${rows.map(r => {
            const deltaClass = (r.delta_variance != null && r.delta_variance > 0) ? "v-adverse"
                             : (r.delta_variance != null && r.delta_variance < 0 ? "v-favourable" : "");
            const ragSpan = (rag) => rag ? `<span class="rag-pill rag-${escapeAttr(rag)}">${rag.toUpperCase()}</span>` : "-";
            return `
              <tr>
                <td><span class="compare-status cs-${escapeAttr(r.status)}">${escapeHtml(r.status)}</span></td>
                <td>${escapeHtml(r.period)}</td>
                <td>${escapeHtml(r.cost_centre)}</td>
                <td>${escapeHtml(r.category)}</td>
                <td>${escapeHtml(r.line_type)}</td>
                <td class="num">${r.variance_a == null ? "-" : formatGBPSigned(r.variance_a)}</td>
                <td class="num">${r.variance_b == null ? "-" : formatGBPSigned(r.variance_b)}</td>
                <td class="num ${deltaClass}">${r.delta_variance == null ? "-" : formatGBPSigned(r.delta_variance)}</td>
                <td>${ragSpan(r.rag_a)} ${r.rag_a && r.rag_b ? "&rarr;" : ""} ${ragSpan(r.rag_b)}</td>
              </tr>
            `;
          }).join("")}
        </tbody>
      </table>
      ${c.rows.length > 200 ? `<p class="muted">Showing the first 200 of ${c.rows.length} rows.</p>` : ""}
    `;
  }

  // --- render ----------------------------------------------------------
  function render(data) {
    lastData = data;
    activeFilter = null;
    activePills.rag.clear();
    activePills.line_type.clear();
    activePills.direction.clear();
    detailSort = null;
    const company = data.metadata && data.metadata.company || "Budget";
    document.title = `${company}. Day 03 variance`;

    setStats(data);
    summaryEl.innerHTML = renderKpis(data);
    renderCommentary(data.commentary);
    renderPivots(data);
    renderHeatmap(data);
    renderDetail(data);
    renderWarnings(data.warnings || []);

    if (data.csv_filename || data.xlsx_filename) {
      exportEl.hidden = false;
      csvLink.href  = data.csv_filename  ? `/api/download/${encodeURIComponent(data.csv_filename)}`  : "#";
      xlsxLink.href = data.xlsx_filename ? `/api/download/${encodeURIComponent(data.xlsx_filename)}` : "#";
      csvLink.textContent  = `↓ ${data.csv_filename}`;
      xlsxLink.textContent = `↓ ${data.xlsx_filename}`;
      if (pptxLink) {
        if (data.pptx_filename) {
          pptxLink.hidden = false;
          pptxLink.href = `/api/download/${encodeURIComponent(data.pptx_filename)}`;
          pptxLink.textContent = `↓ ${data.pptx_filename}`;
        } else {
          pptxLink.hidden = true;
        }
      }
    } else {
      exportEl.hidden = true;
    }
    if (data.cost_log) renderCostLog(data.cost_log);
  }

  function setStats(data) {
    if (!data) {
      statRows.textContent = "0";
      statBudget.textContent = "0";
      statVarPct.textContent = "0";
      statCost.textContent = "$0.0000";
      return;
    }
    const h = data.headline || {};
    statRows.textContent   = String(h.row_count || 0);
    statBudget.textContent = formatGBP(h.total_budget);
    statVarPct.textContent = formatPctSigned(h.total_variance_pct);
    const cost = (data.commentary && typeof data.commentary.cost_usd === "number")
      ? data.commentary.cost_usd
      : (typeof data.total_cost_usd === "number" ? data.total_cost_usd : 0);
    statCost.textContent = `$${cost.toFixed(4)}`;
  }

  function renderKpis(data) {
    const h = data.headline || {};
    const c = h.rag_counts || {};
    const variance = h.total_variance || 0;
    const variance_pct = h.total_variance_pct;
    const variance_class = (variance > 0) ? "kpi-red" : (variance < 0 ? "kpi-amber" : "kpi-navy");
    const anomalyTile = (data.anomaly_count > 0)
      ? `<div class="kpi kpi-red"><h2>${data.anomaly_count}</h2><p>Anomalies (z &ge; 2)</p></div>`
      : "";
    return `
      <div class="kpi kpi-navy"><h2>${formatGBP(h.total_budget)}</h2><p>Total budget</p></div>
      <div class="kpi kpi-navy"><h2>${formatGBP(h.total_actual)}</h2><p>Total actual</p></div>
      <div class="kpi ${variance_class}"><h2>${formatGBPSigned(variance)}</h2><p>Variance ${formatPctSigned(variance_pct)}</p></div>
      <div class="kpi kpi-red"><h2>${c.red || 0}</h2><p>Red lines</p></div>
      <div class="kpi kpi-amber"><h2>${c.amber || 0}</h2><p>Amber lines</p></div>
      <div class="kpi kpi-green"><h2>${c.green || 0}</h2><p>Green lines</p></div>
      <div class="kpi kpi-blue"><h2>${c.favourable || 0}</h2><p>Favourable</p></div>
      ${anomalyTile}
    `;
  }

  function renderCommentary(c) {
    if (!c || c.skipped) {
      commentaryCard.hidden = true;
      return;
    }
    if (c.error) {
      commentaryCard.hidden = false;
      commentaryHeadline.textContent = "AI commentary unavailable";
      commentarySummary.textContent  = c.error;
      commentaryAdverse.innerHTML = "";
      commentaryFav.innerHTML = "";
      commentaryActions.innerHTML = "";
      commentaryCost.textContent = "";
      return;
    }
    commentaryCard.hidden = false;
    commentaryHeadline.textContent = c.headline || "";
    commentarySummary.textContent  = c.summary  || "";
    commentaryAdverse.innerHTML = (c.adverse_drivers || []).map(d => `<li>${escapeHtml(d)}</li>`).join("");
    commentaryFav.innerHTML     = (c.favourable_drivers || []).map(d => `<li>${escapeHtml(d)}</li>`).join("");
    commentaryActions.innerHTML = (c.actions || []).map(d => `<li>${escapeHtml(d)}</li>`).join("");
    if (c.cost_usd) {
      commentaryCost.textContent = `$${(c.cost_usd).toFixed(4)} . ${c.input_tokens} in / ${c.output_tokens} out . ${c.model}`;
    } else {
      commentaryCost.textContent = "";
    }
  }

  function renderPivots(data) {
    pivotRow.hidden = false;
    const periods = (data.by_period || []).map(p => p.key);
    const sparks = buildSparklines(data, periods);
    pivotCostCentre.innerHTML = renderPivot(data.by_cost_centre || [], "Cost centre", "cost_centre", sparks.cost_centre, periods);
    pivotCategory.innerHTML   = renderPivot(data.by_category    || [], "Category",    "category",    sparks.category,    periods);
    if (periods.length > 1) {
      pivotPeriodCard.hidden = false;
      pivotPeriod.innerHTML = renderPivot(data.by_period, "Period", "period", null, periods);
    } else {
      pivotPeriodCard.hidden = true;
    }
    if ((data.by_parent || []).length) {
      pivotParentCard.hidden = false;
      pivotParent.innerHTML = renderPivot(data.by_parent, "Parent", "parent", null, periods);
    } else {
      pivotParentCard.hidden = true;
    }
    // Click-to-filter wiring (mouse + keyboard) across the parent panel + pivot row.
    const allTables = [
      ...(pivotParentCard ? pivotParentCard.querySelectorAll(".pivot-table tr[data-key]") : []),
      ...pivotRow.querySelectorAll(".pivot-table tr[data-key]"),
    ];
    allTables.forEach(tr => {
      tr.tabIndex = 0;
      const toggle = () => {
        const dim = tr.closest(".pivot").dataset.dim;
        const key = tr.dataset.key;
        if (activeFilter && activeFilter.dim === dim && activeFilter.key === key) {
          activeFilter = null;
        } else {
          activeFilter = { dim, key };
        }
        applyFilter();
      };
      tr.addEventListener("click", toggle);
      tr.addEventListener("keydown", (e) => {
        if (e.key === "Enter" || e.key === " ") { e.preventDefault(); toggle(); }
      });
    });
    // Highlight active row across pivots and the parent panel.
    const activeHighlight = (root) => {
      if (!root) return;
      root.querySelectorAll(".pivot-table tr").forEach(tr => {
        tr.classList.toggle("active",
          activeFilter
          && tr.dataset.dim === activeFilter.dim
          && tr.dataset.key === activeFilter.key);
      });
    };
    activeHighlight(pivotParentCard);
    activeHighlight(pivotRow);
  }

  function renderPivot(rows, keyLabel, dim, sparkByKey, periods) {
    if (!rows.length) return `<p class="muted">No rows.</p>`;
    const showSpark = sparkByKey && periods && periods.length > 1;
    const sparkHead = showSpark ? `<th>Trend</th>` : "";
    return `
      <table class="pivot-table">
        <thead>
          <tr><th>${escapeHtml(keyLabel)}</th><th>Budget</th><th>Actual</th><th>Variance</th><th>%</th>${sparkHead}<th>Red</th><th>Amber</th><th>Green</th><th>Fav</th></tr>
        </thead>
        <tbody>
          ${rows.map(a => {
            const c = a.rag_counts || {};
            const varianceClass = a.variance > 0 ? "v-adverse" : (a.variance < 0 ? "v-favourable" : "");
            const sparkCell = showSpark
              ? `<td class="spark-cell">${sparklineSvg(sparkByKey[a.key] || [])}</td>`
              : "";
            return `
              <tr data-key="${escapeAttr(a.key)}" data-dim="${escapeAttr(dim)}">
                <td class="key">${escapeHtml(a.key)}</td>
                <td class="num">${formatGBP(a.budget)}</td>
                <td class="num">${formatGBP(a.actual)}</td>
                <td class="num ${varianceClass}">${formatGBPSigned(a.variance)}</td>
                <td class="num">${formatPctSigned(a.variance_pct)}</td>
                ${sparkCell}
                <td class="num">${c.red || 0}</td>
                <td class="num">${c.amber || 0}</td>
                <td class="num">${c.green || 0}</td>
                <td class="num">${c.favourable || 0}</td>
              </tr>
            `;
          }).join("")}
        </tbody>
      </table>
    `;
  }

  // Build a {dim: {key: [variance per period]}} structure ordered by periods.
  function buildSparklines(data, periods) {
    const out = { cost_centre: {}, category: {} };
    if (periods.length <= 1) return out;
    const pIndex = new Map(periods.map((p, i) => [p, i]));
    const fill = (dim) => {
      const byKey = {};
      for (const r of data.rows || []) {
        const key = r[dim];
        if (!byKey[key]) byKey[key] = new Array(periods.length).fill(0);
        const i = pIndex.get(r.period);
        if (i !== undefined) byKey[key][i] += (r.variance || 0);
      }
      return byKey;
    };
    out.cost_centre = fill("cost_centre");
    out.category    = fill("category");
    return out;
  }

  function sparklineSvg(values) {
    const W = 80, H = 22, padX = 2, padY = 3;
    if (!values || values.length === 0) return "";
    if (values.length === 1) {
      const v = values[0];
      const cls = v > 0 ? "spark-bar-pos" : (v < 0 ? "spark-bar-neg" : "");
      return `<svg viewBox="0 0 ${W} ${H}"><rect x="${W/2-2}" y="${H/2-3}" width="4" height="6" class="${cls}"/></svg>`;
    }
    const max = Math.max(...values.map(Math.abs)) || 1;
    const xStep = (W - padX * 2) / (values.length - 1);
    const midY = H / 2;
    const yScale = (H / 2) - padY;
    const points = values.map((v, i) => {
      const x = padX + i * xStep;
      const y = midY - (v / max) * yScale;
      return `${x.toFixed(1)},${y.toFixed(1)}`;
    }).join(" ");
    const dots = values.map((v, i) => {
      const x = padX + i * xStep;
      const y = midY - (v / max) * yScale;
      const cls = v > 0 ? "spark-dot-pos" : (v < 0 ? "spark-dot-neg" : "");
      return `<circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="1.4" class="${cls}"/>`;
    }).join("");
    return `<svg viewBox="0 0 ${W} ${H}"><line x1="${padX}" y1="${midY}" x2="${W-padX}" y2="${midY}" class="spark-zero"/><polyline points="${points}" class="spark-line"/>${dots}</svg>`;
  }

  // Worst RAG in a bucket wins. Order from most attention-grabbing to least.
  const RAG_PRIORITY = ["red", "amber", "favourable", "green", "na"];
  function worstRag(rags) {
    for (const band of RAG_PRIORITY) {
      if (rags.includes(band)) return band;
    }
    return null;
  }

  function renderHeatmap(data) {
    const rows = data.rows || [];
    if (!rows.length) { heatmapCard.hidden = true; return; }

    const ccs = [];
    const cats = [];
    const ccsSeen = new Set();
    const catsSeen = new Set();
    for (const r of rows) {
      if (!ccsSeen.has(r.cost_centre))   { ccsSeen.add(r.cost_centre);   ccs.push(r.cost_centre); }
      if (!catsSeen.has(r.category))     { catsSeen.add(r.category);     cats.push(r.category); }
    }

    // bucket key: "<cost_centre>||<category>"
    const buckets = new Map();
    for (const r of rows) {
      const key = `${r.cost_centre}||${r.category}`;
      let b = buckets.get(key);
      if (!b) { b = { rows: [], totalVar: 0 }; buckets.set(key, b); }
      b.rows.push(r);
      b.totalVar += (r.variance || 0);
    }

    if (ccs.length > 18 || cats.length > 18) {
      // Avoid an unreadable grid; skip. Pivots already cover this.
      heatmapCard.hidden = true;
      return;
    }
    heatmapCard.hidden = false;

    // Build the grid as a CSS-grid layout. First column is row headers, rest are categories.
    const cols = `220px repeat(${cats.length}, minmax(80px, 1fr))`;
    let html = `<div class="heatmap-grid" style="grid-template-columns:${cols}">`;
    html += `<div class="h-corner"></div>`;
    for (const cat of cats) {
      html += `<div class="h-col-head">${escapeHtml(cat)}</div>`;
    }
    for (const cc of ccs) {
      html += `<div class="h-row-head">${escapeHtml(cc)}</div>`;
      for (const cat of cats) {
        const b = buckets.get(`${cc}||${cat}`);
        if (!b) {
          html += `<div class="h-cell empty" title="No data">.</div>`;
          continue;
        }
        const rag = worstRag(b.rows.map(r => r.rag));
        const tooltip = `${cc} / ${cat}: ${b.rows.length} row${b.rows.length === 1 ? "" : "s"}, ` +
                        `variance ${formatGBPSigned(b.totalVar)}`;
        html += `<div class="h-cell rag-${escapeAttr(rag)}" data-cc="${escapeAttr(cc)}" data-cat="${escapeAttr(cat)}" title="${escapeAttr(tooltip)}">${formatGBPSigned(b.totalVar)}</div>`;
      }
    }
    html += `</div>`;
    heatmapBody.innerHTML = html;

    // Click a cell to filter the Detail table on that cost centre, then scroll to it.
    heatmapBody.querySelectorAll(".h-cell:not(.empty)").forEach(cell => {
      cell.addEventListener("click", () => {
        const cc = cell.dataset.cc;
        activeFilter = { dim: "cost_centre", key: cc };
        applyFilter();
        detailCard.scrollIntoView({ behavior: "smooth", block: "start" });
      });
    });
  }

  function renderDetail(data) {
    detailCard.hidden = false;
    applyFilter();
    detailSearch.value = "";
    detailSearch.removeEventListener("input", _onSearch);
    detailSearch.addEventListener("input", _onSearch);
    detailClear.removeEventListener("click", _onClearFilter);
    detailClear.addEventListener("click", _onClearFilter);
  }

  function _onSearch() {
    applyFilter();
  }
  function _onClearFilter() {
    activeFilter = null;
    activePills.rag.clear();
    activePills.line_type.clear();
    activePills.direction.clear();
    detailSort = null;
    detailSearch.value = "";
    applyFilter();
  }

  // ---- pill bar -----------------------------------------------------
  const PILL_GROUPS = [
    {
      label: "RAG",
      key: "rag",
      options: [
        { value: "red",        label: "Red" },
        { value: "amber",      label: "Amber" },
        { value: "green",      label: "Green" },
        { value: "favourable", label: "Favourable" },
        { value: "na",         label: "n/a" },
      ],
    },
    {
      label: "Line",
      key: "line_type",
      options: [
        { value: "cost",    label: "Cost" },
        { value: "revenue", label: "Revenue" },
      ],
    },
    {
      label: "Direction",
      key: "direction",
      options: [
        { value: "adverse",    label: "Adverse" },
        { value: "favourable", label: "Favourable" },
        { value: "onplan",     label: "On plan" },
      ],
    },
  ];

  function renderPillBar() {
    const host = $("#detail-pills");
    if (!host) return;
    if (!lastData || !(lastData.rows || []).length) {
      host.innerHTML = "";
      return;
    }
    const rows = lastData.rows;
    const counts = {
      rag: {}, line_type: {}, direction: { adverse: 0, favourable: 0, onplan: 0 },
    };
    for (const r of rows) {
      counts.rag[r.rag] = (counts.rag[r.rag] || 0) + 1;
      counts.line_type[r.line_type] = (counts.line_type[r.line_type] || 0) + 1;
      const dir = r.is_adverse ? "adverse" : (r.rag === "favourable" ? "favourable" : "onplan");
      counts.direction[dir] = (counts.direction[dir] || 0) + 1;
    }

    host.innerHTML = PILL_GROUPS.map(group => `
      <div class="pill-group">
        <span class="pill-label">${escapeHtml(group.label)}</span>
        ${group.options.map(opt => {
          const count = counts[group.key][opt.value] || 0;
          const pressed = activePills[group.key].has(opt.value);
          return `<button type="button"
                          class="filter-pill"
                          data-group="${escapeAttr(group.key)}"
                          data-value="${escapeAttr(opt.value)}"
                          aria-pressed="${pressed ? "true" : "false"}"
                          ${count === 0 ? "disabled" : ""}>
                    ${escapeHtml(opt.label)} <span class="pill-count">${count}</span>
                  </button>`;
        }).join("")}
      </div>
    `).join("");

    host.querySelectorAll(".filter-pill").forEach(btn => {
      btn.addEventListener("click", () => {
        const group = btn.dataset.group;
        const value = btn.dataset.value;
        if (activePills[group].has(value)) activePills[group].delete(value);
        else activePills[group].add(value);
        applyFilter();
      });
    });
  }

  function renderRagLegend() {
    const host = $("#rag-legend");
    if (!host) return;
    host.innerHTML = `
      <strong>RAG legend</strong>
      <span class="rag-pill rag-red">Red</span>
      <span class="rag-pill rag-amber">Amber</span>
      <span class="rag-pill rag-green">Green</span>
      <span class="rag-pill rag-favourable">Favourable</span>
      <span class="rag-pill rag-na">n/a</span>
    `;
  }

  function applyFilter() {
    if (!lastData) return;
    let rows = lastData.rows || [];
    const anyPill = activePills.rag.size || activePills.line_type.size || activePills.direction.size;

    if (activeFilter) {
      rows = rows.filter(r => r[activeFilter.dim] === activeFilter.key);
      filterChip.hidden = false;
      filterChip.textContent = `${activeFilter.dim.replace("_", " ")}: ${activeFilter.key}`;
    } else {
      filterChip.hidden = true;
    }
    if (activePills.rag.size) {
      rows = rows.filter(r => activePills.rag.has(r.rag));
    }
    if (activePills.line_type.size) {
      rows = rows.filter(r => activePills.line_type.has(r.line_type));
    }
    if (activePills.direction.size) {
      rows = rows.filter(r => {
        const dir = r.is_adverse ? "adverse" : (r.rag === "favourable" ? "favourable" : "onplan");
        return activePills.direction.has(dir);
      });
    }
    const term = (detailSearch.value || "").trim().toLowerCase();
    if (term) {
      rows = rows.filter(r =>
        r.cost_centre.toLowerCase().includes(term)
        || r.category.toLowerCase().includes(term)
        || r.period.toLowerCase().includes(term)
        || (r.notes || "").toLowerCase().includes(term)
      );
    }
    if (detailSort) {
      const { key, dir } = detailSort;
      const sign = dir === "asc" ? 1 : -1;
      rows = [...rows].sort((a, b) => {
        const va = a[key], vb = b[key];
        if (va == null && vb == null) return 0;
        if (va == null) return 1;
        if (vb == null) return -1;
        if (typeof va === "number" && typeof vb === "number") return sign * (va - vb);
        return sign * String(va).localeCompare(String(vb));
      });
    }
    detailClear.hidden = !(activeFilter || anyPill || term || detailSort);
    renderPillBar();
    renderRagLegend();
    detailBody.innerHTML = renderDetailTable(rows);
    wireNoteInputs(rows);
    // highlight active row in pivots + parent panel
    const tables = [
      ...(pivotParentCard ? pivotParentCard.querySelectorAll(".pivot-table tr") : []),
      ...pivotRow.querySelectorAll(".pivot-table tr"),
    ];
    tables.forEach(tr => {
      tr.classList.toggle("active",
        activeFilter
        && tr.dataset.dim === activeFilter.dim
        && tr.dataset.key === activeFilter.key);
    });
  }

  function wireNoteInputs(rows) {
    if (!lastData) return;
    const company = (lastData.metadata && lastData.metadata.company) || "_";
    detailBody.querySelectorAll(".note-input").forEach(inp => {
      const idx = parseInt(inp.dataset.rowIdx, 10);
      const row = rows[idx];
      if (!row) return;
      const commit = () => {
        setNote(company, row, inp.value);
        inp.classList.toggle("has-note", !!inp.value.trim());
      };
      inp.addEventListener("change", commit);
      inp.addEventListener("blur", commit);
      inp.addEventListener("keydown", (e) => {
        if (e.key === "Enter") { e.preventDefault(); inp.blur(); }
      });
    });
    // Sortable column headers.
    detailBody.querySelectorAll("th.sortable").forEach(th => {
      th.addEventListener("click", () => {
        const key = th.dataset.sortKey;
        if (!key) return;
        if (detailSort && detailSort.key === key) {
          detailSort.dir = detailSort.dir === "asc" ? "desc" : "asc";
        } else {
          detailSort = { key, dir: "asc" };
        }
        applyFilter();
      });
    });
  }

  function renderDetailTable(rows) {
    if (!rows.length) {
      return `<p class="muted">No rows match the current filter.</p>`;
    }
    const company = (lastData && lastData.metadata && lastData.metadata.company) || "_";
    const headers = [
      ["period",       "Period"],
      ["cost_centre",  "Cost centre"],
      ["category",     "Category"],
      ["line_type",    "Type"],
      ["budget",       "Budget"],
      ["actual",       "Actual"],
      ["variance",     "Variance"],
      ["variance_pct", "%"],
      ["rag",          "RAG"],
      ["z_score",      "Z"],
      [null,           "Note"],   // not sortable
    ];
    const arrow = (key) => {
      if (!detailSort || detailSort.key !== key) return `<span class="sort-arrow">&#9651;</span>`;
      return detailSort.dir === "asc"
        ? `<span class="sort-arrow active">&#9650;</span>`
        : `<span class="sort-arrow active">&#9660;</span>`;
    };
    const headerHtml = headers.map(([key, label]) => {
      if (!key) return `<th>${escapeHtml(label)}</th>`;
      return `<th class="sortable" data-sort-key="${escapeAttr(key)}">${escapeHtml(label)} ${arrow(key)}</th>`;
    }).join("");
    return `
      <table class="detail-table">
        <thead>
          <tr>${headerHtml}</tr>
        </thead>
        <tbody>
          ${rows.map((r, i) => {
            const varianceClass = r.is_adverse ? "v-adverse" : (r.rag === "favourable" ? "v-favourable" : "");
            const noteText = getNote(company, r);
            const sig = noteSig(r);
            const z = r.z_score;
            const zClass = (z != null && Math.abs(z) >= 2) ? "z-anomaly" : (z != null && Math.abs(z) >= 1.5 ? "z-watch" : "");
            const zText = (z != null) ? z.toFixed(1) : "-";
            const zTip = (z != null)
              ? `${(r.variance_pct * 100).toFixed(1)}% vs historical mean ${(r.z_mean * 100).toFixed(1)}% (sigma ${(r.z_stdev * 100).toFixed(1)}%, n=${r.z_n})`
              : "needs at least 3 historical observations for this key";
            return `
              <tr data-sig="${escapeAttr(sig)}">
                <td>${escapeHtml(r.period)}</td>
                <td>${escapeHtml(r.cost_centre)}</td>
                <td>${escapeHtml(r.category)}</td>
                <td>${escapeHtml(r.line_type)}</td>
                <td class="num">${formatGBP(r.budget)}</td>
                <td class="num">${formatGBP(r.actual)}</td>
                <td class="num ${varianceClass}">${formatGBPSigned(r.variance)}</td>
                <td class="num">${formatPctSigned(r.variance_pct)}</td>
                <td><span class="rag-pill rag-${escapeAttr(r.rag)}">${r.rag.toUpperCase()}</span></td>
                <td class="num ${zClass}" title="${escapeHtml(zTip)}">${zText}</td>
                <td class="note-cell">
                  <input type="text"
                         class="note-input ${noteText ? "has-note" : ""}"
                         placeholder="add note..."
                         value="${escapeHtml(noteText)}"
                         data-row-idx="${i}">
                </td>
              </tr>
            `;
          }).join("")}
        </tbody>
      </table>
    `;
  }

  function renderWarnings(warnings) {
    if (!warnings.length) {
      warningsCard.hidden = true;
      return;
    }
    warningsCard.hidden = false;
    warningsList.innerHTML = warnings.map(w => `<li>${escapeHtml(w)}</li>`).join("");
  }

  // --- formatters ------------------------------------------------------
  function formatGBP(v) {
    if (v == null) return "n/a";
    return `£${Math.abs(v).toLocaleString(undefined, {minimumFractionDigits: 0, maximumFractionDigits: 0})}`;
  }
  function formatGBPSigned(v) {
    if (v == null) return "n/a";
    const sign = v < 0 ? "-" : (v > 0 ? "+" : "");
    return `${sign}£${Math.abs(v).toLocaleString(undefined, {minimumFractionDigits: 0, maximumFractionDigits: 0})}`;
  }
  function formatPctSigned(v) {
    if (v == null) return "n/a";
    const sign = v < 0 ? "" : "+";
    return `${sign}${(v * 100).toFixed(1)}%`;
  }
  function humanBytes(n) {
    if (n < 1024) return `${n} B`;
    if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
    return `${(n / 1024 / 1024).toFixed(2)} MB`;
  }

  // --- cost log dock card ----------------------------------------------
  function renderCostLog(summary) {
    if (!summary || !cost30dEl) return;
    cost30dEl.textContent   = `$${(summary.cost_usd_30d || 0).toFixed(4)}`;
    costTotalEl.textContent = `$${(summary.cost_usd_total || 0).toFixed(4)}`;
    costRunsEl.textContent  = String(summary.runs || 0);
  }

  async function fetchCostLog() {
    try {
      const r = await fetch("/api/cost-log");
      if (!r.ok) return;
      const d = await r.json();
      renderCostLog(d.summary);
    } catch (_) {}
  }

  if (costClearBtn) {
    costClearBtn.addEventListener("click", async () => {
      if (!confirm("Clear the cost log? Past run history will be lost.")) return;
      try {
        const r = await fetch("/api/cost-log/clear", {
          method: "POST",
          headers: { "X-CSRF-Token": readCookie("csrf_token") },
        });
        if (!r.ok) throw new Error(`HTTP ${r.status}`);
        const d = await r.json();
        renderCostLog(d.summary);
      } catch (e) {
        alert(`Could not clear: ${e.message}`);
      }
    });
  }

  // Prime the dock cost figures from /api/status on load. Quiet on failure.
  fetch("/api/status").then(r => r.ok ? r.json() : null).then(d => {
    if (d && d.cost_log) renderCostLog(d.cost_log);
  }).catch(() => {});

  // --- modals: history + shortcuts ------------------------------------
  function openModal(modal) {
    if (!modal) return;
    modal.hidden = false;
    modal.setAttribute("aria-hidden", "false");
  }
  function closeModal(modal) {
    if (!modal) return;
    modal.hidden = true;
    modal.setAttribute("aria-hidden", "true");
  }
  document.addEventListener("click", (e) => {
    const target = e.target.closest("[data-close='1']");
    if (!target) return;
    const modal = target.closest(".modal");
    if (modal) closeModal(modal);
  });

  // --- run history -----------------------------------------------------
  async function fetchHistory() {
    try {
      const r = await fetch("/api/runs");
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      const data = await r.json();
      renderHistory(data);
    } catch (e) {
      historyBody.innerHTML = `<p class="muted">Could not load history: ${escapeHtml(e.message)}</p>`;
    }
  }

  function renderHistory(data) {
    const entries = data.entries || [];
    if (historySummaryLine) {
      const s = data.summary || {};
      historySummaryLine.textContent = `${entries.length} runs . spent $${(s.cost_usd_total || 0).toFixed(4)} all time`;
    }
    if (!entries.length) {
      historyBody.innerHTML = `<p class="muted">No runs yet. Analyse a workbook to start a history.</p>`;
      return;
    }
    // Mini sparkline of variance % across runs (oldest -> newest, so reverse).
    const ordered = [...entries].reverse();
    const sparkValues = ordered.map(e => (e.total_variance_pct == null ? 0 : e.total_variance_pct));
    const sparkSvg = sparklineSvg(sparkValues);

    historyBody.innerHTML = `
      <div class="history-spark" style="margin-bottom:.8rem;">
        <span class="muted" style="font-family:var(--font-mono);font-size:.7rem;letter-spacing:.1em;text-transform:uppercase;">Variance % across runs</span>
        <div style="margin-top:.2rem;">${sparkSvg}</div>
      </div>
      <table class="history-table">
        <thead>
          <tr>
            <th>When</th><th>Company</th><th>Period</th><th>Source</th>
            <th>Rows</th><th>Variance</th><th>Var %</th><th>Red</th><th>Cost</th><th></th>
          </tr>
        </thead>
        <tbody>
          ${entries.map(e => {
            const ts = (e.ts || "").replace("T", " ").replace("Z", "");
            const cached = e.cached;
            const rowClass = cached ? "history-row cached" : "history-row evicted";
            const reopenBtn = cached
              ? `<button class="btn-ghost btn-tight history-reopen" data-run-id="${escapeAttr(e.id)}" type="button">Reopen</button>`
              : `<span class="muted" title="Result evicted from cache">Evicted</span>`;
            const varClass = e.total_variance == null ? "" : (e.total_variance > 0 ? "v-adverse" : (e.total_variance < 0 ? "v-favourable" : ""));
            return `
              <tr class="${rowClass}" data-run-id="${escapeAttr(e.id || "")}">
                <td><code>${escapeHtml(ts)}</code></td>
                <td>${escapeHtml(e.company || "")}</td>
                <td>${escapeHtml(e.period_label || "")}</td>
                <td><code>${escapeHtml(e.source_filename || "")}</code></td>
                <td class="num">${e.rows || 0}</td>
                <td class="num ${varClass}">${e.total_variance == null ? "-" : formatGBPSigned(e.total_variance)}</td>
                <td class="num">${e.total_variance_pct == null ? "-" : formatPctSigned(e.total_variance_pct)}</td>
                <td class="num">${e.rag_red == null ? "-" : e.rag_red}</td>
                <td class="num">$${(e.cost_usd || 0).toFixed(4)}</td>
                <td class="history-actions">${reopenBtn}</td>
              </tr>
            `;
          }).join("")}
        </tbody>
      </table>
    `;

    historyBody.querySelectorAll(".history-reopen").forEach(btn => {
      btn.addEventListener("click", (ev) => {
        ev.stopPropagation();
        reopenRun(btn.dataset.runId);
      });
    });
    historyBody.querySelectorAll(".history-row.cached").forEach(tr => {
      tr.addEventListener("click", () => reopenRun(tr.dataset.runId));
    });
  }

  async function reopenRun(runId) {
    if (!runId) return;
    try {
      const r = await fetch(`/api/runs/${encodeURIComponent(runId)}`);
      if (!r.ok) {
        const d = await r.json().catch(() => ({}));
        return showError(d.error || `HTTP ${r}`);
      }
      const data = await r.json();
      closeModal(historyModal);
      // Make sure analysis result panels are visible (compare may have hidden them).
      [pivotRow, pivotParentCard, $("#commentary-card"), $("#heatmap-card"),
       detailCard, $("#warnings-card"), exportEl, compareCard].forEach(el => {
         if (el) el.hidden = (el === compareCard);
       });
      render(data);
      showState("results");
      resetBtn.hidden = false;
    } catch (e) {
      showError(`Could not reopen run: ${e.message}`);
    }
  }

  if (historyOpen)    historyOpen.addEventListener("click", () => { fetchHistory(); openModal(historyModal); });
  if (historyRefresh) historyRefresh.addEventListener("click", fetchHistory);

  // Global keyboard shortcuts
  document.addEventListener("keydown", (e) => {
    // Ignore typing inside inputs / contenteditable.
    const tag = (document.activeElement && document.activeElement.tagName) || "";
    const inField = tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT";
    if (e.key === "Escape") {
      [historyModal, shortcutsModal].forEach(m => { if (m && !m.hidden) closeModal(m); });
      return;
    }
    if (inField) return;
    if (e.key === "?" || (e.shiftKey && e.key === "/")) {
      e.preventDefault();
      openModal(shortcutsModal);
    } else if (e.key === "h" || e.key === "H") {
      e.preventDefault();
      fetchHistory();
      openModal(historyModal);
    } else if (e.key === "/") {
      const search = $("#detail-search");
      if (search) { e.preventDefault(); search.focus(); }
    }
  });

  // --- restore last result --------------------------------------------
  (function restoreLast() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      const { at, data } = JSON.parse(raw);
      if (!data || Date.now() - at > 6 * 3600 * 1000) return;
      render(data);
      showState("results");
      resetBtn.hidden = false;
    } catch (_) {}
  })();

  // --- escapers --------------------------------------------------------
  function escapeHtml(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  }
  function escapeAttr(s) { return String(s == null ? "" : s).replace(/[^a-zA-Z0-9_-]/g, ""); }
})();
