(function () {
  const payload = window.__PAYLOAD__;
  if (!payload) return;

  const mappingRoot = document.getElementById("mapping");
  const sourceList = document.getElementById("sourceList");
  const previewTable = document.getElementById("previewTable");
  const rowCountEl = document.getElementById("rowCount");
  const sourceCountEl = document.getElementById("sourceCount");
  const templateCountEl = document.getElementById("templateCount");
  const statusEl = document.getElementById("status");
  const exportBtn = document.getElementById("exportBtn");
  const swapBtn = document.getElementById("swapBtn");
  const exportForm = document.getElementById("exportForm");
  const payloadInput = document.getElementById("payloadInput");
  const mappingInput = document.getElementById("mappingInput");
  const swapForm = document.getElementById("swapForm");
  const swapPayloadInput = document.getElementById("swapPayloadInput");
  const bridgeSvg = document.getElementById("bridgeSvg");

  const mapperDesktop = document.getElementById("mapperDesktop");
  const mapperWizard = document.getElementById("mapperWizard");
  const wizardBody = document.getElementById("wizardBody");
  const wizPos = document.getElementById("wizPos");
  const wizTotal = document.getElementById("wizTotal");
  const wizBack = document.getElementById("wizBack");
  const wizNext = document.getElementById("wizNext");

  rowCountEl.textContent = String(payload.source.row_count ?? "");

  const sourceHeaders = payload.source.headers || [];
  const templateHeaders = payload.template.headers || [];
  const suggestions = payload.suggestions || {};

  sourceCountEl && (sourceCountEl.textContent = String(sourceHeaders.length));
  templateCountEl && (templateCountEl.textContent = String(templateHeaders.length));

  const state = {
    mapping: {}, // templateHeader -> {type, value}
    wizardIndex: 0,
  };

  function setStatus(kind, text) {
    statusEl.classList.remove("danger", "ok");
    if (kind) statusEl.classList.add(kind);
    statusEl.textContent = text || "";
  }

  function isMobileWizard() {
    return window.matchMedia && window.matchMedia("(max-width: 640px)").matches;
  }

  function computeCompleteness() {
    const missing = [];
    for (const th of templateHeaders) {
      const spec = state.mapping[th];
      if (!spec || !spec.type) missing.push(th);
      else if (spec.type === "source" && !spec.value) missing.push(th);
      else if (spec.type === "constant" && spec.value === undefined) missing.push(th);
      else if (spec.type === "blank") {
        // ok
      }
    }
    return missing;
  }

  function getMappedSourceHeaders() {
    const out = new Set();
    for (const th of templateHeaders) {
      const spec = state.mapping[th];
      if (spec?.type === "source" && spec.value) out.add(spec.value);
    }
    return out;
  }

  function refreshGate() {
    const missing = computeCompleteness();
    if (missing.length) {
      exportBtn.disabled = true;
      setStatus("danger", `Missing mappings: ${missing.length} column(s).`);
    } else {
      exportBtn.disabled = false;
      setStatus("ok", "All template columns are mapped. Ready to export.");
    }
    renderSourceBadges();
    drawBridge();
  }

  function bestSuggestion(th) {
    const list = suggestions[th] || [];
    if (!list.length) return null;
    return list[0];
  }

  function buildTargetRow(th, { wizard = false } = {}) {
    const wrap = document.createElement("div");
    wrap.className = wizard ? "targetRow wizardRow" : "targetRow";
    wrap.dataset.target = th;

    const top = document.createElement("div");
    top.className = "targetRowTop";

    const title = document.createElement("div");
    title.className = "targetTitle";
    title.textContent = th;

    const badge = document.createElement("span");
    badge.className = "badge unmapped";
    badge.textContent = "Unmapped";
    badge.dataset.badgeFor = th;

    top.appendChild(title);
    top.appendChild(badge);

    const hint = document.createElement("div");
    hint.className = "hint";
    const best = bestSuggestion(th);
    if (best && best.score >= 0.45) {
      hint.innerHTML = `Suggested: <b>${escapeHtml(best.source)}</b> (score ${best.score})`;
    } else {
      hint.textContent = "Suggested: (none)";
    }

    const controls = document.createElement("div");
    controls.className = "controls";

    const typeSel = document.createElement("select");
    typeSel.innerHTML = `
      <option value="">Choose…</option>
      <option value="source">Source column</option>
      <option value="constant">Constant value</option>
      <option value="blank">Blank</option>
    `;

    const valueInput = document.createElement("input");
    valueInput.type = "text";
    valueInput.placeholder = "Search source column…";
    valueInput.autocomplete = "off";

    const listId = `srcList_${hashId(th)}`;
    const dl = document.createElement("datalist");
    dl.id = listId;
    for (const sh of sourceHeaders) {
      const opt = document.createElement("option");
      opt.value = sh;
      dl.appendChild(opt);
    }
    valueInput.setAttribute("list", listId);

    function setBadgeFromSpec() {
      const spec = state.mapping[th];
      const b = wrap.querySelector(`[data-badge-for="${cssEscape(th)}"]`) || badge;
      if (!spec || !spec.type) {
        b.className = "badge unmapped";
        b.textContent = "Unmapped";
        return;
      }
      b.className = "badge mapped";
      if (spec.type === "source") b.textContent = "Mapped";
      else if (spec.type === "constant") b.textContent = "Constant";
      else if (spec.type === "blank") b.textContent = "Blank";
      else b.textContent = "Mapped";
    }

    function apply() {
      const t = typeSel.value;
      if (!t) {
        delete state.mapping[th];
        valueInput.value = "";
        valueInput.placeholder = "Search source column…";
      } else if (t === "blank") {
        state.mapping[th] = { type: "blank", value: "" };
        valueInput.value = "";
        valueInput.placeholder = "Blank (no value)";
      } else if (t === "constant") {
        state.mapping[th] = { type: "constant", value: valueInput.value || "" };
        valueInput.placeholder = "Enter constant value…";
      } else if (t === "source") {
        state.mapping[th] = { type: "source", value: valueInput.value || "" };
        valueInput.placeholder = "Search source column…";
      }
      setBadgeFromSpec();
      refreshGate();
    }

    typeSel.addEventListener("change", () => {
      const t = typeSel.value;
      if (t === "blank") valueInput.value = "";
      apply();
    });
    valueInput.addEventListener("input", () => {
      const spec = state.mapping[th];
      if (typeSel.value === "constant") {
        state.mapping[th] = { type: "constant", value: valueInput.value || "" };
      } else if (typeSel.value === "source") {
        state.mapping[th] = { type: "source", value: valueInput.value || "" };
      } else if (spec && spec.type === "constant") {
        spec.value = valueInput.value || "";
      }
      setBadgeFromSpec();
      refreshGate();
    });
    valueInput.addEventListener("blur", () => {
      // If source type selected and value not exact match, keep it but it will be blocked on export server-side if invalid.
      setBadgeFromSpec();
      refreshGate();
    });

    // Preselect best suggestion if strong.
    if (best && best.score >= 0.62) {
      typeSel.value = "source";
      valueInput.value = best.source;
      state.mapping[th] = { type: "source", value: best.source };
      setBadgeFromSpec();
    }

    controls.appendChild(typeSel);
    controls.appendChild(valueInput);

    wrap.appendChild(top);
    wrap.appendChild(controls);
    wrap.appendChild(hint);
    wrap.appendChild(dl);
    return wrap;
  }

  function renderMapping() {
    mappingRoot.innerHTML = "";
    for (const th of templateHeaders) {
      mappingRoot.appendChild(buildTargetRow(th));
    }
    refreshGate();
  }

  function renderWizard() {
    if (!wizardBody) return;
    wizardBody.innerHTML = "";
    const th = templateHeaders[state.wizardIndex];
    if (!th) return;
    wizardBody.appendChild(buildTargetRow(th, { wizard: true }));
    wizPos && (wizPos.textContent = String(state.wizardIndex + 1));
    wizTotal && (wizTotal.textContent = String(templateHeaders.length));
    wizBack && (wizBack.disabled = state.wizardIndex <= 0);
    wizNext &&
      (wizNext.textContent = state.wizardIndex >= templateHeaders.length - 1 ? "Done" : "Next");
    refreshGate();
  }

  function renderSourceList() {
    if (!sourceList) return;
    sourceList.innerHTML = "";
    for (const sh of sourceHeaders) {
      const card = document.createElement("div");
      card.className = "sourceCard";
      card.dataset.source = sh;

      const head = document.createElement("div");
      head.className = "sourceCardHeader";

      const title = document.createElement("div");
      title.className = "sourceCardTitle";
      title.textContent = sh;

      const badge = document.createElement("span");
      badge.className = "badge unmapped";
      badge.textContent = "Unused";
      badge.dataset.sourceBadge = sh;

      head.appendChild(title);
      head.appendChild(badge);

      const chips = document.createElement("div");
      chips.className = "sampleChips";
      // We don't have per-column samples in payload yet; keep chips empty for now.

      card.appendChild(head);
      card.appendChild(chips);
      sourceList.appendChild(card);
    }
    renderSourceBadges();
  }

  function renderSourceBadges() {
    if (!sourceList) return;
    const used = getMappedSourceHeaders();
    for (const sh of sourceHeaders) {
      const el = sourceList.querySelector(`[data-source-badge="${cssEscape(sh)}"]`);
      if (!el) continue;
      if (used.has(sh)) {
        el.className = "badge mapped";
        el.textContent = "Used";
      } else {
        el.className = "badge";
        el.textContent = "Unused";
      }
    }
  }

  function drawBridge() {
    if (!bridgeSvg || !mapperDesktop || isMobileWizard()) return;
    if (window.matchMedia("(max-width: 980px)").matches) return;

    const svgRect = bridgeSvg.getBoundingClientRect();
    const w = Math.max(0, Math.floor(svgRect.width));
    const h = Math.max(0, Math.floor(svgRect.height));
    bridgeSvg.setAttribute("viewBox", `0 0 ${w} ${h}`);
    bridgeSvg.setAttribute("width", String(w));
    bridgeSvg.setAttribute("height", String(h));
    bridgeSvg.innerHTML = "";

    // defs
    const defs = document.createElementNS("http://www.w3.org/2000/svg", "defs");
    const glow = document.createElementNS("http://www.w3.org/2000/svg", "filter");
    glow.setAttribute("id", "glow");
    glow.innerHTML =
      '<feGaussianBlur stdDeviation="2.2" result="coloredBlur"/><feMerge><feMergeNode in="coloredBlur"/><feMergeNode in="SourceGraphic"/></feMerge>';
    defs.appendChild(glow);
    bridgeSvg.appendChild(defs);

    function anchorForSource(sh) {
      const el = document.querySelector(`[data-source="${cssEscape(sh)}"]`);
      if (!el) return null;
      const r = el.getBoundingClientRect();
      return {
        x: r.right - svgRect.left,
        y: r.top + r.height / 2 - svgRect.top,
      };
    }
    function anchorForTarget(th) {
      const el = document.querySelector(`[data-target="${cssEscape(th)}"]`);
      if (!el) return null;
      const r = el.getBoundingClientRect();
      return {
        x: r.left - svgRect.left,
        y: r.top + r.height / 2 - svgRect.top,
      };
    }

    for (const th of templateHeaders) {
      const spec = state.mapping[th];
      if (!spec || spec.type !== "source" || !spec.value) continue;
      const a = anchorForSource(spec.value);
      const b = anchorForTarget(th);
      if (!a || !b) continue;

      const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
      const dx = Math.max(40, (b.x - a.x) * 0.7);
      const d = `M ${a.x} ${a.y} C ${a.x + dx} ${a.y}, ${b.x - dx} ${b.y}, ${b.x} ${b.y}`;
      path.setAttribute("d", d);
      path.setAttribute("fill", "none");
      path.setAttribute("stroke", "rgba(106,165,255,0.9)");
      path.setAttribute("stroke-width", "2");
      path.setAttribute("filter", "url(#glow)");
      path.setAttribute("stroke-linecap", "round");
      bridgeSvg.appendChild(path);
    }
  }

  function renderPreview() {
    const headers = payload.source.headers || [];
    const rows = payload.source.preview_rows || [];

    const thead = document.createElement("thead");
    const trh = document.createElement("tr");
    for (const h of headers) {
      const th = document.createElement("th");
      th.textContent = h;
      trh.appendChild(th);
    }
    thead.appendChild(trh);

    const tbody = document.createElement("tbody");
    for (const r of rows) {
      const tr = document.createElement("tr");
      for (const cell of r) {
        const td = document.createElement("td");
        td.textContent = cell ?? "";
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }

    previewTable.innerHTML = "";
    previewTable.appendChild(thead);
    previewTable.appendChild(tbody);
  }

  exportBtn.addEventListener("click", () => {
    const missing = computeCompleteness();
    if (missing.length) {
      setStatus("danger", `Missing mappings: ${missing.join(", ")}`);
      return;
    }
    payloadInput.value = JSON.stringify(payload);
    mappingInput.value = JSON.stringify(state.mapping);
    exportForm.submit();
  });

  swapBtn?.addEventListener("click", () => {
    setStatus("", "Swapping files…");
    swapPayloadInput.value = JSON.stringify(payload);
    swapForm.submit();
  });

  function syncLayoutMode() {
    if (!mapperDesktop || !mapperWizard) return;
    if (isMobileWizard()) {
      mapperDesktop.classList.add("hidden");
      mapperWizard.classList.remove("hidden");
      renderWizard();
    } else {
      mapperWizard.classList.add("hidden");
      mapperDesktop.classList.remove("hidden");
      renderMapping();
      drawBridge();
    }
  }

  wizBack?.addEventListener("click", () => {
    state.wizardIndex = Math.max(0, state.wizardIndex - 1);
    renderWizard();
  });
  wizNext?.addEventListener("click", () => {
    if (state.wizardIndex < templateHeaders.length - 1) {
      state.wizardIndex += 1;
      renderWizard();
    } else {
      // Done: go back to desktop view if user rotates / resizes later.
      setStatus("", "All set. You can Export from the top bar when ready.");
      refreshGate();
    }
  });

  window.addEventListener("resize", () => {
    syncLayoutMode();
    drawBridge();
  });
  window.addEventListener(
    "scroll",
    () => {
      drawBridge();
    },
    { passive: true }
  );

  function escapeHtml(s) {
    return String(s)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function cssEscape(s) {
    // minimal escape for attribute selectors
    return String(s).replaceAll('"', '\\"');
  }

  function hashId(s) {
    let h = 0;
    for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) >>> 0;
    return String(h);
  }

  renderSourceList();
  syncLayoutMode();
  renderPreview();
})();

