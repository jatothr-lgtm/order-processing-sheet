/* Orders Placed — the page: input slots, the master editor, the preview and the
 * Excel export. Every number comes from calc.js; nothing here changes one. */
(function () {
  "use strict";

  const C = window.OrdersCalc;
  const DEFAULT = C.normaliseMaster(window.ORDERS_DEFAULT_MASTER);
  const STORE_KEY = "ordersPlaced.master.v1";

  const $ = id => document.getElementById(id);
  const esc = s => String(s == null ? "" : s).replace(/[&<>"']/g, ch => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch]));
  const fmtKg = v => (v == null ? "" : Math.round(v).toLocaleString("en-IN"));
  const clone = o => JSON.parse(JSON.stringify(o));

  const state = {
    slots: { pvs: null, fr: null, prev: null },
    master: loadMaster(),
    result: null,
  };

  /* ── master storage (this browser only) ─────────────────────────── */

  function loadMaster() {
    try {
      const raw = localStorage.getItem(STORE_KEY);
      if (raw) {
        const m = JSON.parse(raw);
        // Saved before normalisations existed: start from the default ones.
        if (!Array.isArray(m.normalisations)) m.normalisations = clone(DEFAULT.normalisations);
        return C.normaliseMaster(m);
      }
    } catch (e) { /* storage blocked or corrupt: fall back to the default */ }
    return clone(DEFAULT);
  }

  const isEdited = () => JSON.stringify(state.master) !== JSON.stringify(DEFAULT);

  function saveMaster() {
    try {
      if (isEdited()) localStorage.setItem(STORE_KEY, JSON.stringify(state.master));
      else localStorage.removeItem(STORE_KEY);
    } catch (e) { /* not fatal: the edit still applies to this session */ }
    const pill = $("masterPill");
    pill.textContent = isEdited() ? "Edited" : "Default";
    pill.classList.toggle("edited", isEdited());
  }

  /* ── reading files ──────────────────────────────────────────────── */

  // Parse in a worker so a 45,000-line fill rate does not freeze the page.
  // Where a worker cannot start (opened from disk, CDN blocked inside the
  // worker) the same code runs here instead.
  let worker = null, seq = 0;
  function getWorker() {
    if (worker === false) return null;
    if (!worker) {
      try { worker = new Worker("reader.js"); } catch (e) { worker = false; return null; }
    }
    return worker;
  }

  async function readWorkbookFile(file) {
    const buf = await file.arrayBuffer();
    const w = getWorker();
    if (!w) return C.readWorkbook(window.XLSX, buf);
    return new Promise((resolve, reject) => {
      const id = ++seq;
      const done = () => { w.removeEventListener("message", onMsg); w.removeEventListener("error", onErr); };
      const onMsg = e => {
        if (e.data.id !== id) return;
        done();
        if (e.data.ok) resolve(e.data.result); else reject(new Error(e.data.error));
      };
      const onErr = e => {
        done();
        if (e && e.preventDefault) e.preventDefault();
        worker = false;
        try { resolve(C.readWorkbook(window.XLSX, buf)); } catch (err) { reject(err); }
      };
      w.addEventListener("message", onMsg);
      w.addEventListener("error", onErr);
      w.postMessage({ id, buffer: buf });
    });
  }

  /* ── input slots ────────────────────────────────────────────────── */

  const NEEDS = {
    pvs: "the Proj vs SO columns (Month, Year, Item Group, item_parent, SO Qty)",
    fr: "the fill rate columns (Sales Order Date, Stock Qty in KGS, Item Group, NEW MIS ITEM GROUP, Parent Item, Item Name)",
    prev: "the Orders Placed layout (“Group” in column A, months across the header row)",
  };

  // Where a workbook has several matching sheets (a full Proj vs SO tab and
  // single-month copies of it), take the largest.
  function defaultPick(kind, cands) {
    if (kind === "prev") return 0;
    let best = 0;
    cands.forEach((c, i) => { if (c.rows.length > cands[best].rows.length) best = i; });
    return best;
  }

  const current = kind => {
    const s = state.slots[kind];
    return s && s.candidates ? s.candidates[s.pick] : null;
  };

  async function loadFile(kind, file) {
    const token = ++seq;
    state.slots[kind] = { fileName: file.name, busy: true, token };
    renderSlot(kind);
    let wb;
    try {
      wb = await readWorkbookFile(file);
    } catch (err) {
      if (state.slots[kind] && state.slots[kind].token === token) {
        state.slots[kind] = { fileName: file.name, error: `Could not read this file: ${err.message}` };
        renderSlot(kind);
      }
      return;
    }
    if (!state.slots[kind] || state.slots[kind].token !== token) return;   // replaced meanwhile

    const cands = wb[kind];
    state.slots[kind] = cands.length
      ? { fileName: file.name, candidates: cands, pick: defaultPick(kind, cands) }
      : { fileName: file.name, error: `No sheet in this workbook has ${NEEDS[kind]}.` };
    renderSlot(kind);

    // One workbook may carry both inputs; fill the other required slot if it is empty.
    for (const other of ["pvs", "fr"]) {
      const s = state.slots[other];
      if (other === kind || (s && (s.candidates || s.busy)) || !wb[other].length) continue;
      state.slots[other] = { fileName: file.name, candidates: wb[other], pick: defaultPick(other, wb[other]), auto: true };
      renderSlot(other);
    }
    recompute();
  }

  function describe(kind, d) {
    const head = d.headerRow > 1 ? ` · header on row ${d.headerRow}` : "";
    if (kind === "pvs") {
      const months = [...d.months.keys()].sort().map(C.monthLabel).join(", ") || "no readable months";
      return `${d.rows.length.toLocaleString("en-IN")} rows · ${months}${head}`;
    }
    if (kind === "fr") {
      const months = [...d.months.entries()].sort((a, b) => (a[0] < b[0] ? -1 : 1)).map(([m, v]) => `${C.monthLabel(m)} (${v.rows.toLocaleString("en-IN")} lines, ${fmtKg(v.qty)} kg)`).join(", ") || "no readable Sales Order Dates";
      return `${months}${head}`;
    }
    const span = d.months.length ? `${C.monthLabel(d.months[0])} – ${C.monthLabel(d.months[d.months.length - 1])}` : "no months";
    return `${d.groups.length} groups · ${span}${head}`;
  }

  function renderSlot(kind) {
    const el = document.querySelector(`.slot[data-slot="${kind}"]`);
    const st = el.querySelector(".slot-status");
    const s = state.slots[kind];
    el.classList.toggle("loaded", !!(s && s.candidates));
    if (!s) { st.innerHTML = ""; return; }
    const head = `<div class="row"><span class="file">${esc(s.fileName)}</span>${s.busy ? "" : `<button class="icon del" data-act="clear" title="Remove this file" aria-label="Remove this file">✕</button>`}</div>`;
    if (s.busy) { st.innerHTML = `${head}<div class="busy">Reading — a month of fill rate takes a few seconds</div>`; return; }
    if (s.error) { st.innerHTML = `${head}<div class="err">${esc(s.error)}</div>`; return; }
    const d = s.candidates[s.pick];
    const sheet = s.candidates.length > 1
      ? `Sheet <select data-act="pick">${s.candidates.map((c, i) => `<option value="${i}"${i === s.pick ? " selected" : ""}>${esc(c.sheet)}</option>`).join("")}</select>`
      : `Sheet “${esc(d.sheet)}”`;
    st.innerHTML = `${head}<div>${sheet}${s.auto ? " · from the same workbook" : ""}</div><div>${esc(describe(kind, d))}</div>`;
  }

  document.querySelectorAll(".slot").forEach(el => {
    const kind = el.dataset.slot;
    const input = el.querySelector("input[type=file]");
    const drop = el.querySelector(".drop");
    input.addEventListener("change", () => { if (input.files[0]) loadFile(kind, input.files[0]); input.value = ""; });
    drop.addEventListener("dragover", e => { e.preventDefault(); drop.classList.add("over"); });
    drop.addEventListener("dragleave", () => drop.classList.remove("over"));
    drop.addEventListener("drop", e => {
      e.preventDefault();
      drop.classList.remove("over");
      const f = e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) loadFile(kind, f);
    });
    el.querySelector(".slot-status").addEventListener("click", e => {
      if (e.target.closest("[data-act=clear]")) { state.slots[kind] = null; renderSlot(kind); recompute(); }
    });
    el.querySelector(".slot-status").addEventListener("change", e => {
      if (e.target.dataset.act !== "pick") return;
      state.slots[kind].pick = +e.target.value;
      renderSlot(kind);
      recompute();
    });
  });

  /* ── calculation + preview ──────────────────────────────────────── */

  let recomputeTimer = null;
  const recomputeSoon = () => { clearTimeout(recomputeTimer); recomputeTimer = setTimeout(recompute, 250); };

  function recompute() {
    const pvs = current("pvs"), fr = current("fr"), prev = current("prev");
    fillDatalists(pvs, fr);
    state.result = pvs && fr ? C.compute(state.master, pvs, fr, prev) : null;
    renderOutput();
    updateHits();
  }

  function renderOutput() {
    const res = state.result;
    const notes = $("notes"), wrap = $("outWrap"), empty = $("outEmpty"), btn = $("btnDownload");
    if (!res || !res.newMonths.length) {
      notes.innerHTML = "";
      wrap.hidden = true;
      empty.hidden = false;
      btn.disabled = true;
      $("outTitle").textContent = "Output";
      const missing = ["pvs", "fr"].filter(k => !current(k)).map(k => (k === "pvs" ? "the Proj vs SO file" : "the fill rate"));
      empty.textContent = missing.length
        ? `Load ${missing.join(" and ")} to see the output.`
        : "The fill rate has no rows with a readable Sales Order Date, so there is no month to report.";
      return;
    }
    $("outTitle").textContent = `Output — ${res.newMonths.map(C.monthLabel).join(", ")}`;
    notes.innerHTML = res.notes.map(n => `<div class="note ${n.level}">${esc(n.text)}</div>`).join("");

    // The months computed now, plus up to three before them for comparison.
    const firstNew = res.months.indexOf(res.newMonths[0]);
    const hist = res.months.slice(Math.max(0, firstNew - 3), firstNew).filter(m => !res.newMonths.includes(m));
    const cols = [...hist.map(m => ({ m, cls: "hist" })), ...res.newMonths.map(m => ({ m, cls: "new" }))];
    const srcBadge = s => `<span class="src ${s === C.PVS ? "pvs" : s === C.FR ? "fr" : "prev"}">${esc(s)}</span>`;

    const head = `<thead><tr><th>Group</th><th>Source</th><th>Rule</th>${cols.map(c => `<th class="n ${c.cls}">${C.monthLabel(c.m)}${c.cls === "hist" ? "<br><small>previous</small>" : ""}</th>`).join("")}</tr></thead>`;
    const body = res.rows.map(r => `<tr>
        <td class="grp">${esc(r.group)}</td>
        <td>${srcBadge(r.source)}</td>
        <td class="rule">${esc(r.rule)}</td>
        ${cols.map(c => {
          const v = r.values[c.m];
          const tip = c.cls === "new" && r.matched[c.m] != null ? ` title="${r.matched[c.m].toLocaleString("en-IN")} source rows"` : "";
          return `<td class="n ${c.cls}"${tip}>${v == null ? '<span class="blank">—</span>' : fmtKg(v)}</td>`;
        }).join("")}
      </tr>`).join("");
    const foot = `<tr class="total"><td colspan="3">Grand Total Kgs</td>${cols.map(c => `<td class="n ${c.cls}">${fmtKg(res.totalKg[c.m])}</td>`).join("")}</tr>
      <tr class="total"><td colspan="3">Grand Total Tons</td>${cols.map(c => `<td class="n ${c.cls}">${fmtKg(res.totalTons[c.m])}</td>`).join("")}</tr>`;
    $("outTable").innerHTML = `${head}<tbody>${body}${foot}</tbody>`;
    wrap.hidden = false;
    empty.hidden = true;
    btn.disabled = false;
  }

  /* ── the Excel output ───────────────────────────────────────────── */

  const colLetter = n => { let s = ""; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };

  async function exportXlsx() {
    const res = state.result;
    if (!res) return;
    const wb = new ExcelJS.Workbook();
    wb.creator = "Orders Placed";
    wb.created = new Date();
    const font = { name: "Aptos Narrow", size: 11 };
    const bold = { ...font, bold: true };
    const line = { style: "thin", color: { argb: "FFBFC8C2" } };

    // Laid out as the working Sheet1: title, a year above the first month of
    // each year, "Group" and the months, one row per Group, then the totals.
    const ws = wb.addWorksheet("Orders Placed", { views: [{ state: "frozen", xSplit: 1, ySplit: 2 }] });
    ws.getColumn(1).width = 30;
    ws.getCell(1, 1).value = "Orders Placed";
    ws.getCell(1, 1).font = { ...bold, size: 12 };
    ws.getCell(2, 1).value = "Group";
    ws.getCell(2, 1).font = bold;
    res.months.forEach((m, i) => {
      const [y, mm] = C.keyParts(m);
      const col = i + 2;
      ws.getColumn(col).width = 11;
      if (i === 0 || mm === 1) { ws.getCell(1, col).value = y; ws.getCell(1, col).font = bold; }
      const c = ws.getCell(2, col);
      c.value = new Date(Date.UTC(y, mm - 1, 1));
      c.numFmt = "mmm-yy";
      c.font = bold;
      c.alignment = { horizontal: "right" };
      c.border = { bottom: line };
    });
    ws.getCell(2, 1).border = { bottom: line };

    const first = 3, last = 2 + res.rows.length;
    res.rows.forEach((r, j) => {
      const R = first + j;
      ws.getCell(R, 1).value = r.group;
      ws.getCell(R, 1).font = font;
      res.months.forEach((m, i) => {
        const v = r.values[m];
        if (v == null) return;
        const c = ws.getCell(R, i + 2);
        c.value = v;
        c.numFmt = "#,##0";
        c.font = font;
      });
    });

    // Months computed now total by formula over unrounded values; months
    // carried from a previous output keep the totals they had there.
    const kgRow = last + 1, tonRow = last + 2;
    ws.getCell(kgRow, 1).value = "Grand Total Kgs";
    ws.getCell(tonRow, 1).value = "Grand Total Tons";
    res.months.forEach((m, i) => {
      const L = colLetter(i + 2);
      const kg = ws.getCell(kgRow, i + 2), tons = ws.getCell(tonRow, i + 2);
      if (res.totalKg[m] == null) return;
      if (res.totalFrom[m] === "previous") {
        kg.value = res.totalKg[m];
        tons.value = res.totalTons[m];
      } else {
        kg.value = { formula: `SUM(${L}${first}:${L}${last})`, result: res.totalKg[m] };
        tons.value = { formula: `ROUND(${L}${kgRow}/1000,0)`, result: res.totalTons[m] };
      }
      kg.numFmt = "#,##0";
      tons.numFmt = "#,##0";
    });
    ws.getRow(kgRow).font = bold;
    ws.getRow(tonRow).font = bold;
    for (let c = 1; c <= res.months.length + 1; c++) ws.getCell(kgRow, c).border = { top: line };

    // Workings: how each figure computed now was reached.
    const wk = wb.addWorksheet("Workings", { views: [{ state: "frozen", ySplit: 1 }] });
    wk.columns = [
      { header: "Month", key: "month", width: 11 },
      { header: "Group", key: "group", width: 30 },
      { header: "Source", key: "source", width: 13 },
      { header: "Rule", key: "rule", width: 90 },
      { header: "Source rows", key: "rows", width: 12 },
      { header: "Kg (unrounded)", key: "kg", width: 16 },
    ];
    for (const m of res.newMonths) {
      for (const r of res.rows) {
        if (r.history) continue;
        wk.addRow({ month: C.monthLabel(m), group: r.group, source: r.source, rule: r.rule, rows: r.matched[m], kg: r.values[m] });
      }
    }
    wk.getColumn("kg").numFmt = "#,##0.000";
    wk.getRow(1).font = bold;

    // Notes: what went in, and everything flagged on screen.
    const nt = wb.addWorksheet("Notes");
    nt.columns = [{ width: 22 }, { width: 120 }];
    const fileLine = k => {
      const s = state.slots[k], d = current(k);
      return d ? `${s.fileName} — sheet “${d.sheet}”` : "not loaded";
    };
    const lines = [
      ["Generated", new Date().toLocaleString("en-IN")],
      ["Months computed", res.newMonths.map(C.monthLabel).join(", ")],
      ["Proj vs SO", fileLine("pvs")],
      ["Fill rate", fileLine("fr")],
      ["Previous output", fileLine("prev")],
      ["Master", isEdited() ? "edited — see the two Master sheets" : "default — see the two Master sheets"],
      [],
      ...res.notes.map(n => [n.level === "warn" ? "Check" : "Note", n.text]),
    ];
    lines.forEach(l => nt.addRow(l));
    nt.getColumn(1).font = bold;
    nt.getColumn(2).alignment = { wrapText: true, vertical: "top" };

    // The master it was built with — this file can be uploaded as a master.
    for (const s of C.masterToSheets(state.master)) {
      const ms = wb.addWorksheet(s.name, { views: [{ state: "frozen", ySplit: 1 }] });
      s.rows.forEach(r => ms.addRow(r));
      ms.getRow(1).font = bold;
      ms.columns.forEach((col, i) => { col.width = i === 0 ? 30 : 34; });
    }

    const buf = await wb.xlsx.writeBuffer();
    const label = res.newMonths.length === 1
      ? C.monthLabel(res.newMonths[0]).replace(" ", "-")
      : `${C.monthLabel(res.newMonths[0]).replace(" ", "-")}_to_${C.monthLabel(res.newMonths[res.newMonths.length - 1]).replace(" ", "-")}`;
    download(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), `Orders-Placed_${label}.xlsx`);
  }

  function download(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  $("btnDownload").addEventListener("click", () => {
    exportXlsx().catch(err => toast(`Export failed: ${err.message}`));
  });

  /* ── master editor ──────────────────────────────────────────────── */

  const FILTERS = [
    { f: "itemGroup", label: "Item Group", list: "dl-fr-itemGroup", cls: "col-flt" },
    { f: "newMis", label: "NEW MIS ITEM GROUP", list: "dl-fr-newMis", cls: "col-flt" },
    { f: "parentItem", label: "Parent Item", list: "dl-fr-parentItem", cls: "col-wide" },
    { f: "itemName", label: "Item Name contains", list: "", cls: "col-sm" },
  ];

  // Suggestions for a normalisation's values come from the column it rewrites.
  const NORM_LIST = {
    "Fill Rate|Item Group": "dl-fr-itemGroup",
    "Fill Rate|NEW MIS ITEM GROUP": "dl-fr-newMis",
    "Fill Rate|Parent Item": "dl-fr-parentItem",
    "Fill Rate|Item Name": "dl-fr-itemName",
    "Proj vs SO|Item Group": "dl-pvs-itemGroup",
    "Proj vs SO|item_parent": "dl-pvs-itemParent",
  };
  const normColumns = file => (C.NORM_COLUMNS[file] || C.NORM_COLUMNS[C.FR]).map(([label]) => label);

  function renderMaster() {
    renderGroups();
    renderMap();
    renderNorm();
    saveMaster();
  }

  function setCount(kind, n) {
    $(`cnt-${kind}`).textContent = n ? `· ${n}` : "";
  }

  function renderGroups() {
    const g = state.master.groups;
    $("tblGroups").innerHTML = `<thead><tr><th></th><th class="col-group">Group</th><th class="col-src">Source</th>${FILTERS.map(x => `<th class="${x.cls}">${x.label}</th>`).join("")}<th></th></tr></thead>
      <tbody>${g.map((row, i) => {
        const pvs = row.source === C.PVS;
        return `<tr data-i="${i}">
          <td class="idx">${i + 1}</td>
          <td><input data-f="group" value="${esc(row.group)}" placeholder="Group name"></td>
          <td><select data-f="source"><option${pvs ? "" : " selected"}>${C.FR}</option><option${pvs ? " selected" : ""}>${C.PVS}</option></select></td>
          ${FILTERS.map(x => `<td><input data-f="${x.f}" value="${esc(pvs ? "" : row[x.f])}" title="${esc(row[x.f])}"${x.list ? ` list="${x.list}"` : ""}${pvs ? ' disabled placeholder="via the mapping"' : ""}></td>`).join("")}
          <td class="ord">
            <button class="icon" data-act="up" title="Move up" aria-label="Move up"${i === 0 ? " disabled" : ""}>↑</button>
            <button class="icon" data-act="down" title="Move down" aria-label="Move down"${i === g.length - 1 ? " disabled" : ""}>↓</button>
            <button class="icon del" data-act="del" title="Remove" aria-label="Remove">✕</button>
          </td>
        </tr>`;
      }).join("")}</tbody>`;
    setCount("groups", g.length);
  }

  function renderMap() {
    const m = state.master.pvsMap;
    $("tblMap").innerHTML = `<thead><tr><th></th><th class="col-src">Match column</th><th class="col-group">Value</th><th class="col-group">Group</th><th></th></tr></thead>
      <tbody>${m.map((row, i) => {
        const ig = row.column === "Item Group";
        return `<tr data-i="${i}">
          <td class="idx">${i + 1}</td>
          <td><select data-f="column"><option${ig ? " selected" : ""}>Item Group</option><option${ig ? "" : " selected"}>item_parent</option></select></td>
          <td><input data-f="value" value="${esc(row.value)}" list="${ig ? "dl-pvs-itemGroup" : "dl-pvs-itemParent"}" placeholder="as written in Proj vs SO"></td>
          <td><input data-f="group" value="${esc(row.group)}" list="dl-groups-pvs" placeholder="Group it counts towards"></td>
          <td class="ord"><button class="icon del" data-act="del" title="Remove" aria-label="Remove">✕</button></td>
        </tr>`;
      }).join("")}</tbody>`;
    setCount("pvsMap", m.length);
    fillGroupList();
  }

  function renderNorm() {
    const list = state.master.normalisations;
    const opts = (values, current) => {
      const all = values.includes(current) || !current ? values : [...values, current];
      return all.map(v => `<option${v === current ? " selected" : ""}>${esc(v)}</option>`).join("");
    };
    $("tblNorm").innerHTML = `<thead><tr><th></th><th class="col-src">File</th><th class="col-src">Column</th><th class="col-wide">Raw value — as in the file</th><th class="col-wide">Normalised value — used by the rules</th><th class="n" title="Lines or rows this normalisation changed in the files loaded">Changed</th><th></th></tr></thead>
      <tbody>${list.map((row, i) => {
        const dl = NORM_LIST[`${row.file}|${row.column}`] || "";
        return `<tr data-i="${i}">
          <td class="idx">${i + 1}</td>
          <td><select data-f="file">${opts([C.FR, C.PVS], row.file)}</select></td>
          <td><select data-f="column">${opts(normColumns(row.file), row.column)}</select></td>
          <td><input data-f="raw" value="${esc(row.raw)}" title="${esc(row.raw)}"${dl ? ` list="${dl}"` : ""} placeholder="value to replace"></td>
          <td><input data-f="value" value="${esc(row.value)}" title="${esc(row.value)}"${dl ? ` list="${dl}"` : ""} placeholder="replace it with"></td>
          <td class="n hits" data-hit="${i}"></td>
          <td class="ord"><button class="icon del" data-act="del" title="Remove" aria-label="Remove">✕</button></td>
        </tr>`;
      }).join("")}</tbody>`;
    setCount("normalisations", list.length);
    updateHits();
  }

  // How many lines each normalisation changed, from the latest calculation.
  function updateHits() {
    const hits = state.result ? state.result.normHits : null;
    document.querySelectorAll("#tblNorm td.hits").forEach(td => {
      const v = hits ? hits[+td.dataset.hit] : null;
      td.textContent = v == null ? "—" : v.toLocaleString("en-IN");
      td.classList.toggle("zero", v === 0);
      td.title = v == null ? "Shows once both files are loaded" : v === 0 ? "Matches nothing in the files loaded" : `${v} lines or rows changed`;
    });
  }

  function masterChanged() {
    saveMaster();
    recomputeSoon();
  }

  // Groups table: typing edits in place (keeps focus); structure changes re-render.
  $("tblGroups").addEventListener("focusin", e => {
    if (e.target.dataset.f === "group") e.target.dataset.old = e.target.value;
  });
  $("tblGroups").addEventListener("input", e => {
    const tr = e.target.closest("tr[data-i]"), f = e.target.dataset.f;
    if (!tr || !f || f === "source") return;
    state.master.groups[+tr.dataset.i][f] = e.target.value;
    e.target.title = e.target.value;
    masterChanged();
  });
  $("tblGroups").addEventListener("change", e => {
    const tr = e.target.closest("tr[data-i]"), f = e.target.dataset.f;
    if (!tr) return;
    if (f === "source") {
      state.master.groups[+tr.dataset.i].source = e.target.value;
      renderMaster();
      masterChanged();
    } else if (f === "group") {
      // A renamed Group keeps its Proj vs SO mapping. Done once the name is
      // finished, so clearing it on the way to a new one does not orphan rows.
      const old = C.norm(e.target.dataset.old), now = e.target.value;
      if (old && old !== C.norm(now)) {
        state.master.pvsMap.forEach(r => { if (C.norm(r.group) === old) r.group = now; });
        renderMap();
        masterChanged();
      }
      e.target.dataset.old = now;
    }
  });
  $("tblGroups").addEventListener("click", e => {
    const b = e.target.closest("button[data-act]"), tr = e.target.closest("tr[data-i]");
    if (!b || !tr) return;
    const i = +tr.dataset.i, g = state.master.groups;
    if (b.dataset.act === "up" && i > 0) [g[i - 1], g[i]] = [g[i], g[i - 1]];
    else if (b.dataset.act === "down" && i < g.length - 1) [g[i + 1], g[i]] = [g[i], g[i + 1]];
    else if (b.dataset.act === "del") {
      if (!confirm(`Remove “${g[i].group || "this row"}” from the master?`)) return;
      g.splice(i, 1);
    }
    renderMaster();
    masterChanged();
  });

  $("tblMap").addEventListener("input", e => {
    const tr = e.target.closest("tr[data-i]"), f = e.target.dataset.f;
    if (!tr || !f || f === "column") return;
    state.master.pvsMap[+tr.dataset.i][f] = e.target.value;
    masterChanged();
  });
  $("tblMap").addEventListener("change", e => {
    const tr = e.target.closest("tr[data-i]");
    if (!tr || e.target.dataset.f !== "column") return;
    state.master.pvsMap[+tr.dataset.i].column = e.target.value;
    renderMap();
    masterChanged();
  });
  $("tblMap").addEventListener("click", e => {
    const b = e.target.closest("button[data-act=del]"), tr = e.target.closest("tr[data-i]");
    if (!b || !tr) return;
    state.master.pvsMap.splice(+tr.dataset.i, 1);
    renderMap();
    masterChanged();
  });

  $("tblNorm").addEventListener("input", e => {
    const tr = e.target.closest("tr[data-i]"), f = e.target.dataset.f;
    if (!tr || (f !== "raw" && f !== "value")) return;
    state.master.normalisations[+tr.dataset.i][f] = e.target.value;
    e.target.title = e.target.value;
    masterChanged();
  });
  $("tblNorm").addEventListener("change", e => {
    const tr = e.target.closest("tr[data-i]"), f = e.target.dataset.f;
    if (!tr || (f !== "file" && f !== "column")) return;
    const row = state.master.normalisations[+tr.dataset.i];
    row[f] = e.target.value;
    // Another file has other columns: keep the column if it exists there.
    if (f === "file" && !normColumns(row.file).includes(row.column)) row.column = normColumns(row.file)[0];
    renderNorm();
    fillDatalistsNow();
    masterChanged();
  });
  $("tblNorm").addEventListener("click", e => {
    const b = e.target.closest("button[data-act=del]"), tr = e.target.closest("tr[data-i]");
    if (!b || !tr) return;
    state.master.normalisations.splice(+tr.dataset.i, 1);
    renderNorm();
    masterChanged();
  });

  /* ── per-table add, export, import, paste ───────────────────────── */

  const FILE_SLUG = { groups: "groups", pvsMap: "proj-vs-so-mapping", normalisations: "normalisations" };
  const TABLE_ID = { groups: "tblGroups", pvsMap: "tblMap", normalisations: "tblNorm" };
  const FIRST_INPUT = { groups: "group", pvsMap: "value", normalisations: "raw" };

  function addRow(kind) {
    if (kind === "groups") state.master.groups.push({ group: "", source: C.FR, itemGroup: "", newMis: "", parentItem: "", itemName: "" });
    else if (kind === "pvsMap") state.master.pvsMap.push({ column: "Item Group", value: "", group: "" });
    else state.master.normalisations.push({ file: C.FR, column: "Parent Item", raw: "", value: "" });
    renderMaster();
    const inputs = $(TABLE_ID[kind]).querySelectorAll(`input[data-f="${FIRST_INPUT[kind]}"]`);
    inputs[inputs.length - 1].focus();
  }

  function writeSheets(sheets, fileName) {
    const wb = XLSX.utils.book_new();
    for (const s of sheets) {
      const ws = XLSX.utils.aoa_to_sheet(s.rows);
      ws["!cols"] = s.rows[0].map((_, i) => ({ wch: i === 0 ? 26 : 40 }));
      XLSX.utils.book_append_sheet(wb, ws, s.name);
    }
    XLSX.writeFile(wb, fileName);
  }

  function exportTable(kind) {
    writeSheets([C.tableToSheet(kind, state.master[kind])], `Orders-Placed_${FILE_SLUG[kind]}.xlsx`);
  }

  async function sheetsFromFile(file) {
    const wb = XLSX.read(new Uint8Array(await file.arrayBuffer()), { type: "array", dense: true, raw: true, cellDates: false });
    return wb.SheetNames.map(name => ({ name, rows: XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null, blankrows: true }) }));
  }

  // Rows copied from Excel arrive tab-separated; anything else is read as CSV.
  function parsePaste(kind, text) {
    if (!text.trim()) return null;
    let rows;
    if (text.includes("\t")) {
      rows = text.replace(/\r/g, "").split("\n").map(line => line.split("\t"));
    } else {
      const wb = XLSX.read(text, { type: "string", raw: true });
      rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: null });
    }
    return C.readTable(kind, [{ name: "pasted", rows }], { positional: true });
  }

  async function importFile(kind, file) {
    try {
      const sheets = await sheetsFromFile(file);
      // A file without the header row is read in column order, like a paste.
      const parsed = C.readTable(kind, sheets) || (sheets.length === 1 ? C.readTable(kind, sheets, { positional: true }) : null);
      if (!parsed) { toast(`No sheet in ${file.name} has the headers ${C.TABLES[kind].headers.join(", ")}.`); return; }
      openImport(kind, parsed, file.name);
    } catch (err) {
      toast(`Could not read ${file.name}: ${err.message}`);
    }
  }

  let imp = null;   // the import being reviewed: { kind, parsed, from, paste }

  function openImport(kind, parsed, from) {
    imp = { kind, parsed, from, paste: !parsed };
    const t = C.TABLES[kind];
    $("impTitle").textContent = `${imp.paste ? "Paste" : "Import"} — ${t.title}`;
    $("impPasteBox").hidden = !imp.paste;
    if (imp.paste) {
      $("impPasteHint").innerHTML = `Copy the rows in Excel and paste them here, with or without the header row. Columns, in this order: <b>${t.headers.map(esc).join(" · ")}</b>.`;
      $("impText").value = "";
    }
    renderImportSummary();
    // showModal keeps the last returnValue: clear it, or Escape after an
    // earlier "Add & update" would apply this import too.
    $("importDlg").returnValue = "";
    $("importDlg").showModal();
    if (imp.paste) $("impText").focus();
  }

  function renderImportSummary() {
    const box = $("impSummary"), add = $("impAdd"), rep = $("impReplace");
    const p = imp && imp.parsed;
    const current = imp ? state.master[imp.kind] : [];
    const skippedHtml = p && p.skipped.length
      ? `<p class="warn">${p.skipped.length} row${p.skipped.length === 1 ? "" : "s"} skipped:</p><ul>${p.skipped.slice(0, 8).map(s => `<li>Line ${s.line}: ${esc(s.reason)}</li>`).join("")}${p.skipped.length > 8 ? `<li>… and ${p.skipped.length - 8} more</li>` : ""}</ul>`
      : "";
    if (!p || !p.rows.length) {
      box.innerHTML = (imp && imp.paste && !p ? "<p>Nothing pasted yet.</p>" : "<p>No usable rows.</p>") + skippedHtml;
      add.disabled = rep.disabled = true;
      return;
    }
    const a = C.mergeTable(imp.kind, current, p.rows, "add");
    const r = C.mergeTable(imp.kind, current, p.rows, "replace");
    box.innerHTML = `<p><b>${p.rows.length} row${p.rows.length === 1 ? "" : "s"} read</b>${imp.from ? ` from ${esc(imp.from)}` : ""}.</p>${skippedHtml}
      <p><b>Add &amp; update</b> — ${a.added} new, ${a.updated} changed, ${a.unchanged} already the same; ${a.rows.length} rows after.</p>
      <p><b>Replace all</b> — the ${current.length} current rows are replaced by these ${r.rows.length}.</p>`;
    add.disabled = rep.disabled = false;
  }

  let pasteTimer = null;
  $("impText").addEventListener("input", () => {
    clearTimeout(pasteTimer);
    pasteTimer = setTimeout(() => { imp.parsed = parsePaste(imp.kind, $("impText").value); renderImportSummary(); }, 150);
  });

  $("importDlg").addEventListener("close", () => {
    const mode = $("importDlg").returnValue;
    if (!imp || (mode !== "add" && mode !== "replace")) { imp = null; return; }
    if (imp.paste) imp.parsed = parsePaste(imp.kind, $("impText").value);
    if (!imp.parsed || !imp.parsed.rows.length) { imp = null; return; }
    const res = C.mergeTable(imp.kind, state.master[imp.kind], imp.parsed.rows, mode);
    state.master[imp.kind] = res.rows;
    const title = C.TABLES[imp.kind].title;
    imp = null;
    renderMaster();
    fillDatalistsNow();
    masterChanged();
    toast(mode === "replace" ? `${title}: replaced with ${res.rows.length} rows` : `${title}: ${res.added} added, ${res.updated} updated`);
  });

  document.querySelectorAll(".card[data-table]").forEach(card => {
    const kind = card.dataset.table;
    card.querySelector(".card-head").addEventListener("click", e => {
      const b = e.target.closest("button[data-act]");
      if (!b) return;
      if (b.dataset.act === "add") addRow(kind);
      else if (b.dataset.act === "export") exportTable(kind);
      else if (b.dataset.act === "paste") openImport(kind, null, null);
    });
    const input = card.querySelector('input[data-act="import"]');
    input.addEventListener("change", () => {
      const f = input.files[0];
      input.value = "";
      if (f) importFile(kind, f);
    });
  });

  /* ── the whole master ───────────────────────────────────────────── */

  $("btnMasterReset").addEventListener("click", () => {
    if (!confirm("Discard every edit to all three tables and go back to the default master?")) return;
    state.master = clone(DEFAULT);
    renderMaster();
    fillDatalistsNow();
    masterChanged();
    toast("Master reset to the default");
  });
  $("btnMasterDownload").addEventListener("click", () => {
    writeSheets(C.masterToSheets(state.master), "Orders-Placed_master.xlsx");
  });
  $("masterFile").addEventListener("change", async e => {
    const f = e.target.files[0];
    e.target.value = "";
    if (!f) return;
    try {
      const m = C.readMaster(await sheetsFromFile(f));
      // A master saved before normalisations existed has none: keep ours.
      const keepNorm = m.normalisations === null;
      state.master = C.normaliseMaster({ ...m, normalisations: keepNorm ? state.master.normalisations : m.normalisations });
      renderMaster();
      fillDatalistsNow();
      masterChanged();
      const g = state.master;
      toast(`Master loaded: ${g.groups.length} groups, ${g.pvsMap.length} mappings, ${g.normalisations.length} normalisations` +
        (keepNorm ? " (the file has none, so the current ones were kept)" : "") +
        (m.skipped.length ? ` · ${m.skipped.length} rows skipped` : ""));
    } catch (err) {
      toast(`Could not load the master: ${err.message}`);
    }
  });

  /* ── suggestions from the loaded files ──────────────────────────── */

  function setList(id, values) {
    $(id).innerHTML = [...values].sort((a, b) => a.localeCompare(b)).map(v => `<option value="${esc(v)}">`).join("");
  }
  // Values as they appear in the files, plus the normalised values the
  // master introduces, so either can be picked in a rule.
  function fillDatalists(pvs, fr) {
    const targets = id => state.master.normalisations
      .filter(n => NORM_LIST[`${n.file}|${n.column}`] === id && n.value.trim())
      .map(n => n.value.trim());
    const fill = (id, labels) => setList(id, new Set([...(labels ? labels.values() : []), ...targets(id)]));
    fill("dl-fr-itemGroup", fr && fr.labels["Item Group"]);
    fill("dl-fr-newMis", fr && fr.labels["NEW MIS ITEM GROUP"]);
    fill("dl-fr-parentItem", fr && fr.labels["Parent Item"]);
    fill("dl-fr-itemName", fr && fr.labels["Item Name"]);
    fill("dl-pvs-itemGroup", pvs && pvs.labels["Item Group"]);
    fill("dl-pvs-itemParent", pvs && pvs.labels.item_parent);
  }
  const fillDatalistsNow = () => fillDatalists(current("pvs"), current("fr"));
  function fillGroupList() {
    setList("dl-groups-pvs", state.master.groups.filter(g => g.source === C.PVS && g.group.trim()).map(g => g.group));
  }

  /* ── tabs, toast ────────────────────────────────────────────────── */

  document.querySelectorAll(".tab").forEach(t => t.addEventListener("click", () => {
    document.querySelectorAll(".tab").forEach(x => { x.classList.toggle("active", x === t); x.setAttribute("aria-selected", x === t); });
    document.querySelectorAll(".view").forEach(v => v.classList.toggle("active", v.id === `view-${t.dataset.tab}`));
  }));

  let toastTimer = null;
  function toast(msg) {
    const el = $("toast");
    el.textContent = msg;
    el.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => el.classList.remove("show"), 3800);
  }

  renderMaster();
  fillDatalistsNow();
  renderOutput();
})();
