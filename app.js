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
      if (raw) return C.normaliseMaster(JSON.parse(raw));
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

  function renderMaster() {
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
    renderMap();
    saveMaster();
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
    fillGroupList();
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

  $("btnAddGroup").addEventListener("click", () => {
    state.master.groups.push({ group: "", source: C.FR, itemGroup: "", newMis: "", parentItem: "", itemName: "" });
    renderMaster();
    const inputs = $("tblGroups").querySelectorAll('input[data-f="group"]');
    inputs[inputs.length - 1].focus();
  });
  $("btnAddMap").addEventListener("click", () => {
    state.master.pvsMap.push({ column: "Item Group", value: "", group: "" });
    renderMap();
    saveMaster();
    const inputs = $("tblMap").querySelectorAll('input[data-f="value"]');
    inputs[inputs.length - 1].focus();
  });
  $("btnMasterReset").addEventListener("click", () => {
    if (!confirm("Discard every edit and go back to the default master?")) return;
    state.master = clone(DEFAULT);
    renderMaster();
    masterChanged();
    toast("Master reset to the default");
  });
  $("btnMasterDownload").addEventListener("click", () => {
    const wb = XLSX.utils.book_new();
    for (const s of C.masterToSheets(state.master)) {
      const ws = XLSX.utils.aoa_to_sheet(s.rows);
      ws["!cols"] = s.rows[0].map((_, i) => ({ wch: i === 0 ? 30 : 34 }));
      XLSX.utils.book_append_sheet(wb, ws, s.name);
    }
    XLSX.writeFile(wb, "Orders-Placed_master.xlsx");
  });
  $("masterFile").addEventListener("change", async e => {
    const f = e.target.files[0];
    e.target.value = "";
    if (!f) return;
    try {
      const wb = C.readWorkbook(window.XLSX, await f.arrayBuffer());
      if (!wb.master) throw new Error("no Master - Groups and Master - Proj vs SO sheets found. Use a downloaded master or an output file.");
      state.master = C.normaliseMaster(wb.master);
      renderMaster();
      masterChanged();
      toast(`Master loaded: ${state.master.groups.length} groups, ${state.master.pvsMap.length} mappings`);
    } catch (err) {
      toast(`Could not load the master: ${err.message}`);
    }
  });

  /* ── suggestions from the loaded files ──────────────────────────── */

  function setList(id, values) {
    $(id).innerHTML = [...values].sort((a, b) => a.localeCompare(b)).map(v => `<option value="${esc(v)}">`).join("");
  }
  function fillDatalists(pvs, fr) {
    setList("dl-fr-itemGroup", fr ? fr.labels["Item Group"].values() : []);
    setList("dl-fr-newMis", fr ? fr.labels["NEW MIS ITEM GROUP"].values() : []);
    setList("dl-fr-parentItem", fr ? fr.labels["Parent Item"].values() : []);
    setList("dl-pvs-itemGroup", pvs ? pvs.labels["Item Group"].values() : []);
    setList("dl-pvs-itemParent", pvs ? pvs.labels.item_parent.values() : []);
  }
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
  renderOutput();
})();
