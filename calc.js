/* Orders Placed — the calculation. No DOM in here: the page and the Node checks
 * load this same file, so what is verified is what runs.
 *
 * Every function takes sheets as [{ name, rows }] where rows is an array of
 * arrays — SheetJS `sheet_to_json(ws, { header: 1, raw: true })`. Dates arrive
 * as Excel serial numbers and are converted here in UTC: SheetJS's own Date
 * objects are built in local time and can move midnight on the 1st into the
 * previous month, which would put a day's orders in the wrong column.
 */
(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  else root.OrdersCalc = api;
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  const PVS = "Proj vs SO";
  const FR = "Fill Rate";
  const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

  const PVS_COLS = ["Month", "Year", "Item Group", "item_parent", "SO Qty"];
  const FR_COLS = ["Sales Order Date", "Stock Qty in KGS", "Item Group", "NEW MIS ITEM GROUP", "Parent Item", "Item Name"];

  // Fill-rate filter columns: master field → row field → header in the report.
  const FR_FILTERS = [
    { field: "itemGroup", label: "Item Group" },
    { field: "newMis", label: "NEW MIS ITEM GROUP" },
    { field: "parentItem", label: "Parent Item" },
  ];

  const GROUP_HEADERS = ["Group", "Source", "Item Group", "NEW MIS ITEM GROUP", "Parent Item", "Item Name contains"];
  const MAP_HEADERS = ["Match column", "Value", "Group"];

  /* ── text and numbers ───────────────────────────────────────────── */

  // Labels arrive with stray spaces, tabs and mixed case: "Trail mix" vs
  // "Trail Mix", "Satva Mix " with a trailing space.
  const clean = v => (v == null ? "" : String(v)).replace(/\s+/g, " ").trim();
  const norm = v => clean(v).toLowerCase();
  const normHeader = v => norm(v).replace(/[_.]+/g, " ").replace(/\s+/g, " ").trim();

  function toNum(v) {
    if (typeof v === "number") return Number.isFinite(v) ? v : 0;
    if (v == null) return 0;
    const n = parseFloat(String(v).replace(/,/g, ""));
    return Number.isFinite(n) ? n : 0;
  }

  // A cached cell value that really is a number — "#N/A", blanks and text are not.
  function numOrNull(v) {
    if (typeof v === "number") return Number.isFinite(v) ? v : null;
    if (typeof v === "string" && /^-?[\d,]*\.?\d+$/.test(v.trim())) return parseFloat(v.replace(/,/g, ""));
    return null;
  }

  // Sums of kg carry float noise (249.6000000000001); keep them to the gram.
  const toGram = n => Math.round(n * 1000) / 1000;

  // "a, b; c" → ["a", "b", "c"], normalised. Used for every multi-value master cell.
  const splitList = v => String(v == null ? "" : v).split(/[,;\n]/).map(norm).filter(Boolean);
  const showList = v => String(v == null ? "" : v).split(/[,;\n]/).map(clean).filter(Boolean).join(", ");

  /* ── months ─────────────────────────────────────────────────────── */

  const monthKey = (y, m) => `${y}-${String(m).padStart(2, "0")}`;
  const keyParts = k => k.split("-").map(Number);
  const monthLabel = k => { const [y, m] = keyParts(k); return `${MONTHS[m - 1]} ${y}`; };

  const validYM = (y, m) => (y >= 1990 && y <= 2100 && m >= 1 && m <= 12 ? { y, m } : null);
  const fullYear = s => { const n = parseInt(s, 10); return n < 100 ? 2000 + n : n; };
  const monthFromName = s => MONTHS.findIndex(n => n.toLowerCase() === norm(s).slice(0, 3)) + 1;

  function serialYM(n) {
    const d = new Date(Math.round((n - 25569) * 86400) * 1000);
    return validYM(d.getUTCFullYear(), d.getUTCMonth() + 1);
  }

  // Year and month of a date cell: an Excel serial, a Date, "2026-09-30",
  // "30-09-2026" / "30/09/2026" (day first, as these exports write it),
  // "30-Sep-2026", "Sep-26", "September 2026".
  function parseYM(v) {
    if (v == null || v === "") return null;
    if (v instanceof Date) return isNaN(v) ? null : validYM(v.getFullYear(), v.getMonth() + 1);
    if (typeof v === "number") return v >= 20000 && v < 80000 ? serialYM(v) : null;
    const s = String(v).trim();
    let r;
    if ((r = s.match(/^(\d{4})[-/.](\d{1,2})(?:[-/.]\d{1,2})?(?:[ T].*)?$/))) return validYM(+r[1], +r[2]);
    if ((r = s.match(/^\d{1,2}[-/.](\d{1,2})[-/.](\d{4}|\d{2})(?:[ T].*)?$/))) return validYM(fullYear(r[2]), +r[1]);
    if ((r = s.match(/^(?:\d{1,2}[-/ ])?([A-Za-z]{3,9})\.?[-/ ,']+(\d{4}|\d{2})(?:[ T].*)?$/))) {
      const m = monthFromName(r[1]);
      return m ? validYM(fullYear(r[2]), m) : null;
    }
    if (/^\d+(\.\d+)?$/.test(s)) return parseYM(Number(s));
    return null;
  }

  // Proj vs SO carries the month as a name ("September") beside a Year column.
  function pvsMonth(monthCell, yearCell) {
    let y = typeof yearCell === "number" ? Math.trunc(yearCell) : parseInt(clean(yearCell), 10);
    let m = 0;
    const s = typeof monthCell === "string" ? monthCell.trim() : null;
    if (typeof monthCell === "number" && monthCell >= 1 && monthCell <= 12) m = Math.trunc(monthCell);
    else if (s && /^\d{1,2}$/.test(s)) m = +s;
    else if (s && /^[A-Za-z]{3,9}\.?$/.test(s)) m = monthFromName(s);
    else {
      const p = parseYM(monthCell);
      if (p) { m = p.m; if (!y) y = p.y; }
    }
    const ok = validYM(y, m);
    return ok ? monthKey(ok.y, ok.m) : null;
  }

  /* ── reading the inputs ─────────────────────────────────────────── */

  // The header row may sit below a title; look through the first 15 rows.
  function findHeader(rows, required) {
    const want = required.map(normHeader);
    for (let r = 0; r < Math.min(rows.length, 15); r++) {
      const idx = new Map();
      (rows[r] || []).forEach((h, c) => { const k = normHeader(h); if (k && !idx.has(k)) idx.set(k, c); });
      if (want.every(w => idx.has(w))) {
        const col = {};
        required.forEach((name, i) => { col[name] = idx.get(want[i]); });
        return { row: r, col };
      }
    }
    return null;
  }

  // Distinct labels per column, normalised → as first written. Feeds the
  // "value not found" checks and the master editor's suggestions.
  function labelCollector(names) {
    const maps = {};
    names.forEach(n => { maps[n] = new Map(); });
    return {
      add(name, raw) { const k = norm(raw); if (k && !maps[name].has(k)) maps[name].set(k, clean(raw)); },
      maps,
    };
  }

  function readProjVsSo(sheet) {
    const h = findHeader(sheet.rows, PVS_COLS);
    if (!h) return null;
    const c = h.col;
    const labels = labelCollector(["Item Group", "item_parent"]);
    const rows = [];
    const months = new Map();
    let undated = 0, undatedQty = 0;
    for (let r = h.row + 1; r < sheet.rows.length; r++) {
      const row = sheet.rows[r];
      if (!row) continue;
      const qty = toNum(row[c["SO Qty"]]);
      const month = pvsMonth(row[c.Month], row[c.Year]);
      if (!month) {
        if (qty) { undated++; undatedQty += qty; }
        continue;
      }
      labels.add("Item Group", row[c["Item Group"]]);
      labels.add("item_parent", row[c.item_parent]);
      rows.push({ month, itemGroup: norm(row[c["Item Group"]]), itemParent: norm(row[c.item_parent]), qty });
      const mm = months.get(month) || { rows: 0, qty: 0 };
      mm.rows++; mm.qty += qty;
      months.set(month, mm);
    }
    return { kind: "pvs", sheet: sheet.name, headerRow: h.row + 1, rows, months, undated, undatedQty, labels: labels.maps };
  }

  function readFillRate(sheet) {
    const h = findHeader(sheet.rows, FR_COLS);
    if (!h) return null;
    const c = h.col;
    const labels = labelCollector(["Item Group", "NEW MIS ITEM GROUP", "Parent Item"]);
    const rows = [];
    const months = new Map();
    let undated = 0, undatedQty = 0;
    for (let r = h.row + 1; r < sheet.rows.length; r++) {
      const row = sheet.rows[r];
      if (!row) continue;
      const qty = toNum(row[c["Stock Qty in KGS"]]);
      const ym = parseYM(row[c["Sales Order Date"]]);
      // A row with no order date cannot be put in a month. Footers ("Total")
      // carry no kg and are dropped quietly; anything with kg is reported.
      if (!ym) {
        if (qty) { undated++; undatedQty += qty; }
        continue;
      }
      const month = monthKey(ym.y, ym.m);
      labels.add("Item Group", row[c["Item Group"]]);
      labels.add("NEW MIS ITEM GROUP", row[c["NEW MIS ITEM GROUP"]]);
      labels.add("Parent Item", row[c["Parent Item"]]);
      rows.push({
        month,
        itemGroup: norm(row[c["Item Group"]]),
        newMis: norm(row[c["NEW MIS ITEM GROUP"]]),
        parentItem: norm(row[c["Parent Item"]]),
        itemName: norm(row[c["Item Name"]]),
        qty,
      });
      const mm = months.get(month) || { rows: 0, qty: 0 };
      mm.rows++; mm.qty += qty;
      months.set(month, mm);
    }
    return { kind: "fr", sheet: sheet.name, headerRow: h.row + 1, rows, months, undated, undatedQty, labels: labels.maps };
  }

  // A previous output: "Group" in column A of a header row whose other cells
  // are months. Grand Total rows are skipped — they are recomputed.
  function readPrevious(sheet) {
    const rows = sheet.rows;
    for (let r = 0; r < Math.min(rows.length, 15); r++) {
      const hdr = rows[r] || [];
      if (norm(hdr[0]) !== "group") continue;
      const cols = [];
      for (let c = 1; c < hdr.length; c++) {
        const p = parseYM(hdr[c]);
        if (p) cols.push({ c, month: monthKey(p.y, p.m) });
      }
      if (!cols.length) continue;
      const groups = [];
      const totals = { kg: {}, tons: {} };
      for (let i = r + 1; i < rows.length; i++) {
        const row = rows[i] || [];
        const name = clean(row[0]);
        if (!name) continue;
        const values = {};
        for (const { c, month } of cols) {
          const v = numOrNull(row[c]);
          if (v != null) values[month] = v;
        }
        if (/^grand total/i.test(name)) {
          if (/ton/i.test(name)) Object.assign(totals.tons, values);
          else Object.assign(totals.kg, values);
          continue;
        }
        groups.push({ group: name, values });
      }
      const months = [...new Set(cols.map(x => x.month))].sort();
      return { kind: "prev", sheet: sheet.name, headerRow: r + 1, months, groups, totals };
    }
    return null;
  }

  const READERS = { pvs: readProjVsSo, fr: readFillRate, prev: readPrevious };

  // Every sheet in a workbook that reads as the given kind, in workbook order.
  function detect(kind, sheets) {
    const out = [];
    for (const s of sheets) {
      const res = READERS[kind](s);
      if (res) out.push(res);
    }
    return out;
  }

  // Parses a workbook and finds every input it holds, so one file carrying
  // both tabs can fill both slots. XLSX (SheetJS) is passed in so the page,
  // its worker and Node all run this. `raw` keeps CSV text as written — a
  // day-first "05/09/2026" must not be re-read as 9 May.
  function readWorkbook(XLSX, data) {
    const bytes = data instanceof ArrayBuffer ? new Uint8Array(data) : data;
    const wb = XLSX.read(bytes, { type: "array", dense: true, raw: true, cellDates: false, cellText: false, cellNF: false, cellStyles: false, cellHTML: false });
    const sheets = wb.SheetNames.map(name => ({
      name,
      rows: XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null, blankrows: true }),
    }));
    let master = null;
    try { master = readMaster(sheets); } catch (e) { master = null; }
    return { sheetNames: wb.SheetNames, pvs: detect("pvs", sheets), fr: detect("fr", sheets), prev: detect("prev", sheets), master };
  }

  /* ── master ─────────────────────────────────────────────────────── */

  function normSource(v) {
    const s = norm(v);
    if (s.startsWith("proj") || s.includes("vs so") || s === "pvs") return PVS;
    if (s.startsWith("fill") || s === "fr") return FR;
    return null;
  }

  function normMapColumn(v) {
    const s = normHeader(v);
    if (s === "item group") return "itemGroup";
    if (s === "item parent" || s === "parent item" || s === "parent") return "itemParent";
    return null;
  }

  // "16" or "16 g" means a pack size: 16 followed by g / gm / gms / gram(s),
  // and not part of a bigger number — "160 g" and "1.16 g" are not 16 g.
  // Anything else is a plain case-insensitive "contains".
  const UNIT = "(?:g|gm|gms|gram|grams)";
  function nameTest(token) {
    const m = token.match(new RegExp(`^(\\d+(?:\\.\\d+)?)\\s*${UNIT}?$`));
    if (m) {
      const re = new RegExp(`(?:^|[^\\d.])${m[1].replace(".", "\\.")}\\s*${UNIT}(?![a-z])`);
      return s => re.test(s);
    }
    return s => s.includes(token);
  }

  function compileFillRate(g) {
    const tests = [];
    const parts = [];
    const lists = [];
    for (const { field, label } of FR_FILTERS) {
      const list = splitList(g[field]);
      if (!list.length) continue;
      const set = new Set(list);
      tests.push(row => set.has(row[field]));
      parts.push(`${label} ${list.length > 1 ? "in" : "="} ${showList(g[field])}`);
      lists.push({ label, list, shown: showList(g[field]).split(", ") });
    }
    const names = splitList(g.itemName);
    if (names.length) {
      const fns = names.map(nameTest);
      tests.push(row => fns.some(f => f(row.itemName)));
      parts.push(`Item Name has ${showList(g.itemName)}`);
    }
    return { tests, lists, text: parts.length ? parts.join(" AND ") : "no rule — 0" };
  }

  /* ── the calculation ────────────────────────────────────────────── */

  function compute(master, pvs, fr, prev) {
    const notes = [];
    const warn = text => notes.push({ level: "warn", text });
    const info = text => notes.push({ level: "info", text });

    const months = fr ? [...fr.months.keys()].sort() : [];

    // Groups, in output order. A blank name is an unfinished row; a repeat is
    // ignored so a Group is never counted twice in the total.
    const groups = [];
    const byName = new Map();
    for (const raw of master.groups || []) {
      const name = clean(raw.group);
      if (!name) continue;
      if (byName.has(norm(name))) { warn(`"${name}" is listed twice in the master; only the first row is used.`); continue; }
      let source = normSource(raw.source);
      if (!source) { warn(`"${name}" has no valid Source; treated as Fill Rate.`); source = FR; }
      const g = { group: name, source, raw };
      groups.push(g);
      byName.set(norm(name), g);
    }

    // Proj vs SO mapping. A Group matched through several rows counts each
    // Proj vs SO row once — Berry Mix is listed under both Item Group and
    // item_parent and must not double.
    const pvsRules = new Map();
    for (const r of master.pvsMap || []) {
      const gName = clean(r.group), value = norm(r.value);
      if (!gName && !value) continue;
      const field = normMapColumn(r.column);
      if (!field) { warn(`Proj vs SO mapping "${clean(r.value)}" has match column "${clean(r.column)}"; use Item Group or item_parent.`); continue; }
      if (!value) continue;
      const g = byName.get(norm(gName));
      if (!g) { warn(`Proj vs SO mapping "${clean(r.value)}" points to Group "${gName}", which is not in the Groups list.`); continue; }
      if (g.source !== PVS) { warn(`Proj vs SO mapping "${clean(r.value)}" points to "${g.group}", whose Source is Fill Rate — the mapping is not used.`); continue; }
      const colLabel = field === "itemGroup" ? "Item Group" : "item_parent";
      if (!pvsRules.has(g.group)) pvsRules.set(g.group, []);
      pvsRules.get(g.group).push({ field, value, label: `${colLabel} = ${clean(r.value)}`, colLabel, raw: clean(r.value) });
      if (pvs && !pvs.labels[colLabel].has(value)) {
        warn(`${g.group}: ${colLabel} "${clean(r.value)}" does not appear anywhere in the Proj vs SO file.`);
      }
    }

    const missingPvsMonths = pvs ? months.filter(m => !pvs.months.has(m)) : months.slice();
    const noRule = [];
    const rows = [];

    for (const g of groups) {
      const values = {}, exact = {}, matched = {};
      months.forEach(m => { exact[m] = 0; matched[m] = 0; });
      let rule;

      if (g.source === FR) {
        const fr1 = compileFillRate(g.raw);
        rule = fr1.text;
        if (fr) {
          for (const { label, list, shown } of fr1.lists) {
            list.forEach((v, i) => {
              if (!fr.labels[label].has(v)) warn(`${g.group}: ${label} "${shown[i]}" does not appear anywhere in the fill rate.`);
            });
          }
        }
        if (!fr1.tests.length) noRule.push(g.group);
        else if (fr) {
          for (const row of fr.rows) {
            if (!(row.month in exact)) continue;
            if (fr1.tests.every(t => t(row))) { exact[row.month] += row.qty; matched[row.month]++; }
          }
        }
        months.forEach(m => { values[m] = toGram(exact[m]); });
      } else {
        const conds = pvsRules.get(g.group) || [];
        rule = conds.length ? `SO Qty where ${conds.map(c => c.label).join(" OR ")}` : "no Proj vs SO mapping — 0";
        if (!conds.length) warn(`${g.group} is a Proj vs SO Group with no mapping row, so it shows 0.`);
        if (pvs && conds.length) {
          for (const row of pvs.rows) {
            if (!(row.month in exact)) continue;
            if (conds.some(c => row[c.field] === c.value)) { exact[row.month] += row.qty; matched[row.month]++; }
          }
        }
        months.forEach(m => { values[m] = missingPvsMonths.includes(m) ? null : toGram(exact[m]); });
      }
      rows.push({ group: g.group, source: g.source, rule, values, exact, matched, history: false });
    }

    if (noRule.length) info(`No rule, so 0: ${noRule.join(", ")}.`);
    if (pvs && missingPvsMonths.length) {
      warn(`Proj vs SO has no rows for ${missingPvsMonths.map(monthLabel).join(", ")} — its Groups are blank for ${missingPvsMonths.length > 1 ? "those months" : "that month"}.`);
    }
    if (pvs) {
      const unused = [...pvs.months.keys()].filter(m => !months.includes(m)).sort();
      if (unused.length) info(`Proj vs SO also has ${unused.map(monthLabel).join(", ")}; only the fill-rate ${months.length > 1 ? "months are" : "month is"} used.`);
      if (pvs.undated) warn(`${pvs.undated} Proj vs SO rows (${fmt(pvs.undatedQty)} SO Qty) have no readable Month/Year and were left out.`);
    }
    if (fr && fr.undated) warn(`${fr.undated} fill-rate rows (${fmt(fr.undatedQty)} kg) have no readable Sales Order Date and were left out.`);

    // History from a previous output: its months are kept, a month computed
    // now replaces the old column, and Groups the master no longer has keep
    // their history at the bottom.
    let allMonths = months.slice();
    if (prev) {
      allMonths = [...new Set([...prev.months, ...months])].sort();
      const prevBy = new Map(prev.groups.map(p => [norm(p.group), p]));
      const used = new Set();
      for (const row of rows) {
        const p = prevBy.get(norm(row.group));
        if (!p) continue;
        used.add(norm(row.group));
        for (const [m, v] of Object.entries(p.values)) if (!months.includes(m)) row.values[m] = v;
      }
      for (const p of prev.groups) {
        if (used.has(norm(p.group))) continue;
        const values = {};
        for (const [m, v] of Object.entries(p.values)) if (!months.includes(m)) values[m] = v;
        months.forEach(m => { values[m] = null; });
        rows.push({ group: p.group, source: "Previous output", rule: "not in the master — history only", values, exact: {}, matched: {}, history: true });
        warn(`"${p.group}" is in the previous output but not in the master; its history is kept and ${months.map(monthLabel).join(", ")} is blank.`);
      }
      const replaced = months.filter(m => prev.months.includes(m));
      if (replaced.length) info(`${replaced.map(monthLabel).join(", ")} was already in the previous output and has been replaced with the figures computed now.`);
    }

    // Grand Total is a plain sum of the rows, as in the history: the size rows
    // (Makhana Roasted 20 g, Cashewnut Roasted 50 g …) sit inside their parent
    // row and are counted again in the total. Values stay unrounded so the
    // total is exact; they are only displayed as whole kg. A month carried from
    // a previous output keeps the totals it had there — some of those were
    // taken before rounding and differ from the sum of the rows by a kg or two.
    const totalKg = {}, totalTons = {}, totalFrom = {};
    for (const m of allMonths) {
      if (prev && !months.includes(m) && prev.totals.kg[m] != null) {
        totalKg[m] = prev.totals.kg[m];
        totalTons[m] = prev.totals.tons[m] != null ? prev.totals.tons[m] : Math.round(totalKg[m] / 1000);
        totalFrom[m] = "previous";
        continue;
      }
      let sum = 0, any = false;
      for (const row of rows) { const v = row.values[m]; if (v != null) { sum += v; any = true; } }
      totalKg[m] = any ? sum : null;
      totalTons[m] = any ? Math.round(sum / 1000) : null;
      totalFrom[m] = "sum";
    }

    return { months: allMonths, newMonths: months, rows, totalKg, totalTons, totalFrom, notes };
  }

  const fmt = n => Math.round(n).toLocaleString("en-IN");

  /* ── master as a workbook ───────────────────────────────────────── */

  function masterToSheets(master) {
    return [
      {
        name: "Master - Groups",
        rows: [GROUP_HEADERS, ...master.groups.map(g => [g.group, g.source, g.itemGroup, g.newMis, g.parentItem, g.itemName])],
      },
      {
        name: "Master - Proj vs SO",
        rows: [MAP_HEADERS, ...master.pvsMap.map(m => [m.column, m.value, m.group])],
      },
    ];
  }

  // Reads the two master tables back from any workbook that has them —
  // an exported master or an output file, which carries the same sheets.
  function readMaster(sheets) {
    let groups = null, pvsMap = null;
    for (const s of sheets) {
      if (!groups) {
        const h = findHeader(s.rows, GROUP_HEADERS);
        if (h) {
          const c = h.col;
          groups = s.rows.slice(h.row + 1)
            .filter(r => r && clean(r[c.Group]))
            .map(r => ({
              group: clean(r[c.Group]),
              source: normSource(r[c.Source]) || FR,
              itemGroup: clean(r[c["Item Group"]]),
              newMis: clean(r[c["NEW MIS ITEM GROUP"]]),
              parentItem: clean(r[c["Parent Item"]]),
              itemName: clean(r[c["Item Name contains"]]),
            }));
          continue;
        }
      }
      if (!pvsMap) {
        const h = findHeader(s.rows, MAP_HEADERS);
        if (h) {
          const c = h.col;
          pvsMap = s.rows.slice(h.row + 1)
            .filter(r => r && (clean(r[c.Value]) || clean(r[c.Group])))
            .map(r => ({ column: normMapColumn(r[c["Match column"]]) === "itemGroup" ? "Item Group" : "item_parent", value: clean(r[c.Value]), group: clean(r[c.Group]) }));
        }
      }
    }
    if (!groups) throw new Error(`No Groups table found — expected a sheet with the headers ${GROUP_HEADERS.join(", ")}.`);
    if (!pvsMap) throw new Error(`No Proj vs SO mapping found — expected a sheet with the headers ${MAP_HEADERS.join(", ")}.`);
    return { groups, pvsMap };
  }

  // Fills missing fields so a master from storage or an older export is safe to use.
  function normaliseMaster(m) {
    const str = v => (v == null ? "" : String(v));
    return {
      groups: (m && Array.isArray(m.groups) ? m.groups : []).map(g => ({
        group: str(g.group), source: normSource(g.source) || FR,
        itemGroup: str(g.itemGroup), newMis: str(g.newMis), parentItem: str(g.parentItem), itemName: str(g.itemName),
      })),
      pvsMap: (m && Array.isArray(m.pvsMap) ? m.pvsMap : []).map(r => ({
        column: normMapColumn(r.column) === "itemGroup" ? "Item Group" : "item_parent", value: str(r.value), group: str(r.group),
      })),
    };
  }

  return {
    PVS, FR, PVS_COLS, FR_COLS, GROUP_HEADERS, MAP_HEADERS,
    norm, clean, parseYM, pvsMonth, monthKey, keyParts, monthLabel, nameTest,
    detect, readWorkbook, readProjVsSo, readFillRate, readPrevious, compute,
    masterToSheets, readMaster, normaliseMaster,
  };
});
