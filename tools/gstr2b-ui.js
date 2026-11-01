/* Page script for the GSTR-2B reconciler. Everything runs in this browser tab; nothing is uploaded. */
(function () {
  "use strict";
  const R = window.Recon;
  const $ = (id) => document.getElementById(id);
  const el = (tag, attrs, ...kids) => {
    const e = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (v === false || v === null || v === undefined) continue;
      if (k === "class") e.className = v;
      else if (k.startsWith("on")) e.addEventListener(k.slice(2), v);
      else e.setAttribute(k, v === true ? "" : v);
    }
    for (const k of kids.flat()) if (k !== null && k !== undefined && k !== false) e.append(k.nodeType ? k : String(k));
    return e;
  };
  const inr = new Intl.NumberFormat("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const rs = (v) => "₹" + inr.format(v);
  const kb = (n) => (n < 1024 * 1024 ? Math.max(1, Math.round(n / 1024)) + " KB" : (n / 1048576).toFixed(1) + " MB");

  const state = { books: null, g2b: [], booksInfo: null, g2bSheets: null, res: null, filter: "" };

  // ---------- file pickers ----------
  function setupDrop(zoneId, inputId, multi) {
    const zone = $(zoneId), input = $(inputId);
    const take = (files) => {
      files = Array.from(files || []);
      if (!files.length) return;
      if (multi) {
        for (const f of files) if (!state.g2b.some((g) => g.name === f.name && g.size === f.size)) state.g2b.push(f);
      } else state.books = files[0];
      input.value = "";
      changed();
    };
    input.addEventListener("change", () => take(input.files));
    zone.addEventListener("dragover", (e) => { e.preventDefault(); zone.classList.add("over"); });
    zone.addEventListener("dragleave", () => zone.classList.remove("over"));
    zone.addEventListener("drop", (e) => { e.preventDefault(); zone.classList.remove("over"); take(e.dataTransfer.files); });
  }

  function chip(file, onRemove) {
    return el("span", { class: "chip" }, el("span", { class: "fname" }, file.name), el("small", {}, kb(file.size)),
      el("button", { type: "button", class: "x", "aria-label": "Remove " + file.name, onclick: onRemove }, "×"));
  }

  function renderChips() {
    const b = $("booksChips"), g = $("g2bChips");
    b.replaceChildren(state.books ? chip(state.books, () => { state.books = null; changed(); }) : "");
    g.replaceChildren(...state.g2b.map((f, i) => chip(f, () => { state.g2b.splice(i, 1); changed(); })));
    $("booksZone").classList.toggle("done", !!state.books);
    $("g2bZone").classList.toggle("done", state.g2b.length > 0);
    $("runBtn").disabled = !(state.books && state.g2b.length);
  }

  function changed() {
    renderChips();
    state.booksInfo = state.g2bSheets = state.res = null;
    $("mapping").replaceChildren();
    $("results").replaceChildren();
    if (state.books && state.g2b.length) load();
    else setStatus(state.books || state.g2b.length ? "Now add the " + (state.books ? "GSTR-2B file(s)." : "purchase register.") : "", "");
  }

  function setStatus(msg, kind) {
    const s = $("status");
    s.replaceChildren(kind === "busy" ? el("span", { class: "spin", "aria-hidden": "true" }) : "", msg);
    s.className = kind || "";
  }

  // ---------- reading and running ----------
  async function load() {
    setStatus("Reading the files on this computer…", "busy");
    await new Promise((r) => setTimeout(r, 20));
    try {
      state.booksInfo = R.inspectBooks(await R.readWorkbook(state.books.name, await state.books.arrayBuffer()));
      state.g2bSheets = [];
      for (const f of state.g2b) state.g2bSheets.push({ file: f.name, sheets: await R.readWorkbook(f.name, await f.arrayBuffer()) });
      run();
    } catch (e) {
      setStatus(e.message || String(e), "err");
    }
  }

  function run() {
    $("results").replaceChildren();
    try {
      const tol = Math.max(0, parseFloat($("tol").value) || 0);
      renderMapping();
      const books = R.buildBooks(state.booksInfo);
      const g2b = R.build2B(state.g2bSheets);
      if (!g2b.rows.length) throw new Error(g2b.notes.join(" ") || "No GSTR-2B rows found.");
      state.res = R.reconcile(books.rows, g2b.rows, tol);
      const notes = g2b.notes.slice();
      if (books.flipped) notes.push(`${books.flipped} credit note(s) in the books had positive amounts and were treated as negative.`);
      render(state.res, notes);
      const S = state.res.summary, ok = S.tieBooks && S.tie2B && S.bridgeTies;
      setStatus(`Reconciled ${S.books.rows} books rows with ${S.g2b.rows} GSTR-2B rows. ` +
        (ok ? "All control totals tie." : "A control total does not tie: see the summary."), ok ? "ok" : "err");
    } catch (e) {
      setStatus(e.message || String(e), "err");
      const d = $("mapping").querySelector("details");
      if (d) d.open = true;
    }
  }

  function renderMapping() {
    const info = state.booksInfo, box = $("mapping");
    const found = R.BOOK_FIELDS.filter(([k]) => info.map[k] !== undefined).length;
    const missingReq = R.BOOK_FIELDS.some(([k, , req]) => req && info.map[k] === undefined);
    const grid = el("div", { class: "mapgrid" });
    for (const [key, label, req] of R.BOOK_FIELDS) {
      const sel = el("select", { id: "map-" + key, onchange: () => { info.map[key] = sel.value === "" ? undefined : +sel.value; run(); } },
        el("option", { value: "" }, req ? "(choose a column)" : "(not in file)"),
        info.headers.map((h, i) => (h ? el("option", { value: String(i) }, h) : null)));
      sel.value = info.map[key] === undefined ? "" : String(info.map[key]);
      grid.append(el("label", { for: "map-" + key }, label, req ? el("span", { class: "req" }, " *") : ""), sel);
    }
    box.replaceChildren(el("details", { class: "card", open: missingReq },
      el("summary", {}, `Columns read from the purchase register (${found} of ${R.BOOK_FIELDS.length} found)`),
      el("p", { class: "note" }, `Sheet "${info.sheet}", headers on row ${info.headerRow + 1}. If a column was picked wrongly, change it here; the result updates at once. * = needed.`),
      grid));
  }

  // ---------- results ----------
  const SECTIONS = [
    { key: "mismatch", tab: "mismatched values", title: "Mismatched values", pair: true,
      todo: "Check the invoice. Correct the books, or ask the supplier to amend. Where the GSTIN differs, fix the vendor master." },
    { key: "books_only", tab: "in-books-not-in-2B", title: "In books, not in GSTR-2B",
      todo: "Don't take this ITC yet. Follow up with the supplier and take it in the month it appears in 2B, within the section 16(4) time limit." },
    { key: "gstr2b_only", tab: "in-2B-not-in-books", title: "In GSTR-2B, not in books",
      todo: "Ask the client: goods not received, invoice not booked, or not their purchase. Credit notes here reduce ITC." },
    { key: "duplicate", tab: "possible-duplicates", title: "Possible duplicates in books",
      todo: "Probably entered twice. Reverse the extra entry: ITC can be taken only once." },
    { key: "flagged", tab: "flagged-for-review", title: "Flagged for review",
      todo: "GSTINs that fail the check digit, ledger heads that suggest blocked credit under section 17(5), and rows the 2B marks as ITC not available. A clean match does not make ITC eligible." },
    { key: "tolerance", tab: "matched-within-tolerance", title: "Matched within tolerance", pair: true,
      todo: "Rounding-size differences. Nothing to fix, but record the tolerance you used in the working paper." },
    { key: "matched", tab: "matched", title: "Matched", pair: true, todo: "Agree in full. Nothing to do." },
  ];
  const PAIR_COLS = ["Reason", "Supplier", "Supplier GSTIN (books)", "Supplier GSTIN (2B)", "Invoice no (books)", "Invoice no (2B)",
    "Invoice date", "Taxable (books)", "Taxable (2B)", "ITC (books)", "ITC (2B)", "ITC difference (2B less books)", "Books voucher", "2B ref"];
  const SHORT = { "Supplier GSTIN (books)": "GSTIN (books)", "Supplier GSTIN (2B)": "GSTIN (2B)", "Invoice no (books)": "Invoice (books)",
    "Invoice no (2B)": "Invoice (2B)", "ITC difference (2B less books)": "ITC diff", "Ledger head / 2B period": "Ledger / period",
    "Supplier GSTIN": "GSTIN", "Books voucher": "Voucher" };
  const NUMERIC_INT = new Set(["Row", "Books row"]);

  function sectionData(tabs, s) {
    let rows = tabs.find(([n]) => n === s.tab)[1];
    if (s.pair) {
      const head = rows[0], idx = PAIR_COLS.map((c) => head.indexOf(c));
      rows = rows.map((r) => idx.map((i) => r[i]));
      // show the 2B GSTIN only when it differs
      rows = rows.map((r, ri) => (ri && r[2] === r[3] ? Object.assign(r.slice(), { 3: "" }) : r));
    }
    return rows;
  }

  function itcOf(key) {
    const r = state.res;
    if (key === "flagged") return null;
    if (key === "gstr2b_only") return r.g2b.filter((x) => x.status === key).reduce((s, x) => s + x.itc, 0);
    if (["matched", "tolerance", "mismatch"].includes(key)) return r.paired.filter((p) => p.status === key).reduce((s, p) => s + p.b.itc, 0);
    return r.books.filter((x) => x.status === key).reduce((s, x) => s + x.itc, 0);
  }

  function render(res, notes) {
    const out = $("results"), tabs = R.tables(res), S = res.summary;
    out.append(el("h2", { id: "result" }, "Result"));
    for (const n of notes) out.append(el("p", { class: "warn" }, n));

    const badge = (ok, label) => el("span", { class: "badge " + (ok ? "good" : "bad") }, (ok ? "✓ " : "✗ ") + label);
    out.append(el("div", { class: "badges" }, badge(S.tieBooks, "Books tie out"), badge(S.tie2B, "GSTR-2B ties out"),
      badge(S.bridgeTies, "ITC bridge ties"), el("span", { class: "badge" }, "Tolerance " + rs(res.tol))));

    const tiles = el("div", { class: "tiles" });
    const ORDER = ["matched", "tolerance", "mismatch", "books_only", "gstr2b_only", "duplicate", "flagged"];
    for (const s of ORDER.map((k) => SECTIONS.find((x) => x.key === k))) {
      const n = sectionData(tabs, s).length - 1, itc = itcOf(s.key);
      const hot = n && !["matched", "tolerance"].includes(s.key);
      tiles.append(el("a", { class: "tile" + (hot ? " hot" : "") + (n ? "" : " zero"), href: "#sec-" + s.key,
        onclick: () => { const d = $("sec-" + s.key); if (d) d.open = true; } },
        el("b", {}, String(n)), el("span", {}, s.title), itc !== null && n ? el("small", {}, "ITC " + rs(itc)) : ""));
    }
    out.append(tiles);

    const filter = el("input", { type: "search", id: "filter", placeholder: "Filter by GSTIN, supplier, invoice or voucher", "aria-label": "Filter the lists",
      value: state.filter, oninput: (e) => { state.filter = e.target.value; renderSections(); } });
    out.append(el("div", { class: "actions" },
      el("button", { type: "button", class: "cta", onclick: download }, "Download Excel (8 sheets)"), filter));

    // control totals and ITC bridge
    const T = (t) => [t.rows, t.taxable, t.itc];
    const ctl = el("table", { class: "grid" },
      el("thead", {}, el("tr", {}, el("th", {}, ""), el("th", { class: "n" }, "Rows"), el("th", { class: "n" }, "Taxable value"), el("th", { class: "n" }, "ITC"))),
      el("tbody", {}, [["Purchase register", S.books], ["GSTR-2B (credit notes negative)", S.g2b]].map(([l, t]) =>
        el("tr", {}, el("td", {}, l), ...T(t).map((v, i) => el("td", { class: "n" }, i ? inr.format(v) : String(v)))))));
    const br = el("table", { class: "grid bridge" }, el("tbody", {},
      S.bridge.map(([l, v], i) => el("tr", i ? {} : { class: "strong" }, el("td", {}, l.replace(/^(less|add): /, (m) => m)), el("td", { class: "n" }, inr.format(v)))),
      el("tr", { class: "strong" }, el("td", {}, "= ITC as per GSTR-2B (calculated)"), el("td", { class: "n" }, inr.format(S.bridgeCalc))),
      el("tr", {}, el("td", {}, "ITC as per GSTR-2B (file)"), el("td", { class: "n" }, inr.format(S.g2b.itc)))));
    out.append(el("div", { class: "two" }, el("div", {}, el("h3", {}, "Control totals"), ctl,
      el("p", { class: "note" }, "Every row lands in exactly one list; the Excel summary sheet shows the split by list and by tax head.")),
      el("div", {}, el("h3", {}, "ITC bridge: books to GSTR-2B"), br)));

    out.append(el("div", { id: "sections" }));
    renderSections();
  }

  function renderSections() {
    const box = $("sections");
    if (!box || !state.res) return;
    const tabs = R.tables(state.res), q = state.filter.trim().toLowerCase(), LIMIT = 500;
    box.replaceChildren();
    for (const s of SECTIONS) {
      const rows = sectionData(tabs, s), head = rows[0], all = rows.slice(1);
      const shown = q ? all.filter((r) => r.some((v) => String(v ?? "").toLowerCase().includes(q))) : all;
      const open = (q ? shown.length > 0 : all.length > 0 && !["matched", "tolerance"].includes(s.key)) || ($("sec-" + s.key)?.open ?? false);
      const d = el("details", { id: "sec-" + s.key, class: "sec", open },
        el("summary", {}, s.title, el("span", { class: "count" }, q ? `${shown.length} of ${all.length}` : String(all.length))),
        el("p", { class: "todo" }, s.todo));
      if (shown.length) {
        const numCols = head.map((h, i) => all.some((r) => typeof r[i] === "number"));
        const t = el("table", { class: "list" },
          el("thead", {}, el("tr", {}, head.map((h, i) => el("th", numCols[i] ? { class: "n" } : {}, SHORT[h] || h)))),
          el("tbody", {}, shown.slice(0, LIMIT).map((r) => el("tr", {}, head.map((h, i) => {
            const v = r[i];
            const txt = typeof v === "number" ? (NUMERIC_INT.has(h) ? String(v) : inr.format(v)) : v ?? "";
            return el("td", { class: (numCols[i] ? "n" : "") + (typeof v === "number" && v < 0 ? " neg" : "") + (/^Reason$|^Flag$/.test(h) && v ? " why" : "") }, txt);
          })))));
        d.append(el("div", { class: "scroll" }, t));
        if (shown.length > LIMIT) d.append(el("p", { class: "note" }, `Showing ${LIMIT} of ${shown.length}. The Excel download has all of them.`));
      } else d.append(el("p", { class: "note" }, q ? "No rows match the filter." : "None."));
      box.append(d);
    }
  }

  function download() {
    if (!state.res) return;
    const blob = new Blob([R.writeXlsx(R.tables(state.res))], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const a = el("a", { href: URL.createObjectURL(blob), download: "gstr2b-reconciliation.xlsx" });
    document.body.append(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }

  // ---------- start ----------
  setupDrop("booksZone", "booksFile", false);
  setupDrop("g2bZone", "g2bFiles", true);
  $("runBtn").addEventListener("click", () => { if (state.books && state.g2b.length) load(); });
  $("tol").addEventListener("change", () => { if (state.booksInfo) run(); });
  $("resetBtn").addEventListener("click", () => { state.books = null; state.g2b = []; state.filter = ""; changed(); setStatus("", ""); });
  renderChips();
  if (!("DecompressionStream" in window)) setStatus("This browser is too old to read .xlsx files. Use a current Chrome, Edge, Firefox or Safari, or save the files as .csv.", "err");
})();
