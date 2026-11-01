/* GSTR-2B vs purchase register reconciliation engine.
 * Runs entirely in the browser (or in Node for testing). Reads .xlsx and .csv files itself; no libraries,
 * no network. Logic follows reference_solution.py from the AI for CAs bonus pack (Chapter 14).
 * (c) Chandan Gupta. Free to use. Not professional advice: check every result before relying on it.
 */
(function (root) {
  "use strict";

  // ---------- small helpers ----------
  const r2 = (x) => Math.round((x + (x >= 0 ? 1e-9 : -1e-9)) * 100) / 100;
  const AMTS = ["taxable", "igst", "cgst", "sgst", "cess"];
  const TAX = ["igst", "cgst", "sgst", "cess"];
  const ENT = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };
  const unxml = (s) => s.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (m, e) =>
    e[0] === "#" ? String.fromCodePoint(e[1].toLowerCase() === "x" ? parseInt(e.slice(2), 16) : +e.slice(1))
      : (ENT[e] ?? m));
  const td = new TextDecoder("utf-8");

  function num(v) {
    if (v === null || v === undefined || v === "") return 0;
    if (typeof v === "number") return v;
    let s = String(v).replace(/[₹,\s]/g, "").replace(/^Rs\.?/i, "");
    let neg = false;
    if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
    if (/Cr$/i.test(s)) { neg = true; s = s.slice(0, -2); } else if (/Dr$/i.test(s)) s = s.slice(0, -2);
    const n = parseFloat(s);
    return isNaN(n) ? 0 : (neg ? -n : n);
  }

  const MON = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
  function toDate(v) { // returns ms since epoch (UTC midnight) or null
    if (v === null || v === undefined || v === "") return null;
    if (typeof v === "number") return v > 59 && v < 80000 ? Date.UTC(1899, 11, 30) + Math.round(v) * 864e5 : null;
    const s = String(v).trim();
    let m;
    if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return Date.UTC(+m[1], m[2] - 1, +m[3]);
    if ((m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/))) {
      const y = +m[3] < 100 ? 2000 + +m[3] : +m[3];
      return Date.UTC(y, m[2] - 1, +m[1]); // Indian order: day first
    }
    if ((m = s.match(/^(\d{1,2})[-\s/]([A-Za-z]{3})[A-Za-z]*[-\s/,]*(\d{2,4})$/))) {
      const mo = MON[m[2].toLowerCase()];
      if (mo === undefined) return null;
      const y = +m[3] < 100 ? 2000 + +m[3] : +m[3];
      return Date.UTC(y, mo, +m[1]);
    }
    if (/^\d+(\.\d+)?$/.test(s)) return toDate(+s);
    return null;
  }
  const fmtDate = (ms) => {
    if (ms === null) return "";
    const d = new Date(ms);
    return String(d.getUTCDate()).padStart(2, "0") + "-" + String(d.getUTCMonth() + 1).padStart(2, "0") + "-" + d.getUTCFullYear();
  };

  // ---------- reading .xlsx (zip + XML) ----------
  async function inflateRaw(u8) {
    const stream = new Blob([u8]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  async function unzip(buf) {
    const u8 = new Uint8Array(buf), dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 66000); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error("This is not an .xlsx file (or it is damaged). Save it as .xlsx or .csv and try again.");
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const files = {};
    for (let k = 0; k < count; k++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true), elen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
      const off = dv.getUint32(p + 42, true);
      const name = td.decode(u8.subarray(p + 46, p + 46 + nlen));
      const lnlen = dv.getUint16(off + 26, true), lelen = dv.getUint16(off + 28, true);
      const start = off + 30 + lnlen + lelen;
      files[name] = { method, data: u8.subarray(start, start + csize) };
      p += 46 + nlen + elen + clen;
    }
    return {
      has: (n) => n in files,
      names: () => Object.keys(files),
      async text(n) {
        const f = files[n];
        if (!f) return null;
        if (f.method === 0) return td.decode(f.data);
        if (f.method === 8) return td.decode(await inflateRaw(f.data));
        throw new Error("Unsupported compression in the .xlsx file.");
      },
    };
  }

  function colIndex(ref) {
    const letters = ref.match(/^[A-Z]+/)[0];
    let n = 0;
    for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }

  function parseSheet(xml, shared) {
    const rows = [];
    const rowRe = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g;
    let rm, auto = 0;
    while ((rm = rowRe.exec(xml))) {
      const ra = rm[1].match(/\br="(\d+)"/);
      const ri = ra ? +ra[1] - 1 : auto;
      auto = ri + 1;
      const row = [];
      const body = rm[2] || "";
      const cRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
      let cm, cauto = 0;
      while ((cm = cRe.exec(body))) {
        const attrs = cm[1], inner = cm[2] || "";
        const rr = attrs.match(/\br="([A-Z]+)\d+"/);
        const ci = rr ? colIndex(rr[1]) : cauto;
        cauto = ci + 1;
        const t = (attrs.match(/\bt="(\w+)"/) || [])[1];
        const vm = inner.match(/<v>([\s\S]*?)<\/v>/);
        let val = null;
        if (t === "s") val = vm ? shared[+vm[1]] : null;
        else if (t === "inlineStr") val = unxml((inner.match(/<t[^>]*>([\s\S]*?)<\/t>/g) || []).map((x) => x.replace(/<[^>]+>/g, "")).join(""));
        else if (t === "str" || t === "e") val = vm ? unxml(vm[1]) : null;
        else if (t === "b") val = vm ? vm[1] === "1" : null;
        else val = vm ? +vm[1] : null;
        row[ci] = val;
      }
      for (let i = 0; i < row.length; i++) if (row[i] === undefined) row[i] = null;
      rows[ri] = row;
    }
    for (let i = 0; i < rows.length; i++) if (!rows[i]) rows[i] = [];
    return rows;
  }

  async function readXlsx(buf) {
    const z = await unzip(buf);
    const wb = await z.text("xl/workbook.xml");
    if (!wb) throw new Error("This .xlsx file has no workbook inside. Is it a real Excel file?");
    const rels = (await z.text("xl/_rels/workbook.xml.rels")) || "";
    const relMap = {};
    for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
      const id = (m[0].match(/\bId="([^"]+)"/) || [])[1], tg = (m[0].match(/\bTarget="([^"]+)"/) || [])[1];
      if (id && tg) relMap[id] = tg.startsWith("/") ? tg.slice(1) : "xl/" + tg.replace(/^\.\//, "");
    }
    const ssXml = (await z.text("xl/sharedStrings.xml")) || "";
    const shared = [];
    for (const m of ssXml.matchAll(/<si>([\s\S]*?)<\/si>/g)) {
      shared.push(unxml((m[1].match(/<t[^>]*>([\s\S]*?)<\/t>/g) || []).map((x) => x.replace(/<[^>]+>/g, "")).join("")));
    }
    const sheets = [];
    for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
      const name = unxml((m[0].match(/\bname="([^"]*)"/) || [])[1] || "");
      const rid = (m[0].match(/\br:id="([^"]+)"/) || m[0].match(/\bid="([^"]+)"/) || [])[1];
      const path = relMap[rid];
      const xml = path ? await z.text(path) : null;
      if (xml) sheets.push({ name, rows: parseSheet(xml, shared) });
    }
    return sheets;
  }

  function readCsv(text) {
    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
    const first = text.split(/\r?\n/, 1)[0];
    const delim = (first.match(/;/g) || []).length > (first.match(/,/g) || []).length ? ";" :
      (first.match(/\t/g) || []).length > (first.match(/,/g) || []).length ? "\t" : ",";
    const rows = [];
    let row = [], f = "", q = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (q) {
        if (c === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += c;
      } else if (c === '"') q = true;
      else if (c === delim) { row.push(f); f = ""; }
      else if (c === "\n" || c === "\r") {
        if (c === "\r" && text[i + 1] === "\n") i++;
        row.push(f); rows.push(row); row = []; f = "";
      } else f += c;
    }
    if (f !== "" || row.length) { row.push(f); rows.push(row); }
    return [{ name: "CSV", rows: rows.map((r) => r.map((v) => (v.trim() === "" ? null : v.trim()))) }];
  }

  // GSTR-2B JSON from the portal: a .json file, or a .zip holding one or more of them.
  async function readJsonOrZip(name, buf) {
    if (/\.json$/i.test(name)) return { json: [JSON.parse(td.decode(new Uint8Array(buf)).replace(/^\uFEFF/, ""))] };
    const z = await unzip(buf);
    const names = z.names().filter((n) => /\.json$/i.test(n));
    if (names.length) return { json: await Promise.all(names.map(async (n) => JSON.parse((await z.text(n)).replace(/^\uFEFF/, "")))) };
    if (z.has("xl/workbook.xml")) return null;
    throw new Error(`${name}: no GSTR-2B .json or Excel file inside this .zip.`);
  }

  async function readWorkbook(name, buf) {
    if (/\.json$|\.zip$/i.test(name)) {
      const j = await readJsonOrZip(name, buf);
      if (j) return j;
    }
    if (/\.csv$|\.txt$/i.test(name)) return readCsv(td.decode(new Uint8Array(buf)));
    if (/\.xls$/i.test(name)) throw new Error(`${name}: old .xls format. Open it in Excel and save as .xlsx (or .csv).`);
    return readXlsx(buf);
  }

  // ---------- finding headers and columns ----------
  const cellText = (v) => (v === null || v === undefined ? "" : String(v)).replace(/\s+/g, " ").trim();
  const lc = (v) => cellText(v).toLowerCase();

  const BOOK_FIELDS = [ // order matters: specific fields claim their column first
    ["gstin", "Supplier GSTIN", true, /gstin|gst\s*(no|number|reg)/],
    ["taxable", "Taxable value", true, /taxable|assessable/],
    ["igst", "IGST", true, /igst|integrated/],
    ["cgst", "CGST", true, /cgst|central/],
    ["sgst", "SGST / UTGST", true, /sgst|utgst|state/],
    ["cess", "Cess", false, /cess/],
    ["date", "Invoice date", false, [/(invoice|bill|supplier|doc).*date/, /^date$/, /date/]],
    ["inv", "Invoice no", true, /(invoice|bill|supplier inv|supplier ref|ref).*(no|num|#)|^invoice$|^bill$|supplier invoice/],
    ["voucher", "Voucher no", false, [/(voucher|vch).*(no|num|#)/, /voucher|vch/]],
    ["doc", "Document type", false, /doc.*type|voucher type|vch type|^type$|document/],
    ["ledger", "Ledger / expense head", false, /ledger|head|account|expense|purchase a\/c/],
    ["name", "Supplier name", false, /supplier|party|vendor|particulars|name/],
  ];

  function findBooksHeader(rows) {
    for (let i = 0; i < Math.min(40, rows.length); i++) {
      const cells = (rows[i] || []).map(lc);
      if (cells.some((c) => /gstin|gst\s*(no|number|reg)/.test(c)) && cells.some((c) => /taxable|assessable/.test(c))) return i;
    }
    return -1;
  }

  function autoMap(headers, fields) {
    const map = {}, used = new Set();
    for (const [key, , , res] of fields) {
      outer: for (const re of [].concat(res)) { // several patterns: the first, most specific, wins
        for (let i = 0; i < headers.length; i++) {
          const h = lc(headers[i]);
          if (!h || used.has(i)) continue;
          if ((key === "inv" || key === "voucher") && /date/.test(h)) continue;
          if (key === "name" && /gstin/.test(h)) continue;
          if (re.test(h)) { map[key] = i; used.add(i); break outer; }
        }
      }
    }
    return map;
  }

  // Inspect the books file: which sheet, header row and column mapping. The UI may override `map`.
  function inspectBooks(sheets) {
    for (const s of sheets) {
      const h = findBooksHeader(s.rows);
      if (h >= 0) {
        const headers = s.rows[h].map(cellText);
        return { sheet: s.name, headerRow: h, headers, map: autoMap(headers, BOOK_FIELDS), rows: s.rows };
      }
    }
    throw new Error("Couldn't find the header row in the purchase register: it needs columns for supplier GSTIN and taxable value. Check the file, or save it as .csv with one header row.");
  }

  function buildBooks(info) {
    const { rows, headerRow, map } = info;
    const missing = BOOK_FIELDS.filter(([k, , req]) => req && map[k] === undefined).map(([, label]) => label);
    if (missing.length) throw new Error("Purchase register: choose the column for " + missing.join(", ") + ".");
    const get = (r, k) => (map[k] === undefined ? null : r[map[k]]);
    const out = [];
    let flipped = 0;
    for (let i = headerRow + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const gstin = cellText(get(r, "gstin")).toUpperCase(), inv = cellText(get(r, "inv"));
      if (!gstin && !inv) continue; // blank and total lines
      if (!gstin && /total/i.test(r.map(cellText).join(" "))) continue;
      const docText = lc(get(r, "doc"));
      const rec = {
        ref: cellText(get(r, "voucher")) || `Row ${i + 1}`, excelRow: i + 1, gstin,
        name: cellText(get(r, "name")), inv, date: toDate(get(r, "date")),
        ledger: cellText(get(r, "ledger")),
      };
      for (const a of AMTS) rec[a] = r2(num(get(r, a)));
      const isCredit = /credit/.test(docText) || (map.doc === undefined && rec.taxable < 0);
      rec.doc = isCredit ? "CN" : /debit/.test(docText) ? "DN" : "INV";
      if (rec.doc === "CN" && rec.taxable > 0) { for (const a of AMTS) rec[a] = -rec[a]; flipped++; }
      out.push(rec);
    }
    return { rows: out, flipped };
  }

  const G2B_COLS = {
    gstin: /gstin of supplier/, name: /trade\/legal name|trade name|legal name/,
    inv: /invoice number|note number/, date: /invoice date|note date/, noteType: /note type/,
    taxable: /taxable value/, igst: /integrated tax/, cgst: /central tax/, sgst: /state\/ut tax|state tax/,
    cess: /^cess/, period: /gstr-1\/iff\/gstr-5 period|^period/, itcAvail: /itc availability/,
  };

  // ---------- GSTR-2B JSON (portal download) ----------
  const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const prd = (p) => { const m = String(p || "").match(/^(\d{2})(\d{4})$/); return m ? `${MONTHS[+m[1] - 1]}'${m[2].slice(2)}` : String(p || ""); };
  const lower = (o) => { const r = {}; for (const [k, v] of Object.entries(o || {})) r[k.toLowerCase()] = v; return r; };

  function json2B(file, obj, multi) {
    const out = [], notes = [];
    let root = lower(obj);
    if (root.data && typeof root.data === "object") root = lower(root.data);
    const dd = lower(root.docdata || root);
    const tag = multi ? file + " " : "";
    for (const [sec, listKey, numKey, isNote] of [["b2b", "inv", "inum", false], ["cdnr", "nt", "ntnum", true]]) {
      (dd[sec] || []).forEach((sup, si) => {
        const S = lower(sup);
        (S[listKey] || []).forEach((doc, di) => {
          const D = lower(doc);
          let sign = 1, kind = "INV";
          if (isNote) { const t = String(D.typ || D.ntty || "").toUpperCase(); if (t.startsWith("C")) { sign = -1; kind = "CN"; } else kind = "DN"; }
          const avail = String(D.itcavl || "").toUpperCase();
          const rec = {
            ref: `${tag}JSON ${sec}[${si + 1}].${listKey}[${di + 1}]`, excelRow: "", gstin: String(S.ctin || "").trim().toUpperCase(),
            name: String(S.trdnm || ""), inv: String(D[numKey] ?? D.inum ?? D.nt_num ?? "").trim(), date: toDate(D.dt || D.idt || D.nt_dt),
            doc: kind, period: prd(S.supprd || root.rtnprd), itcAvail: avail === "Y" ? "Yes" : avail === "N" ? "No" : avail,
          };
          const amt = { taxable: D.txval, igst: D.igst ?? D.iamt, cgst: D.cgst ?? D.camt, sgst: D.sgst ?? D.samt, cess: D.cess ?? D.csamt };
          for (const a of AMTS) rec[a] = r2(num(amt[a]) * sign);
          out.push(rec);
        });
      });
    }
    for (const k of ["b2ba", "cdnra"]) if ((dd[k] || []).length) notes.push(`${file}: ${k.toUpperCase()} (amendments) was not used. Review amendments separately.`);
    if (!out.length) notes.push(`${file}: no B2B or CDNR documents found in this JSON. Is it the GSTR-2B JSON from the GST portal?`);
    return { out, notes };
  }

  function build2B(sheetsList) { // sheetsList: [{file, sheets}] where sheets is a list of sheets, or {json: [objects]}
    const out = [], notes = [];
    for (const { file, sheets } of sheetsList) {
      if (sheets && sheets.json) {
        for (const obj of sheets.json) {
          const r = json2B(file, obj, sheetsList.length > 1 || sheets.json.length > 1);
          out.push(...r.out); notes.push(...r.notes);
        }
        continue;
      }
      let used = 0;
      for (const s of sheets) {
        const nm = s.name.trim().toUpperCase().replace(/\s+/g, "");
        let kind = nm === "B2B" ? "INV" : /^B2B-?CDNR$/.test(nm) ? "NOTE" : null;
        if (!kind && sheets.length === 1) { // a single sheet or CSV saved from the 2B: tell invoices from notes by the headers
          const head = s.rows.slice(0, 15).map((r) => (r || []).map(lc).join("|")).join("|");
          if (/gstin of supplier/.test(head)) kind = /note number|note type/.test(head) ? "NOTE" : "INV";
        }
        if (!kind) {
          if (/^B2BA$|^B2B-?CDNRA$/.test(nm) && s.rows.length > 7) notes.push(`${file}: sheet ${s.name} (amendments) was not used. Review amendments separately.`);
          continue;
        }
        let h = -1;
        for (let i = 0; i < Math.min(15, s.rows.length); i++) if ((s.rows[i] || []).some((c) => /gstin of supplier/i.test(cellText(c)))) h = i;
        // h is the LAST row mentioning "GSTIN of supplier": the column-name row under the merged group titles
        if (h < 0) { notes.push(`${file}: sheet ${s.name} has no "GSTIN of supplier" header; skipped.`); continue; }
        const top = h > 0 ? s.rows[h - 1] || [] : [];
        const sub = s.rows[h] || [];
        const headers = [];
        for (let i = 0; i < Math.max(top.length, sub.length); i++) headers.push(cellText(sub[i]) || cellText(top[i]));
        const map = {};
        for (const [k, re] of Object.entries(G2B_COLS)) {
          const i = headers.findIndex((x) => re.test(x.toLowerCase()));
          if (i >= 0) map[k] = i;
        }
        for (const k of ["gstin", "inv", "taxable"]) if (map[k] === undefined) throw new Error(`${file}, sheet ${s.name}: couldn't find the ${k === "inv" ? "invoice/note number" : k} column.`);
        const get = (r, k) => (map[k] === undefined ? null : r[map[k]]);
        for (let i = h + 1; i < s.rows.length; i++) {
          const r = s.rows[i] || [];
          const gstin = cellText(get(r, "gstin")).toUpperCase();
          if (!gstin || /total/i.test(gstin)) continue;
          let sign = 1, doc = "INV";
          if (kind === "NOTE") {
            const t = lc(get(r, "noteType"));
            if (/credit/.test(t)) { sign = -1; doc = "CN"; } else doc = "DN";
          }
          const rec = {
            ref: `${sheetsList.length > 1 ? file + " " : ""}${s.name}!row${i + 1}`, excelRow: i + 1, gstin,
            name: cellText(get(r, "name")), inv: cellText(get(r, "inv")), date: toDate(get(r, "date")), doc,
            period: cellText(get(r, "period")), itcAvail: cellText(get(r, "itcAvail")),
          };
          for (const a of AMTS) rec[a] = r2(num(get(r, a)) * sign);
          out.push(rec);
          used++;
        }
      }
      if (!used) notes.push(`${file}: no B2B or B2B-CDNR rows found. Is this the GSTR-2B Excel or JSON download from the GST portal?`);
    }
    return { rows: out, notes };
  }

  // ---------- GSTIN check digit ----------
  const CHARS = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  function gstinValid(g) {
    g = String(g).trim().toUpperCase();
    if (!/^\d{2}[A-Z]{5}\d{4}[A-Z][0-9A-Z]Z[0-9A-Z]$/.test(g)) return false;
    let total = 0;
    for (let i = 0; i < 14; i++) {
      const p = CHARS.indexOf(g[i]) * (i % 2 ? 2 : 1);
      total += Math.floor(p / 36) + (p % 36);
    }
    return CHARS[(36 - (total % 36)) % 36] === g[14];
  }

  function normInv(inv) {
    const s = String(inv).toUpperCase().replace(/20\d\d-\d\d|\d\d-\d\d/g, "");
    return s.replace(/\D/g, "").replace(/^0+/, "") || String(inv).trim().toUpperCase();
  }

  const BLOCKED_WORDS = [
    ["motor car", "Motor vehicle for transport of persons: s.17(5)(a), check the exceptions"],
    ["motor vehicle", "Motor vehicle for transport of persons: s.17(5)(a), check the exceptions"],
    ["club", "Membership of a club: s.17(5)(b)"],
    ["gym", "Health and fitness centre: s.17(5)(b)"],
    ["food", "Food and beverages: s.17(5)(b), check the exceptions"],
    ["catering", "Outdoor catering: s.17(5)(b), check the exceptions"],
    ["staff welfare", "May include food, health or travel benefits: s.17(5)(b)"],
    ["gift", "Gifts: s.17(5)(h)"],
    ["insurance", "Life or health insurance, or motor insurance: s.17(5)(b) or (ab)"],
    ["construction", "Works contract or construction of immovable property: s.17(5)(c) or (d)"],
  ];

  // ---------- the reconciliation ----------
  function reconcile(books, g2b, tol) {
    tol = tol === undefined ? 1 : tol;
    for (const d of [books, g2b]) for (const x of d) {
      x.core = normInv(x.inv);
      x.itc = r2(TAX.reduce((s, t) => s + x[t], 0));
      x.status = null;
    }
    // possible duplicates in books: same GSTIN, document type and cleaned number; keep the first
    const seen = new Map();
    for (const b of books) {
      const k = b.gstin + "|" + b.doc + "|" + b.core;
      if (seen.has(k)) { b.status = "duplicate"; b.dupOf = seen.get(k).ref; } else seen.set(k, b);
    }
    const pairs = [];
    function match(key, basis, extra) {
      const idx = new Map();
      for (const g of g2b) if (g.status === null) {
        const k = key(g);
        if (!idx.has(k)) idx.set(k, []);
        idx.get(k).push(g);
      }
      for (const b of books) {
        if (b.status !== null) continue;
        for (const g of idx.get(key(b)) || []) {
          if (g.status === null && (!extra || extra(b, g))) {
            b.status = g.status = "paired";
            pairs.push([b, g, basis]);
            break;
          }
        }
      }
    }
    match((d) => d.gstin + "|" + d.doc + "|" + d.inv.toUpperCase(), "1 GSTIN + invoice no as written");
    match((d) => d.gstin + "|" + d.doc + "|" + d.core, "2 GSTIN + cleaned invoice no");
    match((d) => d.doc + "|" + d.core, "3 cleaned invoice no + value, GSTIN differs", (b, g) =>
      Math.abs(b.taxable - g.taxable) <= tol && b.date !== null && g.date !== null &&
      Math.abs(b.date - g.date) <= 7 * 864e5);

    const paired = [];
    for (const [b, g, basis] of pairs) {
      const diffs = {};
      for (const a of AMTS) diffs[a] = r2(g[a] - b[a]);
      const worst = Math.max(...AMTS.map((a) => Math.abs(diffs[a])));
      const reasons = [];
      if (b.gstin !== g.gstin) reasons.push("GSTIN differs");
      if (Math.abs(diffs.taxable) > tol) reasons.push("taxable value differs");
      if (TAX.some((t) => Math.abs(diffs[t]) > tol)) reasons.push("tax differs");
      let status;
      if (reasons.length) status = "mismatch";
      else if (worst < 0.005) status = "matched";
      else { status = "tolerance"; reasons.push(`largest difference Rs ${worst.toFixed(2)}`); }
      b.status = g.status = status;
      paired.push({ status, basis, reason: reasons.join("; "), b, g, diffs, itcDiff: r2(g.itc - b.itc) });
    }
    for (const b of books) if (b.status === null) b.status = "books_only";
    for (const g of g2b) if (g.status === null) g.status = "gstr2b_only";

    const flagged = [];
    for (const b of books) {
      if (!gstinValid(b.gstin)) flagged.push({ flag: "GSTIN fails format or check digit", source: "Books", rec: b,
        note: b.gstin ? "Correct the vendor master; see the 2B GSTIN in mismatched values if a probable pair was found" : "No GSTIN in the books: no ITC unless the supplier is registered" });
      const led = b.ledger.toLowerCase();
      for (const [w, why] of BLOCKED_WORDS) if (led.includes(w)) {
        flagged.push({ flag: "Possibly blocked credit", source: "Books", rec: b, note: why + ". Matched in 2B does not make it eligible." });
        break;
      }
    }
    for (const g of g2b) {
      if (!gstinValid(g.gstin)) flagged.push({ flag: "GSTIN fails format or check digit", source: "2B", rec: g, note: "" });
      if (g.itcAvail.trim().toLowerCase() === "no") flagged.push({ flag: "2B says ITC not available", source: "2B", rec: g, note: "Read the Reason column in 2B" });
    }

    const totals = (rows) => {
      const t = { rows: rows.length };
      for (const a of AMTS.concat(["itc"])) t[a] = r2(rows.reduce((s, x) => s + x[a], 0));
      return t;
    };
    const by = (rows, st) => rows.filter((x) => x.status === st);
    const bookParts = ["matched", "tolerance", "mismatch", "books_only", "duplicate"].map((st) => [st, totals(by(books, st))]);
    const g2bParts = ["matched", "tolerance", "mismatch", "gstr2b_only"].map((st) => [st, totals(by(g2b, st))]);
    const tie = (whole, parts) => AMTS.concat(["itc", "rows"]).every((a) =>
      Math.abs(whole[a] - parts.reduce((s, [, t]) => s + t[a], 0)) < 0.01);
    const tb = totals(books), tg = totals(g2b);
    const sumDiff = (st) => r2(paired.filter((p) => p.status === st).reduce((s, p) => s + p.itcDiff, 0));
    const bridge = [
      ["ITC as per books", tb.itc],
      ["less: in books, not in 2B", -totals(by(books, "books_only")).itc],
      ["less: possible duplicates in books", -totals(by(books, "duplicate")).itc],
      ["add: in 2B, not in books", totals(by(g2b, "gstr2b_only")).itc],
      ["add: differences on mismatched pairs (2B less books)", sumDiff("mismatch")],
      ["add: differences within tolerance (2B less books)", sumDiff("tolerance")],
    ];
    const calc = r2(bridge.reduce((s, [, v]) => s + v, 0));
    return {
      tol, paired, flagged, books, g2b,
      summary: { books: tb, g2b: tg, bookParts, g2bParts, tieBooks: tie(tb, bookParts), tie2B: tie(tg, g2bParts),
        bridge, bridgeCalc: calc, bridgeTies: Math.abs(calc - tg.itc) < 0.01 },
    };
  }

  // ---------- output tables (used for screen and for the Excel download) ----------
  const LABELS = { matched: "Matched", tolerance: "Matched within tolerance", mismatch: "Mismatched values",
    books_only: "In books, not in 2B", gstr2b_only: "In 2B, not in books", duplicate: "Possible duplicate in books" };

  function tables(res) {
    const S = res.summary, T = ["rows", "taxable", "igst", "cgst", "sgst", "cess", "itc"];
    const line = (label, t) => [label].concat(T.map((k) => t[k]));
    const summary = [["Line", "Rows", "Taxable", "IGST", "CGST", "SGST", "Cess", "Total ITC"],
      line("BOOKS: purchase register, all rows", S.books)];
    for (const [st, t] of S.bookParts) summary.push(line("  " + LABELS[st], t));
    summary.push(["  ties to the purchase register: " + (S.tieBooks ? "YES" : "NO")], [],
      line("GSTR-2B: B2B plus B2B-CDNR (credit notes negative)", S.g2b));
    for (const [st, t] of S.g2bParts) summary.push(line("  " + LABELS[st], t));
    summary.push(["  ties to GSTR-2B: " + (S.tie2B ? "YES" : "NO")], [], ["ITC BRIDGE: books to GSTR-2B"]);
    for (const [lab, v] of S.bridge) summary.push(["  " + lab, "", "", "", "", "", "", v]);
    summary.push(["  = ITC as per GSTR-2B (calculated)", "", "", "", "", "", "", S.bridgeCalc],
      ["  ITC as per GSTR-2B (file): ties " + (S.bridgeTies ? "YES" : "NO"), "", "", "", "", "", "", S.g2b.itc],
      [], ["Tolerance used: Rs " + res.tol.toFixed(2) + " per amount column (a working choice, not a rule in the law)."]);

    const pairHead = ["Status", "Match basis", "Reason", "Books voucher", "Books row", "2B ref", "Supplier GSTIN (books)",
      "Supplier GSTIN (2B)", "Supplier", "Doc", "Invoice no (books)", "Invoice no (2B)", "Invoice date", "2B period",
      "Ledger head", "Taxable (books)", "Taxable (2B)", "ITC (books)", "ITC (2B)", "ITC difference (2B less books)",
      "Taxable diff", "IGST diff", "CGST diff", "SGST diff", "Cess diff"];
    const pairRows = (st) => [pairHead].concat(res.paired.filter((p) => p.status === st).map((p) => [
      LABELS[p.status], p.basis, p.reason, p.b.ref, p.b.excelRow, p.g.ref, p.b.gstin, p.g.gstin, p.g.name || p.b.name,
      p.b.doc, p.b.inv, p.g.inv, fmtDate(p.g.date), p.g.period, p.b.ledger, p.b.taxable, p.g.taxable, p.b.itc, p.g.itc,
      p.itcDiff].concat(AMTS.map((a) => p.diffs[a]))));
    const sideHead = ["Ref", "Row", "Supplier GSTIN", "Supplier", "Invoice no", "Invoice date", "Doc", "Ledger head / 2B period",
      "Taxable", "IGST", "CGST", "SGST", "Cess", "ITC"];
    const side = (rows, extra) => [sideHead.concat(extra ? [extra[0]] : [])].concat(rows.map((x) => [x.ref, x.excelRow,
      x.gstin, x.name, x.inv, fmtDate(x.date), x.doc, x.ledger !== undefined ? x.ledger : x.period].concat(
      AMTS.map((a) => x[a]), [x.itc], extra ? [extra[1](x)] : [])));
    const flagged = [["Flag", "Source", "Ref", "Supplier", "GSTIN", "Invoice no", "Ledger head", "ITC", "Status", "Note"]]
      .concat(res.flagged.map((f) => [f.flag, f.source, f.rec.ref, f.rec.name, f.rec.gstin, f.rec.inv, f.rec.ledger || "",
        f.rec.itc, LABELS[f.rec.status] || f.rec.status, f.note]));
    return [
      ["summary", summary],
      ["matched", pairRows("matched")],
      ["matched-within-tolerance", pairRows("tolerance")],
      ["mismatched values", pairRows("mismatch")],
      ["in-books-not-in-2B", side(res.books.filter((x) => x.status === "books_only"))],
      ["in-2B-not-in-books", side(res.g2b.filter((x) => x.status === "gstr2b_only"))],
      ["possible-duplicates", side(res.books.filter((x) => x.status === "duplicate"), ["Same as voucher", (x) => x.dupOf])],
      ["flagged-for-review", flagged],
    ];
  }

  // ---------- writing .xlsx (stored zip, inline strings) ----------
  const CRC = (() => { const t = []; for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = (u8) => { let c = 0xffffffff; for (let i = 0; i < u8.length; i++) c = CRC[(c ^ u8[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  const xesc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]))
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "");
  const colName = (i) => { let s = ""; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };

  function zipStore(files) { // files: [[name, string]]
    const te = new TextEncoder(), parts = [], central = [];
    let off = 0;
    for (const [name, text] of files) {
      const nb = te.encode(name), data = te.encode(text), crc = crc32(data);
      const h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true);
      h.setUint32(14, crc, true); h.setUint32(18, data.length, true); h.setUint32(22, data.length, true);
      h.setUint16(26, nb.length, true);
      parts.push(new Uint8Array(h.buffer), nb, data);
      const c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true);
      c.setUint32(16, crc, true); c.setUint32(20, data.length, true); c.setUint32(24, data.length, true);
      c.setUint16(28, nb.length, true); c.setUint32(42, off, true);
      central.push(new Uint8Array(c.buffer), nb);
      off += 30 + nb.length + data.length;
    }
    const cdSize = central.reduce((s, x) => s + x.length, 0);
    const e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true);
    e.setUint32(12, cdSize, true); e.setUint32(16, off, true);
    const all = parts.concat(central, [new Uint8Array(e.buffer)]);
    const out = new Uint8Array(all.reduce((s, x) => s + x.length, 0));
    let p = 0;
    for (const x of all) { out.set(x, p); p += x.length; }
    return out;
  }

  function writeXlsx(sheetTables) {
    const files = [];
    const sheetXml = (rows) => {
      const widths = [];
      for (const r of rows.slice(0, 60)) (r || []).forEach((v, i) => { widths[i] = Math.min(48, Math.max(widths[i] || 10, String(v ?? "").length + 2)); });
      const cols = widths.length ? "<cols>" + widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("") + "</cols>" : "";
      const body = rows.map((r, ri) => `<row r="${ri + 1}">` + (r || []).map((v, ci) => {
        const ref = colName(ci) + (ri + 1);
        if (v === null || v === undefined || v === "") return "";
        if (typeof v === "number" && isFinite(v)) return `<c r="${ref}"><v>${v}</v></c>`;
        return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xesc(v)}</t></is></c>`;
      }).join("") + "</row>").join("");
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>${cols}<sheetData>${body}</sheetData></worksheet>`;
    };
    files.push(["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      sheetTables.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("") + "</Types>"]);
    files.push(["_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`]);
    files.push(["xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>` +
      sheetTables.map(([n], i) => `<sheet name="${xesc(n.slice(0, 31))}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("") + "</sheets></workbook>"]);
    files.push(["xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      sheetTables.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("") + "</Relationships>"]);
    sheetTables.forEach(([, rows], i) => files.push([`xl/worksheets/sheet${i + 1}.xml`, sheetXml(rows)]));
    return zipStore(files);
  }

  const API = { readWorkbook, inspectBooks, buildBooks, build2B, reconcile, tables, writeXlsx, gstinValid, normInv,
    toDate, fmtDate, BOOK_FIELDS, LABELS };
  if (typeof module !== "undefined" && module.exports) module.exports = API;
  else root.Recon = API;
})(typeof self !== "undefined" ? self : this);
