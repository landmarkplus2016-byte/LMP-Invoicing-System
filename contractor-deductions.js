/* ============================================================
   LMP Invoicing System — Contractor Deductions (window.ConDeductions)
   Shared by all four Contractor Invoices sub-tabs.

   Sources:
     • Expenses  — Analysis Tool 2026.xlsm › "Total Expense"
     • Ledger    — Invoice-Deductions 2026.xlsm › "Deduction" (rows already deducted)

   For one contractor's draft rows:
     1. Site ID + Job Code combos of the draft
     2. every expense row of those combos (ticked categories, cost ≠ 0)
     3. minus rows already in the ledger — matched row by row, count-aware:
        pass 1 same date + site + job code + cost + category + sub category,
        pass 2 the same without the date (the ledger has a few mistyped dates)
     4. flags: the combo also has tasks of another contractor in the tracking,
        or the expense row names another contractor
   Never writes to the source workbooks.
   ============================================================ */
(function () {
'use strict';

const $ = id => document.getElementById(id);
const CAT_KEY = 'lmp-con-ded-cats-off';     // localStorage: categories the user unticked

let _exp = null;      // { rows, byCombo: Map, categories: [{name,count}], file }
let _led = null;      // { rows, file }
// Trucks are paid weekly now, so they are not deducted unless the user ticks them.
const DEFAULT_OFF = ['Trucks'];
let _catsOff = new Set(DEFAULT_OFF);
try {
  const saved = localStorage.getItem(CAT_KEY);
  if (saved != null) _catsOff = new Set(JSON.parse(saved));
} catch { _catsOff = new Set(DEFAULT_OFF); }

// ---------------------------------------------------------------------------
// Keys
// ---------------------------------------------------------------------------
function siteKey(v) {
  let s = String(v ?? '').trim().toUpperCase().replace(/\s+/g, '');
  if (/^\d+(\.0+)?$/.test(s)) s = String(parseInt(s, 10));      // 0782 / 782.0 → 782
  return s;
}
function jcKey(v) { return String(v ?? '').trim().toUpperCase().replace(/\s+/g, ''); }
function comboKey(site, jc) { return siteKey(site) + '|' + jcKey(jc); }
function txtKey(v) { return String(v ?? '').trim().toLowerCase().replace(/\s+/g, ' '); }
function num(v) {
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v ?? '').replace(/[^0-9.-]/g, ''));
  return isNaN(n) ? 0 : n;
}
// Excel serial day (integer) from a SheetJS Date (local midnight, may be a few s early) or a number
function daySerial(v) {
  if (v instanceof Date && !isNaN(v)) return Math.round((v.getTime() - v.getTimezoneOffset() * 60000) / 86400000 + 25569);
  if (typeof v === 'number' && v > 20000 && v < 80000) return Math.round(v);
  if (typeof v === 'string' && v.trim()) {
    const d = new Date(v);
    if (!isNaN(d)) return Math.round((d.getTime() - d.getTimezoneOffset() * 60000) / 86400000 + 25569);
  }
  return null;
}
function serialToDate(s) { return s == null ? null : new Date(Date.UTC(1899, 11, 30) + s * 86400000); }
function costKey(v) { return (Math.round(num(v) * 100) / 100).toFixed(2); }
function catName(v) {
  const s = String(v ?? '').trim();
  if (!s) return '(blank)';
  return s.replace(/\s+/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
}

// ---------------------------------------------------------------------------
// Reading
// ---------------------------------------------------------------------------
function findSheet(wb, name) {
  const n = wb.SheetNames.find(s => s.trim().toLowerCase() === name.toLowerCase());
  return n ? wb.Sheets[n] : null;
}
function findHeader(rows) {
  for (let i = 0; i < Math.min(40, rows.length); i++) {
    const t = (rows[i] || []).map(c => String(c ?? '').trim().toLowerCase());
    if (t.includes('site id') && t.includes('job code') && t.some(x => x.startsWith('cost'))) return i;
  }
  return -1;
}
function colMap(header) {
  const t = header.map(c => String(c ?? '').trim().toLowerCase());
  const exact = (...names) => t.findIndex(x => names.includes(x));
  const has   = (fn) => t.findIndex(fn);
  return {
    date:  exact('date'),
    site:  exact('site id'),
    jc:    exact('job code'),
    cost:  has(x => x.startsWith('cost')),
    cat:   exact('category'),
    sub:   exact('sub category', 'sub-category', 'subcategory'),
    coord: exact('coordinator'),
    name:  exact('name'),
    desc:  exact('item description'),
    comment: exact('comment'),
    contractor: exact('contractor'),
    vf:    has(x => x.startsWith('vf inv')),
    conInv: has(x => x.includes('contractor') && x.includes('inv')),
  };
}

function readExpenses(wb, fileName) {
  const sh = findSheet(wb, 'Total Expense');
  if (!sh) throw new Error('Expenses file: sheet "Total Expense" not found (it has: ' + wb.SheetNames.slice(0, 8).join(', ') + ').');
  const all = XLSX.utils.sheet_to_json(sh, { header: 1, defval: null, raw: true });
  const h = findHeader(all);
  if (h < 0) throw new Error('Expenses file: header row with "Site ID", "Job Code" and "Cost/Site" not found in "Total Expense".');
  const c = colMap(all[h]);
  const rows = [], byCombo = new Map(), cats = new Map();
  for (let i = h + 1; i < all.length; i++) {
    const r = all[i];
    if (!r) continue;
    const site = r[c.site], jc = r[c.jc];
    if (site == null || String(site).trim() === '' || jc == null || String(jc).trim() === '') continue;
    const cost = num(r[c.cost]);
    if (!cost) continue;
    const e = {
      excelRow: i + 1,
      day: c.date >= 0 ? daySerial(r[c.date]) : null,
      site: String(site).trim(), jc: String(jc).trim(),
      combo: comboKey(site, jc),
      cost,
      cat: c.cat >= 0 ? String(r[c.cat] ?? '').trim() : '',
      sub: c.sub >= 0 ? String(r[c.sub] ?? '').trim() : '',
      coord: c.coord >= 0 ? String(r[c.coord] ?? '').trim() : '',
      comment: c.comment >= 0 ? String(r[c.comment] ?? '').trim() : '',
      desc: c.desc >= 0 ? String(r[c.desc] ?? '').trim() : '',
      contractor: c.contractor >= 0 ? String(r[c.contractor] ?? '').trim() : '',
    };
    e.catName = catName(e.cat);
    rows.push(e);
    if (!byCombo.has(e.combo)) byCombo.set(e.combo, []);
    byCombo.get(e.combo).push(e);
    cats.set(e.catName, (cats.get(e.catName) || 0) + 1);
  }
  if (!rows.length) throw new Error('Expenses file: no expense rows found in "Total Expense".');
  const categories = [...cats.entries()].sort((a, b) => b[1] - a[1]).map(([name, count]) => ({ name, count }));
  return { rows, byCombo, categories, file: fileName };
}

function readLedger(wb, fileName) {
  const sh = findSheet(wb, 'Deduction');
  if (!sh) throw new Error('Ledger file: sheet "Deduction" not found (it has: ' + wb.SheetNames.slice(0, 8).join(', ') + ').');
  const all = XLSX.utils.sheet_to_json(sh, { header: 1, defval: null, raw: true });
  const h = findHeader(all);
  if (h < 0) throw new Error('Ledger file: header row with "Site ID", "Job Code" and "Cost/Site" not found in "Deduction".');
  const c = colMap(all[h]);
  const rows = [];
  for (let i = h + 1; i < all.length; i++) {
    const r = all[i];
    if (!r) continue;
    const site = r[c.site], jc = r[c.jc];
    if (site == null || String(site).trim() === '' || jc == null) continue;
    rows.push({
      excelRow: i + 1,
      day: c.date >= 0 ? daySerial(r[c.date]) : null,
      combo: comboKey(site, jc),
      cost: num(r[c.cost]),
      cat: c.cat >= 0 ? String(r[c.cat] ?? '').trim() : '',
      sub: c.sub >= 0 ? String(r[c.sub] ?? '').trim() : '',
      contractor: c.contractor >= 0 ? String(r[c.contractor] ?? '').trim() : '',
      vf: c.vf >= 0 ? String(r[c.vf] ?? '').trim() : '',
      conInv: c.conInv >= 0 ? String(r[c.conInv] ?? '').trim() : '',
    });
  }
  return { rows, file: fileName };
}

const fpFull  = r => [r.day ?? '', r.combo, costKey(r.cost), txtKey(r.cat), txtKey(r.sub)].join('|');
const fpLoose = r => [r.combo, costKey(r.cost), txtKey(r.cat), txtKey(r.sub)].join('|');

// ---------------------------------------------------------------------------
// Compute deductions for one contractor
//   rows   — draft rows ({ jobCode, siteId, … })
//   opts.owners — Map(comboKey → Set(contractor names)) from the whole tracking
//   opts.normalize — contractor-name normaliser (same as the drafts use)
// ---------------------------------------------------------------------------
function compute(contractor, rows, opts) {
  opts = opts || {};
  if (!_exp) return null;
  const combos = new Map();                      // key → { site, jc }
  rows.forEach(r => {
    if (!String(r.siteId ?? '').trim() || !String(r.jobCode ?? '').trim()) return;
    const k = comboKey(r.siteId, r.jobCode);
    if (!combos.has(k)) combos.set(k, { site: r.siteId, jc: r.jobCode });
  });

  // candidate expense rows
  const cand = [];
  let catSkipped = 0, catSkippedAmt = 0;
  for (const k of combos.keys()) {
    for (const e of (_exp.byCombo.get(k) || [])) {
      if (_catsOff.has(e.catName)) { catSkipped++; catSkippedAmt += e.cost; continue; }
      cand.push(e);
    }
  }
  cand.sort((a, b) => (a.combo < b.combo ? -1 : a.combo > b.combo ? 1 : 0) || ((a.day ?? 0) - (b.day ?? 0)) || a.excelRow - b.excelRow);

  // ledger matching (count-aware, only ledger rows of these combos)
  const ledFull = new Map(), ledLoose = new Map();
  if (_led) {
    _led.rows.forEach((l, idx) => {
      if (!combos.has(l.combo)) return;
      const f = fpFull(l), g = fpLoose(l);
      if (!ledFull.has(f)) ledFull.set(f, []);
      ledFull.get(f).push(idx);
      if (!ledLoose.has(g)) ledLoose.set(g, []);
      ledLoose.get(g).push(idx);
    });
  }
  const used = new Set();
  const take = (list) => { if (!list) return -1; for (const i of list) if (!used.has(i)) { used.add(i); return i; } return -1; };
  const match = new Map();                       // expense → { led, loose }
  cand.forEach(e => { const i = take(ledFull.get(fpFull(e))); if (i >= 0) match.set(e, { led: _led.rows[i], loose: false }); });
  cand.forEach(e => { if (match.has(e)) return; const i = take(ledLoose.get(fpLoose(e))); if (i >= 0) match.set(e, { led: _led.rows[i], loose: true }); });

  const norm = opts.normalize || (s => s);
  const me = norm(contractor);
  const deduct = [], already = [];
  for (const e of cand) {
    const m = match.get(e);
    if (m) { already.push({ e, led: m.led, loose: m.loose }); continue; }
    const flags = [];
    const owners = opts.owners && opts.owners.get(e.combo);
    if (owners) {
      const others = [...owners].filter(o => o && o !== me);
      if (others.length) flags.push('Site + Job Code also has tasks of ' + others.join(', '));
    }
    if (e.contractor) {
      const ec = norm(e.contractor);
      if (ec && ec !== me) flags.push('Expense row says contractor: ' + e.contractor);
    }
    deduct.push({ e, flags });
  }
  const sum = list => list.reduce((s, x) => s + x.e.cost, 0);
  return {
    contractor, combos: combos.size,
    deduct, already,
    total: sum(deduct), alreadyTotal: sum(already),
    flagged: deduct.filter(d => d.flags.length).length,
    catSkipped, catSkippedAmt,
    combosWithExpenses: new Set(cand.map(e => e.combo)).size,
  };
}

// ---------------------------------------------------------------------------
// Excel output (ExcelJS) — Deduction sheet in the ledger's column order, so the
// rows can be pasted straight into Invoice-Deductions 2026 › Deduction.
// ---------------------------------------------------------------------------
const LEDGER_COLS = ['Date', 'Site ID', 'Job Code', 'Cost/Site', 'Category', 'Sub Category', 'Coordinator', 'Comment',
                     'Contractor', 'VF INV', 'Contractor Invoice', 'Old/New', 'concat'];
const WIDTHS = [12, 10, 11, 11, 11, 15, 18, 22, 14, 10, 12, 9, 18];
const fillOf = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
const THIN = { style: 'thin', color: { argb: 'FF000000' } };
const BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };

function siteOut(s) { return /^\d+$/.test(s) ? Number(s) : s; }

function writeSheet(wb, result, meta) {
  meta = meta || {};
  const ws = wb.addWorksheet('Deduction');
  const cols = LEDGER_COLS.concat(['Check']);
  cols.forEach((h, i) => { ws.getColumn(i + 1).width = i < WIDTHS.length ? WIDTHS[i] : 48; });
  const head = ws.getRow(1);
  cols.forEach((h, i) => {
    const c = head.getCell(i + 1);
    c.value = h;
    c.fill = fillOf(i === cols.length - 1 ? 'FFFFE699' : 'FF9BC2E6');
    c.font = { bold: true, size: 11 };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    c.border = BORDER;
  });
  head.height = 18;
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  let rn = 2;
  for (const d of result.deduct) {
    const e = d.e;
    const vals = [serialToDate(e.day), siteOut(e.site), e.jc, e.cost, e.cat, e.sub, e.coord, e.comment,
                  result.contractor, meta.vfInv ?? '', meta.conInv ?? '', '', e.site + '-' + e.jc, d.flags.join(' · ')];
    const row = ws.getRow(rn);
    vals.forEach((v, i) => {
      const c = row.getCell(i + 1);
      c.value = v === '' ? null : v;
      c.border = BORDER;
      c.font = { size: 11 };
      if (i === 0 && v) c.numFmt = 'dd-mmm-yy';
      if (i === 3) c.numFmt = '#,##0.00';
      if (d.flags.length) c.fill = fillOf('FFFFFF00');
    });
    rn++;
  }
  // Total row
  const tr = ws.getRow(rn);
  for (let i = 1; i <= cols.length; i++) {
    const c = tr.getCell(i);
    c.fill = fillOf('FF00FF00'); c.border = BORDER; c.font = { bold: true, size: 11 };
  }
  tr.getCell(1).value = 'Total';
  tr.getCell(4).value = result.deduct.length
    ? { formula: 'SUM(D2:D' + (rn - 1) + ')', result: Math.round(result.total * 100) / 100 }
    : 0;
  tr.getCell(4).numFmt = '#,##0.00';
  tr.getCell(5).value = result.deduct.length + ' row(s)';

  if (!result.deduct.length) {
    ws.getRow(rn + 2).getCell(1).value = 'No new deductions found for this contractor’s Site ID + Job Code combos.';
  }

  // Already deducted (for checking) — not to be pasted
  if (result.already.length) {
    const ws2 = wb.addWorksheet('Already Deducted');
    const h2 = ['Date', 'Site ID', 'Job Code', 'Cost/Site', 'Category', 'Sub Category', 'Coordinator',
                'Ledger Contractor', 'Ledger VF INV', 'Ledger Contractor Invoice', 'Ledger Row', 'Note'];
    const w2 = [12, 10, 11, 11, 11, 15, 18, 16, 13, 18, 11, 40];
    h2.forEach((h, i) => {
      ws2.getColumn(i + 1).width = w2[i];
      const c = ws2.getRow(1).getCell(i + 1);
      c.value = h; c.fill = fillOf('FFD9D9D9'); c.font = { bold: true, size: 11 }; c.border = BORDER;
      c.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    ws2.views = [{ state: 'frozen', ySplit: 1 }];
    result.already.forEach((a, k) => {
      const e = a.e;
      const vals = [serialToDate(e.day), siteOut(e.site), e.jc, e.cost, e.cat, e.sub, e.coord,
                    a.led.contractor, a.led.vf, a.led.conInv, a.led.excelRow,
                    a.loose ? 'Matched without the date (ledger date ' + (fmtDay(a.led.day) || 'blank') + ')' : ''];
      const row = ws2.getRow(k + 2);
      vals.forEach((v, i) => {
        const c = row.getCell(i + 1);
        c.value = v === '' ? null : v; c.border = BORDER; c.font = { size: 11 };
        if (i === 0 && v) c.numFmt = 'dd-mmm-yy';
        if (i === 3) c.numFmt = '#,##0.00';
      });
    });
  }
  return ws;
}
function fmtDay(s) {
  const d = serialToDate(s);
  if (!d) return '';
  return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: '2-digit', timeZone: 'UTC' });
}

// ---------------------------------------------------------------------------
// UI
// ---------------------------------------------------------------------------
function fire() { document.dispatchEvent(new CustomEvent('con-ded-change')); }
function setErr(msg) {
  const el = $('con-ded-error');
  if (!el) return;
  el.textContent = msg || '';
  el.style.display = msg ? 'block' : 'none';
}
function renderStatus() {
  const el = $('con-ded-status');
  if (!el) return;
  if (!_exp) {
    el.innerHTML = 'Load the expenses file to fill the <em>Deduction</em> sheet automatically. Without it the sheet stays blank for manual entry.';
  } else {
    el.innerHTML = '✅ ' + _exp.rows.length.toLocaleString() + ' expense rows read' +
      (_led ? ' · ' + _led.rows.length.toLocaleString() + ' ledger rows (already deducted) read' :
              ' · <strong>ledger not loaded</strong> — rows deducted before will NOT be removed') + '.';
  }
  renderCats();
}
function renderCats() {
  const wrap = $('con-ded-cats');
  if (!wrap) return;
  if (!_exp) { wrap.innerHTML = ''; wrap.style.display = 'none'; return; }
  wrap.style.display = 'flex';
  wrap.innerHTML = '<span class="con-ded-cats-label">Deduct these categories:</span>' +
    _exp.categories.map(c =>
      '<label class="con-ded-cat"><input type="checkbox" data-cat="' + escA(c.name) + '"' + (_catsOff.has(c.name) ? '' : ' checked') + '> ' +
      esc(c.name) + '</label>').join('');
  wrap.querySelectorAll('input[type=checkbox]').forEach(cb => cb.addEventListener('change', () => {
    if (cb.checked) _catsOff.delete(cb.dataset.cat); else _catsOff.add(cb.dataset.cat);
    try { localStorage.setItem(CAT_KEY, JSON.stringify([..._catsOff])); } catch { /* ignore */ }
    fire();
  }));
}
function esc(s) { return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function escA(s) { return esc(s).replace(/"/g, '&quot;'); }

function bindInput(inputId, btnId, cardId, nameId, progId, reader, assign) {
  const input = $(inputId);
  if (!input) return;
  const btn = $(btnId);
  if (btn) btn.addEventListener('click', () => input.click());
  input.addEventListener('change', async (ev) => {
    const file = ev.target.files && ev.target.files[0];
    if (!file) return;
    setErr('');
    const prog = $(progId);
    if (prog) prog.style.display = 'block';
    $(nameId).textContent = 'Reading ' + file.name + '…';
    await new Promise(r => setTimeout(r, 30));
    try {
      const buf = await file.arrayBuffer();
      const sheet = reader === readExpenses ? 'Total Expense' : 'Deduction';
      const wb = XLSX.read(buf, { type: 'array', cellDates: true, sheets: [sheet] });
      assign(reader(wb, file.name));
      $(nameId).textContent = file.name;
      $(cardId).classList.add('loaded');
    } catch (err) {
      assign(null);
      $(nameId).textContent = 'No file selected';
      $(cardId).classList.remove('loaded');
      setErr(err.message || String(err));
    } finally {
      if (prog) prog.style.display = 'none';
      renderStatus();
      fire();
    }
  });
}

bindInput('con-ded-exp-input', 'btn-pick-con-ded-exp', 'card-con-ded-exp', 'con-ded-exp-filename', 'con-ded-exp-progress',
          readExpenses, v => { _exp = v; });
bindInput('con-ded-led-input', 'btn-pick-con-ded-led', 'card-con-ded-led', 'con-ded-led-filename', 'con-ded-led-progress',
          readLedger, v => { _led = v; });
renderStatus();

window.ConDeductions = {
  ready: () => !!_exp,
  hasLedger: () => !!_led,
  compute, writeSheet, comboKey,
  _test: { siteKey, jcKey, daySerial, readExpenses, readLedger },
};

})();
