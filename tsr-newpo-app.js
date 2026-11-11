// =============================================================================
// LMP Invoicing System — TSR Planner › New PO
// Plans the Request Form of a NEW TSR / PO from the master tracking, using the
// files already loaded in the TSR Planner (tracking + the newest TSR file):
//   1. tasks already in the TSR file are "sent" (ID list for the tracking)
//   2. combos tagged with this TSR (or split from it) stay in it if they fit
//      its free quantity — otherwise they move to the new PO
//   3. every other open combo (whole Site + Job Code) is ranked:
//        1 all Done + FAC + week · 2 all Done + FAC · 3 all Done ·
//        4 partly Done (the open lines go with it)
//      oldest first, whole combos, up to the budget (Request Form value)
// TOC items (EX01/EX02…) ride with their combo. Element IDs use tnp-*.
// =============================================================================
(function () {
'use strict';
const C = window.TSRCore;
if (!C) { console.error('tsr-core.js must load before tsr-newpo-app.js'); return; }
const { str, escHtml, fmtEGP, fmtQty, round4 } = C;
const el = id => document.getElementById(id);
if (!el('tnp-card')) return;

const TOC_CODES = new Set(['EX01', 'EX02']);
const isToc  = t => TOC_CODES.has(t.code) || /^toc$/i.test(t.accStatus) || /^toc/i.test(t.weekRaw);
const isUso  = t => /\buso\d*\b/i.test(t.comments);
const isDone = t => t.status === 'done';
const TIER_LABEL = {
  1: '1 · Done + FAC + week',
  2: '2 · Done + FAC, no week',
  3: '3 · Done, no FAC',
  4: '4 · Partly done (open lines included)',
  0: 'Not started'
};
const DEFAULT_BUDGET = 11990000;
let _res = null;

// ---------------------------------------------------------------------------
// UI
// ---------------------------------------------------------------------------
function state() { return window.TSRPlanner || {}; }
function showError(msg) {
  const d = el('tnp-error');
  d.textContent = msg; d.style.display = 'block';
  d.scrollIntoView({ behavior: 'smooth', block: 'center' });
}
function clearError() { const d = el('tnp-error'); d.textContent = ''; d.style.display = 'none'; }
function setStatus(t) { el('tnp-export-status').textContent = t; }

// "11,990,000", "11.99M", "500K" → number
function parseAmount(v) {
  const s = str(v).replace(/[, ]/g, '').toUpperCase();
  const m = s.match(/^(\d+(?:\.\d+)?)([KM]?)$/);
  if (!m) return NaN;
  return parseFloat(m[1]) * (m[2] === 'M' ? 1e6 : m[2] === 'K' ? 1e3 : 1);
}

function checkReady() {
  const { trk, tsr } = state();
  const b = parseAmount(el('tnp-budget').value);
  el('tnp-btn-run').disabled = !(trk && tsr && b > 0);
  el('tnp-hint').textContent = !(trk && tsr)
    ? 'Load the Task Tracking file and the newest TSR file above (the mails folder is not needed for a new PO).'
    : (b > 0 ? 'Budget ' + fmtEGP(b) + ' — TSR ' + (str(el('tsp-tsrno').value) || tsr.tsrNo || '?') + ' file loaded.' : 'Type the budget, e.g. 11,990,000 or 11.99M.');
}
document.addEventListener('tsp-files', checkReady);
el('tnp-budget').addEventListener('input', checkReady);
el('tsp-tsrno').addEventListener('input', checkReady);
if (!str(el('tnp-budget').value)) el('tnp-budget').value = DEFAULT_BUDGET.toLocaleString('en-US');
checkReady();

el('tnp-btn-run').addEventListener('click', () => {
  clearError();
  el('tnp-results').style.display = 'none';
  try {
    const { trk, tsr } = state();
    _res = runNewPo(trk, tsr, str(el('tsp-tsrno').value) || tsr.tsrNo, parseAmount(el('tnp-budget').value));
    render(_res);
    el('tnp-results').style.display = 'block';
    el('tnp-results').scrollIntoView({ behavior: 'smooth' });
  } catch (err) {
    console.error(err);
    showError(err.message || 'An unexpected error occurred.');
  }
});
el('tnp-btn-export').addEventListener('click', () => exportNewPo().catch(e => showError('Export failed: ' + e.message)));
el('tnp-btn-sent').addEventListener('click', () => exportSent().catch(e => showError('Export failed: ' + e.message)));

// ---------------------------------------------------------------------------
// Core logic
// ---------------------------------------------------------------------------
// Quantity, price and Request Form line of one task
function prepTask(t, tsr) {
  const listed = tsr.siteList.get(C.siteListKey(C.tsrSite(t)));
  const vf = listed ? C.distanceFactor(listed.distance) : null;
  const f = vf ?? t.factor;
  const item = C.tsrItemFor(tsr, t);
  const qty = t.abs * (f ?? 1);
  const price = item ? item.price : (t.price || (t.qty ? t.total / t.qty : t.total));
  t.np = {
    factor: f ?? 1,
    distance: listed ? listed.distance : t.distanceRaw,
    distSrc: vf != null ? 'VF site list' : (t.factor != null ? 'tracking' : 'unknown — 1.0 used'),
    qty, item, price, value: qty * price,
    line: item ? item.name : 'X:' + t.key         // Request Form line key
  };
}

// Earliest date (ms) of a field over tasks; Infinity when none
function minDate(tasks, k) {
  let best = Infinity;
  for (const t of tasks) if (t[k] instanceof Date && !isNaN(t[k])) best = Math.min(best, t[k].getTime());
  return best;
}

function comboOrder(a, b) {
  return (a.tier || 9) - (b.tier || 9) || a.weekKey - b.weekKey || a.facMin - b.facMin ||
         a.dateMin - b.dateMin || a.firstRow - b.firstRow;
}

function runNewPo(trk, tsr, tsrNo, budget) {
  if (!trk || !tsr) throw new Error('Load the tracking file and the TSR file first.');
  if (!(budget > 0)) throw new Error('Type a budget, e.g. 11,990,000.');
  tsrNo = str(tsrNo);
  const tag = tsrNo.toLowerCase();
  const live = trk.tasks.filter(t => t.status !== 'dead');
  const { inTsr, pairs, unmatched } = C.matchExisting(tsr, live, tag);
  const myTagged  = C.tagTest(tag);
  // PO status "ok" = picked for a submission, not received; for this TSR the
  // TSR file is the truth (same rule as the Planner)
  const openOk    = t => /^ok$/i.test(t.poStatus) && (myTagged(t) || !t.sub);
  const submitted = t => inTsr.has(t) || (!!t.poStatus && !openOk(t));
  const otherTsr  = t => !submitted(t) && !!t.sub && !myTagged(t);

  // ── Combos (Job Code + Site), TOC tasks ride with their combo ─────────────
  const combos = new Map();
  for (const t of live) {
    const id = t.jcKey + '|' + t.site;
    if (!combos.has(id)) combos.set(id, { id, jc: t.jc, site: t.site, siteRaw: t.siteRaw, all: [] });
    combos.get(id).all.push(t);
  }
  const list = [], uso = [], other = [];
  for (const c of combos.values()) {
    c.tasks = c.all.filter(t => !submitted(t) && !otherTsr(t));
    c.otherTasks = c.all.filter(otherTsr);
    if (c.otherTasks.length) other.push(c);
    if (!c.tasks.length) continue;
    if (c.all.some(isUso)) { uso.push(c); continue; }
    c.tasks.forEach(t => prepTask(t, tsr));
    const fac = c.tasks.filter(t => !isToc(t));
    if (c.tasks.every(isDone)) {
      c.tier = fac.every(t => t.fac && t.weekSegs.length) ? 1 : fac.every(t => t.fac) ? 2 : 3;
    } else {
      c.tier = c.all.some(isDone) ? 4 : 0;   // a Done line anywhere in the combo (sent or not)
    }
    c.split = c.all.some(submitted) || c.otherTasks.length > 0;
    c.current = c.all.some(t => inTsr.has(t) || myTagged(t));   // belongs to this TSR
    c.tagged = c.tasks.filter(myTagged);
    c.weekKey = Math.min(...c.tasks.map(t => t.weekKey));
    c.weekLabel = c.tasks.map(t => t.weekRaw).filter(Boolean).sort((a, b) => C.weekSortKey(a) - C.weekSortKey(b))[0] || '';
    c.facMin = minDate(c.tasks, 'fac');
    c.dateMin = minDate(c.tasks, 'taskDate');
    c.firstRow = Math.min(...c.tasks.map(t => t.excelRow));
    c.amount = c.tasks.reduce((s, t) => s + t.total, 0);
    c.value = c.tasks.reduce((s, t) => s + t.np.value, 0);
    c.need = new Map();
    for (const t of c.tasks) c.need.set(t.np.line, (c.need.get(t.np.line) || 0) + t.np.qty);
    c.notes = [];
    if (c.split) {
      const parts = [...new Set(c.all.filter(t => !c.tasks.includes(t)).map(t =>
        inTsr.has(t) ? 'TSR ' + tsrNo + ' Submission #' + inTsr.get(t) : (t.subRaw || t.poStatus)))];
      c.notes.push('Split — other lines: ' + parts.join(', '));
    }
    if (c.tasks.some(isToc)) c.notes.push('TOC line(s): ' + [...new Set(c.tasks.filter(isToc).map(t => t.code))].join(', '));
    const unk = c.tasks.filter(t => t.np.distSrc.startsWith('unknown'));
    if (unk.length) c.notes.push('Distance unknown for ' + unk.map(t => t.code).join(', ') + ' (factor 1.0 used)');
    const noItem = c.tasks.filter(t => !t.np.item);
    if (noItem.length) c.notes.push('Not in TSR ' + tsrNo + ' price list: ' + [...new Set(noItem.map(t => t.code || t.item))].join(', ') + ' (tracking price used)');
    list.push(c);
  }
  list.sort(comboOrder);

  // ── 1. Combos of this TSR: stay if they fit its free quantity ─────────────
  const { avail, pendingUse, pendingSubs } = C.tsrAvailable(tsr, pairs);
  const left = new Map(avail);
  const stays = [], moved = [];
  for (const c of list.filter(x => x.current)) {
    const missing = c.tasks.filter(t => !t.np.item);
    const short = [...c.need].filter(([n, q]) => !n.startsWith('X:') && q > (left.get(n) ?? 0) + 0.005);
    if (missing.length || short.length) {
      c.why = [
        ...(missing.length ? ['Item not in TSR ' + tsrNo + ': ' + [...new Set(missing.map(t => t.code || t.item))].join(', ')] : []),
        ...short.map(([n, q]) => C.itemCode(n) + ' needs ' + fmtQty(q) + ', ' + fmtQty(Math.max(left.get(n) ?? 0, 0)) + ' free in TSR ' + tsrNo)
      ].join('; ');
      c.notes.unshift('Moved from TSR ' + tsrNo + ' — ' + c.why);
      moved.push(c);
      continue;
    }
    for (const [n, q] of c.need) left.set(n, left.get(n) - q);
    stays.push(c);
  }

  // ── 2. New PO: ranked combos, whole, up to the budget ─────────────────────
  // Request Form quantities are whole numbers, so each line is rounded up and
  // the budget is checked on Σ ceil(qty) × unit price.
  const lineInfo = new Map();
  for (const c of list) for (const t of c.tasks) {
    if (!lineInfo.has(t.np.line)) lineInfo.set(t.np.line, t.np.item
      ? { name: t.np.item.name, price: t.np.item.price, item: t.np.item }
      : { name: t.item, price: t.np.price, item: null });
  }
  const ceilQ = q => Math.ceil(round4(q) - 1e-9);
  const lineQty = new Map();
  let value = 0;
  const planned = [], over = [], notStarted = [];
  for (const c of list) {
    if (stays.includes(c)) continue;
    if (!c.tier) { notStarted.push(c); continue; }
    let delta = 0;
    for (const [n, q] of c.need) {
      const cur = lineQty.get(n) || 0, price = lineInfo.get(n).price;
      delta += (ceilQ(cur + q) - ceilQ(cur)) * price;
    }
    if (value + delta > budget + 0.005) { c.overBy = value + delta - budget; over.push(c); continue; }
    for (const [n, q] of c.need) lineQty.set(n, (lineQty.get(n) || 0) + q);
    value += delta;
    planned.push(c);
  }

  // Request Form lines: every TSR item in Request Form order, then items not in it
  const lines = tsr.items.map(it => ({ name: it.name, item: it, price: it.price, exact: lineQty.get(it.name) || 0 }));
  for (const [n, q] of lineQty) if (n.startsWith('X:')) lines.push({ name: lineInfo.get(n).name, item: null, price: lineInfo.get(n).price, exact: q });
  for (const l of lines) { l.qty = l.exact ? ceilQ(l.exact) : 0; l.total = l.qty * l.price; }

  const tierSum = {};
  for (const c of planned) {
    const k = c.tier;
    tierSum[k] = tierSum[k] || { combos: 0, tasks: 0, value: 0 };
    tierSum[k].combos++; tierSum[k].tasks += c.tasks.length; tierSum[k].value += c.value;
  }
  return {
    tsrNo, budget, value, planned, stays, moved, over, notStarted, uso, other, lines, tierSum,
    pairs, unmatched, avail, left, pendingUse, pendingSubs,
    rows: planned.flatMap(c => c.tasks.map(t => ({ t, c }))),
    stayRows: stays.flatMap(c => c.tasks.map(t => ({ t, c })))
  };
}

// ---------------------------------------------------------------------------
// Render
// ---------------------------------------------------------------------------
function render(res) {
  const stat = (label, value, amount, cls) =>
    '<div class="poc-stat-box ' + (cls || '') + '"><span class="poc-stat-label">' + label + '</span>' +
    '<span class="poc-stat-value">' + value + '</span><span class="poc-stat-amount">' + amount + '</span></div>';
  const sumV = a => a.reduce((s, c) => s + c.value, 0);
  const movedIn = res.planned.filter(c => res.moved.includes(c));
  const sentUpd = res.pairs.filter(p => needsUpdate(p, res)).length;

  let banners = '<div class="acc-banner acc-banner-ok">New PO: ' + res.planned.length + ' whole combos, ' + res.rows.length +
    ' tasks — Request Form value ' + fmtEGP(res.value) + ' of ' + fmtEGP(res.budget) + ' (' + fmtEGP(res.budget - res.value) + ' left).' +
    ' Quantities rounded up to whole numbers per item.</div>';
  if (res.moved.length) banners += '<div class="acc-banner tso-banner-fix">' + res.moved.length + ' combo' + (res.moved.length > 1 ? 's' : '') +
    ' tagged with / split from TSR ' + escHtml(res.tsrNo) + ' do not fit its free quantity — ' + movedIn.length + ' moved into the new PO' +
    (res.moved.length > movedIn.length ? ', ' + (res.moved.length - movedIn.length) + ' left out (over budget or not started)' : '') + '.</div>';
  if (res.unmatched.length) banners += '<div class="acc-banner tso-banner-fix">' + res.unmatched.length + ' row' + (res.unmatched.length > 1 ? 's' : '') +
    ' of the TSR file could not be matched to a live tracking task (Site + item) — listed at the end of the Sent sheet.</div>';
  if (sentUpd) banners += '<div class="acc-banner tso-banner-fix">' + sentUpd + ' sent task' + (sentUpd > 1 ? 's have' : ' has') +
    ' a different TSR Sub# in the tracking — use the Sent sheet (ID# first) to VLOOKUP the new value.</div>';

  el('tnp-summary').innerHTML = '<div class="poc-stats-row acc-stats-row">' +
    stat('New PO', res.planned.length + ' combos', fmtEGP(res.value), 'acc-stat-pass') +
    stat('Stays in TSR ' + escHtml(res.tsrNo), res.stays.length + ' combos', fmtEGP(sumV(res.stays))) +
    stat('Sent in TSR ' + escHtml(res.tsrNo), res.pairs.length + ' tasks', res.pendingSubs.length ? 'not received: #' + res.pendingSubs.join(', #') : '') +
    stat('Over budget', res.over.length + ' combos', fmtEGP(sumV(res.over)), res.over.length ? 'acc-stat-fail' : '') +
    stat('Not started', res.notStarted.length + ' combos', fmtEGP(sumV(res.notStarted))) +
    '</div>' + banners;

  const section = (title, count, extra, table, open) =>
    '<details class="tso-details"' + (open ? ' open' : '') + '><summary><span class="tso-sum-title">' + title + '</span>' +
    '<span class="tso-sum-count">' + count + '</span>' + (extra ? '<span class="tso-sum-extra">' + extra + '</span>' : '') +
    '</summary><div class="table-wrapper">' + table + '</div></details>';

  // Tiers
  const tiers = [1, 2, 3, 4].map(k => {
    const s = res.tierSum[k] || { combos: 0, tasks: 0, value: 0 };
    return '<tr><td>' + TIER_LABEL[k] + '</td><td class="num">' + s.combos + '</td><td class="num">' + s.tasks + '</td><td class="num">' + fmtEGP(s.value) + '</td></tr>';
  }).join('');
  el('tnp-tier-wrap').innerHTML = section('By priority', res.planned.length + ' combos', fmtEGP(sumV(res.planned)) + ' exact qty',
    '<table class="acc-table"><thead><tr><th>Priority</th><th>Combos</th><th>Tasks</th><th>Value (exact qty)</th></tr></thead><tbody>' + tiers + '</tbody></table>', true);

  // Request Form
  const used = res.lines.filter(l => l.qty);
  el('tnp-rf-wrap').innerHTML = section('Request Form — new PO quantities', used.length + ' items', fmtEGP(res.value),
    '<table class="acc-table"><thead><tr><th>Item</th><th>Exact qty</th><th>Qty (rounded up)</th><th>Unit Price</th><th>Total</th></tr></thead><tbody>' +
    used.map(l => '<tr' + (l.item ? '' : ' class="acc-row-fail"') + '><td class="acc-item-col">' + escHtml(l.name) +
      (l.item ? '' : '<div class="acc-note acc-item-bad">not in TSR ' + escHtml(res.tsrNo) + ' price list — tracking price</div>') +
      '</td><td class="num">' + fmtQty(l.exact) + '</td><td class="num">' + l.qty + '</td><td class="num">' + fmtEGP(l.price) +
      '</td><td class="num">' + fmtEGP(l.total) + '</td></tr>').join('') + '</tbody></table>', true);

  const comboTable = (arr, extraOf) => '<table class="acc-table"><thead><tr><th>Priority</th><th>Job Code</th><th>Site ID</th><th>Acc. Week</th>' +
    '<th>Items</th><th>Value</th><th>Tracking Rows</th><th>Notes</th></tr></thead><tbody>' +
    arr.slice(0, 400).map(c => '<tr><td>' + escHtml(TIER_LABEL[c.tier]) + '</td><td>' + escHtml(c.jc) + '</td><td class="acc-site">' + escHtml(str(c.siteRaw)) +
      '</td><td>' + escHtml(c.weekLabel) + '</td><td class="acc-sub">' + escHtml(c.tasks.map(t => t.code + (isDone(t) ? '' : '*')).join(', ')) +
      '</td><td class="num">' + fmtEGP(c.value) + '</td><td class="acc-sub">' + c.tasks.map(t => t.excelRow).join(', ') + '</td><td class="acc-notes-col">' +
      [...(extraOf ? [extraOf(c)] : []), ...c.notes].filter(Boolean).map(x => '<div class="acc-note">' + escHtml(x) + '</div>').join('') + '</td></tr>').join('') +
    '</tbody></table>' + (arr.length > 400 ? '<p class="acc-sub">First 400 shown — the full list is in the exported file.</p>' : '');
  const star = ' <span class="acc-sub">(* = line not Done yet)</span>';
  el('tnp-planned-wrap').innerHTML = section('New PO combos' + star, res.planned.length + ' combos', fmtEGP(res.value), comboTable(res.planned), false);
  el('tnp-stays-wrap').innerHTML = res.stays.length ? section('Stays in TSR ' + escHtml(res.tsrNo) + ' (fits its free quantity)', res.stays.length + ' combos',
    fmtEGP(sumV(res.stays)), comboTable(res.stays), false) : '';
  el('tnp-over-wrap').innerHTML = res.over.length ? section('Over budget — not in the new PO', res.over.length + ' combos', fmtEGP(sumV(res.over)),
    comboTable(res.over, c => 'Would exceed the budget by ' + fmtEGP(c.overBy)), false) : '';

  el('tnp-btn-export').disabled = false;
  el('tnp-btn-sent').disabled = !res.pairs.length && !res.unmatched.length;
  setStatus('');
}

// ---------------------------------------------------------------------------
// Export
// ---------------------------------------------------------------------------
const THIN = { style: 'thin', color: { argb: 'FFBFBFBF' } };
const BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };
const fill = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const today = () => new Date().toISOString().slice(0, 10);

function headerSheet(wb, name, headers, widths, startRow) {
  const ws = wb.addWorksheet(name);
  headers.forEach((h, i) => { ws.getColumn(i + 1).width = widths[i] || 14; });
  const r = ws.getRow(startRow || 1);
  headers.forEach((h, i) => { r.getCell(i + 1).value = h; });
  r.eachCell(c => { c.fill = fill('FF0070C0'); c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; c.border = BORDER;
                    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; });
  ws.views = [{ state: 'frozen', ySplit: startRow || 1 }];
  return ws;
}
function addRows(ws, rows, numFmts) {
  for (const vals of rows) {
    const r = ws.addRow(vals.map(C.excelDate));
    r.eachCell({ includeEmpty: true }, (c, n) => {
      c.border = BORDER;
      if (c.value instanceof Date) c.numFmt = 'dd-mmm-yy';
      if (numFmts && numFmts[n - 1]) c.numFmt = numFmts[n - 1];
    });
  }
}
const subTag = (res, n) => res.tsrNo + ' sub-' + n;
const normSub = v => str(v).toLowerCase().replace(/\s+/g, ' ').replace(/sub\s*-?\s*/, 'sub-');
function needsUpdate(p, res) { return normSub(p.task.subRaw) !== normSub(subTag(res, p.ex.subNo)); }

// Sent tasks: ID# first so the tracking can VLOOKUP the TSR Sub# by ID
function addSentSheet(wb, res) {
  const ws = headerSheet(wb, 'Sent in TSR ' + res.tsrNo,
    ['ID#', 'Tracking Row', 'Job Code', 'Site ID', 'Line Item', 'TSR Row', 'Submission', 'TSR Sub# (write in tracking)', 'TSR Sub# in tracking now', 'Update?'],
    [24, 12, 12, 12, 48, 9, 12, 22, 22, 9]);
  const sorted = res.pairs.slice().sort((a, b) => (a.ex.subNo || 0) - (b.ex.subNo || 0) || a.ex.excelRow - b.ex.excelRow);
  addRows(ws, sorted.map(p => [str(p.task.id), p.task.excelRow, p.task.jc, p.task.siteRaw, p.task.item, p.ex.excelRow,
    p.ex.subNo ? 'Submission #' + p.ex.subNo : '', p.ex.subNo ? subTag(res, p.ex.subNo) : '', p.task.subRaw, needsUpdate(p, res) ? 'Yes' : '']));
  ws.eachRow((r, n) => { if (n > 1 && r.getCell(10).value === 'Yes') r.getCell(8).fill = fill('FFFFFF00'); });
  if (res.unmatched.length) {
    addRows(ws, res.unmatched.map(ex => ['', '', '', ex.site, ex.itemName, ex.excelRow, ex.subNo ? 'Submission #' + ex.subNo : '',
      ex.subNo ? subTag(res, ex.subNo) : '', 'NOT FOUND in tracking (live Site + item)', '']));
    ws.eachRow((r, n) => { if (n > 1 && /NOT FOUND/.test(str(r.getCell(9).value))) r.eachCell(c => { c.fill = fill('FFFFC7CE'); }); });
  }
  ws.autoFilter = 'A1:J' + ws.rowCount;
  return ws;
}

const TASK_H = ['ID#', 'Tracking Row', 'Job Code', 'Site ID', 'Line Item', 'Request Form Item', 'Status', 'FAC Date', 'Acceptance Week',
  'Priority', 'Distance', 'Distance from', 'Factor', 'Absolute Qty', 'Actual Qty', 'Unit Price', 'Value', 'Tracking New Total',
  'TSR Sub# in tracking now', 'Notes'];
const TASK_W = [24, 12, 12, 12, 44, 44, 10, 11, 18, 26, 14, 14, 8, 10, 10, 12, 14, 14, 18, 60];
const TASK_F = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '0.00', '0.##', '0.00', '#,##0.00', '#,##0.00', '#,##0.00'];
const taskRow = (t, c) => [str(t.id), t.excelRow, c.jc, t.siteRaw, t.item, t.np.item ? t.np.item.name : '(not in TSR price list)',
  t.statusRaw, t.fac || '', t.weekRaw, TIER_LABEL[c.tier], t.np.distance, t.np.distSrc, t.np.factor, t.abs, round4(t.np.qty),
  t.np.price, round4(t.np.value), t.total, t.subRaw, c.notes.join(' | ')];

async function exportNewPo() {
  const res = _res; if (!res) return;
  const tsr = state().tsr;
  const wb = new ExcelJS.Workbook(); wb.creator = 'LMP Invoicing System';

  // 1. Request Form — same item order as the TSR Request Form, so the Qty
  //    column can be pasted straight into a new Request Form
  const rf = headerSheet(wb, 'Request Form',
    ['#', 'Item Description (Catalog only)', 'Material Group', 'EVO Material Number', 'Price Agreement Number', 'PA Item Number',
     'Qty', 'Unit', 'Unit price', 'Total Value', 'Exact Qty (× distance)', 'Tasks'], [5, 62, 10, 14, 14, 9, 8, 12, 12, 15, 12, 7], 3);
  rf.getCell('A1').value = 'New PO Request — budget ' + fmtEGP(res.budget) + ' — value ' + fmtEGP(res.value) +
    ' — built ' + today() + ' from TSR ' + res.tsrNo + ' price list';
  rf.getCell('A1').font = { bold: true, size: 13, color: { argb: 'FF1A3A5C' } };
  rf.getCell('A2').value = 'Qty = Σ (Absolute Qty × distance factor) per item, rounded up. Items in the same order as "Request Form - VF" — paste column G into the Qty column of the new TSR. Red rows = item not in the TSR ' + res.tsrNo + ' price list (tracking price).';
  rf.getCell('A2').font = { italic: true, size: 9 };
  const taskCount = new Map();
  for (const r of res.rows) taskCount.set(r.t.np.item ? r.t.np.item.name : r.t.item, (taskCount.get(r.t.np.item ? r.t.np.item.name : r.t.item) || 0) + 1);
  res.lines.forEach((l, i) => {
    const it = l.item || {};
    const r = rf.addRow([i + 1, l.name, it.matGroup || '', it.evo ? (Number(it.evo) || it.evo) : '', it.pa ? (Number(it.pa) || it.pa) : '',
      it.paItem ? (Number(it.paItem) || it.paItem) : '', l.qty || null, it.unit || '', l.price, null, l.exact ? round4(l.exact) : null,
      taskCount.get(l.name) || null]);
    r.getCell(10).value = { formula: 'IF(G' + r.number + '="",0,G' + r.number + '*I' + r.number + ')' };
    r.eachCell({ includeEmpty: true }, c => { c.border = BORDER; });
    r.getCell(9).numFmt = '#,##0.00'; r.getCell(10).numFmt = '#,##0.00'; r.getCell(11).numFmt = '0.00';
    if (!l.item) r.eachCell({ includeEmpty: true }, c => { c.fill = fill('FFFFC7CE'); });
    else if (l.qty) r.getCell(7).fill = fill('FFE2EFDA');
  });
  const tr = rf.addRow(['', 'Grand Total', '', '', '', '', null, '', '', null]);
  const last = tr.number - 1;
  tr.getCell(7).value = { formula: 'SUM(G4:G' + last + ')' };
  tr.getCell(10).value = { formula: 'SUM(J4:J' + last + ')' };
  tr.getCell(10).numFmt = '#,##0.00';
  for (let j = 1; j <= 12; j++) { const c = tr.getCell(j); c.font = { bold: true }; c.fill = fill('FFC0E6F5'); c.border = BORDER; }
  const br = rf.addRow(['', 'Budget', '', '', '', '', '', '', '', res.budget]);
  br.getCell(10).numFmt = '#,##0.00';
  const lr = rf.addRow(['', 'Left', '', '', '', '', '', '', '', null]);
  lr.getCell(10).value = { formula: 'J' + br.number + '-J' + tr.number }; lr.getCell(10).numFmt = '#,##0.00';
  [br, lr].forEach(r => { r.getCell(2).font = { bold: true }; r.getCell(10).font = { bold: true }; });

  // 2. New PO tasks (ID# first — write "New PO" in the TSR Sub# column)
  const nt = headerSheet(wb, 'New PO Tasks', [...TASK_H, 'TSR Sub# (write in tracking)'], [...TASK_W, 20]);
  addRows(nt, res.rows.map(({ t, c }) => [...taskRow(t, c), 'New PO']), TASK_F);
  nt.eachRow((r, n) => { if (n > 1 && !/^done$/i.test(str(r.getCell(7).value))) r.getCell(7).fill = fill('FFFFF2CC'); });
  nt.autoFilter = 'A1:U' + nt.rowCount;

  // 3. Stays in the current TSR + its quantity check
  if (res.stayRows.length) {
    const st = headerSheet(wb, 'Stays in TSR ' + res.tsrNo, TASK_H, TASK_W);
    addRows(st, res.stayRows.map(({ t, c }) => taskRow(t, c)), TASK_F);
    st.autoFilter = 'A1:T' + st.rowCount;
  }
  const qs = headerSheet(wb, 'Qty Check TSR ' + res.tsrNo, ['Item', 'TSR Qty', 'Remaining (Request Form)', 'Sent, not yet received' +
    (res.pendingSubs.length ? ' (#' + res.pendingSubs.join(', #') + ')' : ''), 'Stays in TSR (not sent yet)', 'Balance'], [60, 10, 14, 16, 16, 12]);
  const stayQty = new Map();
  for (const c of res.stays) for (const [n, q] of c.need) stayQty.set(n, (stayQty.get(n) || 0) + q);
  tsr.items.filter(it => it.qty || it.remaining || stayQty.get(it.name) || res.pendingUse.get(it.name)).forEach(it => {
    const r = qs.addRow([it.name, it.qty, round4(it.remaining), round4(res.pendingUse.get(it.name) || 0), round4(stayQty.get(it.name) || 0), null]);
    r.getCell(6).value = { formula: 'C' + r.number + '-D' + r.number + '-E' + r.number };
    r.eachCell({ includeEmpty: true }, c => { c.border = BORDER; });
    [3, 4, 5, 6].forEach(k => { r.getCell(k).numFmt = '0.00'; });
  });

  // 4. Sent in the current TSR (ID list for the tracking)
  addSentSheet(wb, res);

  // 5. Left out
  const cH = ['Priority', 'Job Code', 'Site ID', 'Acceptance Week', 'Items (* = not Done)', 'Value (EGP)', 'Tracking Rows', 'ID#s', 'Notes'];
  const cW = [26, 12, 12, 18, 30, 14, 30, 40, 90];
  const comboRows = (arr, extra) => arr.map(c => [TIER_LABEL[c.tier], c.jc, c.siteRaw, c.weekLabel,
    c.tasks.map(t => t.code + (isDone(t) ? '' : '*')).join(', '), round4(c.value), c.tasks.map(t => t.excelRow).join(', '),
    c.tasks.map(t => str(t.id)).join(', '), [...(extra ? [extra(c)] : []), ...c.notes].filter(Boolean).join(' | ')]);
  const nf = [0, 0, 0, 0, 0, '#,##0'];
  if (res.over.length) addRows(headerSheet(wb, 'Over Budget', cH, cW), comboRows(res.over, c => 'Over the budget by ' + fmtEGP(c.overBy)), nf);
  if (res.notStarted.length) addRows(headerSheet(wb, 'Not Started', cH, cW), comboRows(res.notStarted, () => 'No line of this combo is Done'), nf);
  if (res.other.length) {
    const ws = headerSheet(wb, 'In Another TSR', ['Job Code', 'Site ID', 'Line Item', 'Status', 'TSR Sub# in tracking', 'Tracking Row', 'ID#'], [12, 12, 48, 10, 20, 12, 24]);
    addRows(ws, res.other.flatMap(c => c.otherTasks.map(t => [c.jc, c.siteRaw, t.item, t.statusRaw, t.subRaw, t.excelRow, str(t.id)])));
  }
  if (res.uso.length) addRows(headerSheet(wb, 'USO (excluded)', cH, cW), res.uso.map(c => ['', c.jc, c.siteRaw, '', c.tasks.map(t => t.code).join(', '),
    c.tasks.reduce((s, t) => s + t.total, 0), c.tasks.map(t => t.excelRow).join(', '), c.tasks.map(t => str(t.id)).join(', '), 'USO in Comments']), nf);

  const name = 'New_PO_Request_' + today() + '.xlsx';
  C.downloadBlob(new Blob([await wb.xlsx.writeBuffer()], { type: XLSX_TYPE }), name);
  setStatus('✅ Downloaded ' + name + ' — ' + res.rows.length + ' tasks, ' + fmtEGP(res.value) + '.');
}

async function exportSent() {
  const res = _res; if (!res) return;
  const wb = new ExcelJS.Workbook(); wb.creator = 'LMP Invoicing System';
  addSentSheet(wb, res);
  const name = 'TSR_' + res.tsrNo + '_Sent_Tasks_' + today() + '.xlsx';
  C.downloadBlob(new Blob([await wb.xlsx.writeBuffer()], { type: XLSX_TYPE }), name);
  setStatus('✅ Downloaded ' + name + ' — ' + res.pairs.length + ' sent tasks (ID# in column A for VLOOKUP).');
}

window.__tnpDebug = { get result() { return _res; }, runNewPo, parseAmount };
})();
