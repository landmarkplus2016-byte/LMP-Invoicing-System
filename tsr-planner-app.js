// =============================================================================
// LMP Invoicing System — TSR Planner
// Plans the next submission of a TSR from the master tracking:
//   1. readiness check of every Site + Job Code combo (FAC track; EX01/EX02 and
//      TOC items are a separate track, planned later)
//   2. oldest acceptance first, whole combos only, within the TSR Remaining
//   3. paste-ready blocks for the protected PO Break Down sheet, the receiving
//      template for VF invoicing, the numbered mail folders and the allocation list
// Shared rules come from tsr-core.js (window.TSRCore). Element IDs use tsp-*.
// =============================================================================
(function () {
'use strict';
const C = window.TSRCore;
if (!C) { console.error('tsr-core.js must load before tsr-planner-app.js'); return; }
const { str, escHtml, fmtEGP, fmtQty, round4, isBlank, siteKey } = C;

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------
let _trk = null, _tsr = null, _mails = [], _ignored = [];
let _mailCache = new Map();
let _result = null;
const _streamOverride = new Map();   // combo id → stream picked in the app

const el = id => document.getElementById(id);
if (!el('panel-tsrplan')) return;

// General Stream rule: a combo takes the stream of its main item (first match)
const STREAM_RULES = [
  [k => k === 'RF-MACRO-5PLUS' || k === 'RF-MACRO-UPTO4', 'New Physical Sites'],
  [k => k === 'TX04', 'Upgrade-MW Link E-Band 2+0'],
  [k => k === 'TX03', 'Upgrade-MW Link E-Band 1+0'],
  [k => k === 'TX05', 'Upgrade-MW Link XPIC 2+0'],
  [k => k === 'TX09', 'Upgrade-IDU Upgrade "Cards /Nodes Addition"'],
  [k => k === 'TX16', 'Additional Cards with Cables'],
  [k => k === 'RF09', 'Add Sector'],
  [k => k === 'PW02', 'Add Batteries & Rectifiers'],
  [k => k === 'TX19', 'Upgrade-MW Link XPIC 2+0'],
  [k => /^TX(10|11|12|13|15)$/.test(k), 'Upgrade-TX Cutover']
];
const TOC_CODES = new Set(['EX01', 'EX02']);

// ---------------------------------------------------------------------------
// UI helpers
// ---------------------------------------------------------------------------
function showError(msg) {
  const d = el('tsp-error');
  d.textContent = msg; d.style.display = 'block';
  d.scrollIntoView({ behavior: 'smooth', block: 'center' });
}
function clearError() { const d = el('tsp-error'); d.textContent = ''; d.style.display = 'none'; }
function showProgress(id, show) { const e = el(id); if (e) e.style.display = show ? 'block' : 'none'; }
function setLoadingText(t) { const p = el('tsp-loading').querySelector('p'); if (p) p.textContent = t; }
function setStatus(t) { el('tsp-export-status').textContent = t; }

function checkReady() {
  el('tsp-btn-run').disabled = !(_trk && _tsr && _mails.length && parseInt(el('tsp-sub').value, 10) > 0);
}

function wireFilePicker(btnId, inputId, cardId, nameId, progressId, onLoad) {
  el(btnId).addEventListener('click', () => el(inputId).click());
  el(inputId).addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      clearError();
      showProgress(progressId, true);
      await new Promise(r => setTimeout(r, 20));
      const data = await file.arrayBuffer();
      onLoad(XLSX.read(data, { type: 'array', cellDates: true }), file);
      el(nameId).textContent = file.name;
      el(cardId).classList.add('loaded');
    } catch (err) {
      showError('Failed to open ' + file.name + ': ' + err.message);
    } finally {
      showProgress(progressId, false);
      checkReady();
      document.dispatchEvent(new CustomEvent('tsp-files'));   // tells the New PO section
    }
  });
}

wireFilePicker('tsp-btn-track', 'tsp-input-track', 'tsp-card-track', 'tsp-track-filename', 'tsp-track-progress',
  wb => { _trk = null; _trk = C.readTracking(wb); fillDefaults(); });
wireFilePicker('tsp-btn-tsr', 'tsp-input-tsr', 'tsp-card-tsr', 'tsp-tsr-filename', 'tsp-tsr-progress',
  (wb, file) => { _tsr = null; _tsr = C.readTsrFile(wb, file.name); fillDefaults(true); });

el('tsp-btn-folder').addEventListener('click', () => el('tsp-input-folder').click());
el('tsp-input-folder').addEventListener('change', e => loadFolderFiles(Array.from(e.target.files)));
el('tsp-input-folder').addEventListener('folderdrop', e => loadFolderFiles(e.detail));

function loadFolderFiles(files) {
  if (!files.length) return;   // errors from the file cards stay visible
  ({ mails: _mails, ignored: _ignored } = C.buildMailList(files));
  _mailCache = new Map();
  const root = (files[0].webkitRelativePath || '').split('/')[0] || 'Folder';
  const counts = ['U', 'A', 'D'].map(a => C.AREA_LABEL[a] + ' ' + _mails.filter(m => m.area === a).length);
  el('tsp-folder-filename').textContent = root + '  (' + counts.join(' · ') + ' mails)';
  el('tsp-card-folder').classList.toggle('loaded', _mails.length > 0);
  if (!_mails.length) showError('No acceptance mails found. Expected Upper-Cairo / Alex / Delta folders with mails named like "… Status W34-2026".');
  checkReady();
}

// Settings pre-filled from the files (all editable)
function fillDefaults(fromTsr) {
  if (_tsr && fromTsr) {
    el('tsp-tsrno').value = _tsr.tsrNo || '';
    el('tsp-sub').value = _tsr.nextSub;
    el('tsp-row').value = _tsr.nextRow;
    const done = [..._tsr.submissions].sort((a, b) => a - b);
    el('tsp-tsr-hint').textContent = 'TSR ' + (_tsr.tsrNo || '?') + ': ' +
      (done.length ? 'Submission #' + done.join(', #') + ' already in the file (' + _tsr.existing.length + ' rows, last row ' + _tsr.lastRow + ').'
                   : 'no submission in the file yet.') +
      (_tsr.warnings.length ? '  ⚠ ' + _tsr.warnings.join(' ') : '');
  }
  if (_trk && _tsr && !str(el('tsp-po').value)) {
    const tag = str(el('tsp-tsrno').value).toLowerCase();
    const po = _trk.tasks.find(t => tag && t.sub.startsWith(tag) && /^\d{6,}$/.test(t.poNumber));
    if (po) el('tsp-po').value = po.poNumber;
  }
  checkReady();
}

['tsp-sub', 'tsp-tsrno'].forEach(id => el(id).addEventListener('input', checkReady));

el('tsp-btn-run').addEventListener('click', async () => {
  clearError();
  el('tsp-results').style.display = 'none';
  el('tsp-loading').style.display = 'flex';
  await new Promise(r => setTimeout(r, 30));
  try {
    _streamOverride.clear();
    _result = await runPlan();
    renderResults(_result);
    el('tsp-results').style.display = 'block';
    el('tsp-results').scrollIntoView({ behavior: 'smooth' });
  } catch (err) {
    console.error(err);
    showError(err.message || 'An unexpected error occurred.');
  } finally {
    el('tsp-loading').style.display = 'none';
  }
});

el('tsp-btn-paste').addEventListener('click', () => exportPaste().catch(e => showError('Export failed: ' + e.message)));
el('tsp-btn-recv').addEventListener('click', () => exportReceiving().catch(e => showError('Export failed: ' + e.message)));
el('tsp-btn-folders').addEventListener('click', () => exportFolders().catch(e => showError('Folder download failed: ' + e.message)));

// ---------------------------------------------------------------------------
// Core logic
// ---------------------------------------------------------------------------
async function getMail(m) {
  if (!_mailCache.has(m.path)) {
    setLoadingText('Reading mail: ' + m.name + '…');
    _mailCache.set(m.path, await C.parseMailFile(m.file));
  }
  return _mailCache.get(m.path);
}

const isToc = t => TOC_CODES.has(t.code) || /^toc$/i.test(t.accStatus) || /^toc/i.test(t.weekRaw);
const isUso = t => /\buso\d*\b/i.test(t.comments);

function streamFor(combo, siteHasMacro) {
  const keys = combo.tasks.map(t => t.key);
  let stream = null, rule = null;
  for (const [test, s] of STREAM_RULES) {
    if (keys.some(test)) { stream = s; rule = keys.find(test); break; }
  }
  const valid = validStreams();
  let check = '';
  if (!stream) { stream = 'Upgrade-TX Cutover'; check = 'No main item found — please pick the stream'; }
  else if (siteHasMacro && /^TX0[345]$/.test(rule)) {
    stream = 'New Physical Sites';
    check = 'Site also has a new-macro (RF) task — New Physical Sites suggested, or New rollout-MW Link XPIC 2+0';
  }
  if (valid.length && !valid.some(v => v.toLowerCase() === stream.toLowerCase())) {
    check = (check ? check + '; ' : '') + '"' + stream + '" is not in the TSR stream list';
  }
  const exact = valid.find(v => v.toLowerCase() === stream.toLowerCase());
  return { stream: exact || stream, check };
}

function validStreams() {
  if (!_tsr) return [];
  const set = new Set(_tsr.streams);
  for (const r of _tsr.receiving) if (r.stream) set.add(r.stream);
  return [...set].sort((a, b) => a.localeCompare(b));
}

async function runPlan() {
  const tsr   = _tsr;
  const tsrNo = str(el('tsp-tsrno').value);
  const subNo = parseInt(el('tsp-sub').value, 10);
  const tag   = tsrNo.toLowerCase();
  const live  = _trk.tasks.filter(t => t.status !== 'dead');

  // ── 1. Tasks already in this TSR file (earlier submissions) ──────────────
  const { inTsr, pairs } = C.matchExisting(tsr, live, tag);   // task → submission #
  const certOf = v => str(v).replace(/\.0+$/, '');
  // For THIS TSR the loaded TSR file is the truth. PO status "ok" means "picked
  // for a submission, not received yet": if the task is tagged with this TSR #
  // ("1326", "1326 sub-2", "check 1326") or not tagged at all and it is not in
  // the file, it has not been submitted here yet, so it stays available.
  // "Received", "REJ", other PO statuses and other TSRs' tags still exclude it.
  const myTagged  = C.tagTest(tag);
  const openOk    = t => /^ok$/i.test(t.poStatus) && (myTagged(t) || !t.sub);
  const submitted = t => inTsr.has(t) || (!!t.poStatus && !openOk(t));
  const otherTsr  = t => !submitted(t) && t.sub && !myTagged(t);

  // ── 2. Combos (Job Code + Site), FAC track only ──────────────────────────
  const tocTasks = [];
  const combos = new Map();
  for (const t of live) {
    if (isToc(t)) { if (!submitted(t) && C.tsrItemFor(tsr, t)) tocTasks.push(t); continue; }
    const id = t.jcKey + '|' + t.site;
    if (!combos.has(id)) combos.set(id, { id, jc: t.jc, site: t.site, siteRaw: t.siteRaw, all: [] });
    combos.get(id).all.push(t);
  }
  const macroSites = new Set(live.filter(t => /^RF-MACRO/.test(t.key)).map(t => t.site));

  const list = [];
  for (const c of combos.values()) {
    c.tasks = c.all.filter(t => !submitted(t));
    if (!c.tasks.length) continue;                               // fully submitted already
    if (!c.tasks.some(t => C.tsrItemFor(tsr, t))) continue;      // nothing for this TSR
    if (c.all.some(isUso)) continue;                             // USO scope
    c.done = c.all.filter(submitted);
    c.weekKey = Math.min(...c.tasks.map(t => t.weekKey));
    c.weekLabel = (c.tasks.map(t => t.weekRaw).filter(Boolean).sort((a, b) => C.weekSortKey(a) - C.weekSortKey(b))[0]) || '';
    c.firstRow = Math.min(...c.tasks.map(t => t.excelRow));
    c.amount = c.tasks.reduce((s, t) => s + t.total, 0);
    c.other = c.tasks.filter(otherTsr);
    list.push(c);
  }
  list.sort((a, b) => a.weekKey - b.weekKey || a.weekLabel.localeCompare(b.weekLabel) || a.firstRow - b.firstRow);

  // ── 3. Readiness (mail rows are claimed oldest combo first) ──────────────
  const used = new Set();
  for (const c of list) {
    c.reasons = [];
    if (c.other.length) {
      c.reasons.push('Planned in another TSR (' + [...new Set(c.other.map(t => t.subRaw))].join(', ') + ')');
    }
    if (c.done.length) {
      c.reasons.push('Part already submitted (' + c.done.length + ' task' + (c.done.length > 1 ? 's' : '') + ': ' +
        [...new Set(c.done.map(t => inTsr.has(t) ? 'Submission #' + inTsr.get(t) : (t.subRaw || t.poStatus)))].join(', ') + ')');
    }
    for (const t of c.tasks) {
      t.why = []; t.mailHit = null; t.fix = null;
      t.tsrItem = C.tsrItemFor(tsr, t);
      // The TSR counts Qty × the distance of VF's own site list; the tracking
      // distance is only used for sites VF has not listed yet (#N/A in the TSR).
      const listed = tsr.siteList.get(C.siteListKey(tsrSite(t)));
      t.tsrFactor = listed ? C.distanceFactor(listed.distance) : null;
      t.effFactor = t.tsrFactor ?? t.factor;
      t.effQty = t.abs * (t.effFactor ?? 1);
      if (t.status !== 'done') t.why.push('Not done (' + (t.statusRaw || 'blank') + ')');
      else {
        if (!t.fac) t.why.push('No FAC date');
        if (!t.weekSegs.length) t.why.push(t.weekRaw ? 'Acceptance week unreadable: ' + t.weekRaw : 'No acceptance week');
      }
      if (!t.tsrItem) t.why.push('Item not in TSR ' + tsrNo);
      if (t.effFactor == null) t.why.push('Distance missing (not in VF site list, blank in tracking)');
      if (t.status === 'done' && t.weekSegs.length) {
        const cands = [];
        for (const seg of t.weekSegs) for (const m of C.mailsFor(_mails, seg)) if (!cands.includes(m)) cands.push(m);
        let hit = null;
        for (const m of cands) {
          const parsed = await getMail(m);
          if (parsed.error) continue;
          const free = parsed.rows.filter(r => r.site === t.site && r.key === t.key && !used.has(m.path + '|' + r.pos));
          if (!free.length) continue;
          free.sort((a, b) =>
            (certOf(b.request) === certOf(t.cert)) - (certOf(a.request) === certOf(t.cert)) ||
            (b.facing === t.facing) - (a.facing === t.facing));
          hit = { mail: m, row: free[0] }; break;
        }
        if (hit) {
          used.add(hit.mail.path + '|' + hit.row.pos);
          t.mailHit = hit;
          t.fix = C.mailFixes(t, hit.row);
        } else {
          t.why.push(cands.length ? 'Not in acceptance mail for ' + t.weekRaw : 'No acceptance mail for ' + t.weekRaw);
        }
      }
      for (const w of t.why) c.reasons.push(t.code + ': ' + w);
    }
    c.reasons = [...new Set(c.reasons)];
  }

  // ── 4. Oldest first, whole combos, within the TSR Remaining ──────────────
  // Remaining on the Request Form only drops when VF receives a submission, so
  // rows of submissions already in the file but not yet received are deducted.
  const { avail, pendingUse, pendingSubs } = C.tsrAvailable(tsr, pairs);
  const ready = list.filter(c => !c.reasons.length);
  const planned = [], waiting = [];
  for (const c of ready) {
    const need = new Map();
    for (const t of c.tasks) need.set(t.tsrItem.name, (need.get(t.tsrItem.name) || 0) + t.effQty);
    const short = [...need].filter(([n, q]) => q > (avail.get(n) ?? 0) + 0.005);
    if (short.length) {
      c.short = short.map(([n, q]) => C.itemCode(n) + ' needs ' + fmtQty(q) + ', ' + fmtQty(Math.max(avail.get(n) ?? 0, 0)) + ' left').join('; ');
      waiting.push(c); continue;
    }
    for (const [n, q] of need) avail.set(n, avail.get(n) - q);
    c.need = need;
    const st = streamFor(c, macroSites.has(c.site));
    c.stream = st.stream; c.streamCheck = st.check;
    planned.push(c);
  }

  // ── 5. Rows in mail order, folders numbered 1, 2, 3… ─────────────────────
  const rows = [];
  for (const c of planned) for (const t of c.tasks) rows.push({ t, c });
  rows.sort((a, b) => a.t.mailHit.mail.rank - b.t.mailHit.mail.rank || a.t.mailHit.row.pos - b.t.mailHit.row.pos);
  const folders = [...new Set(rows.map(r => r.t.mailHit.mail))];
  const folderOf = new Map(folders.map((m, i) => [m, i + 1]));
  for (const r of rows) r.folder = folderOf.get(r.t.mailHit.mail);

  const blocked = list.filter(c => c.reasons.length && !c.other.length);
  const pending = list.filter(c => c.other.length);
  return {
    tsrNo, subNo, pasteRow: parseInt(el('tsp-row').value, 10) || tsr.nextRow,
    subDate: el('tsp-date').value ? new Date(el('tsp-date').value + 'T00:00:00') : null,
    poNo: str(el('tsp-po').value),
    list, ready, planned, waiting, blocked, pending, tocTasks, rows, folders, avail,
    pendingUse, pendingSubs,
    total: rows.reduce((s, r) => s + r.t.total, 0)
  };
}

const tsrSite = C.tsrSite;

// Paste values for one planned row (TSR column letter → value)
function pasteRow(r, res) {
  const t = r.t, tsr = _tsr;
  const fix = t.fix || { cert: t.cert, facing: t.facingRaw };
  const site = tsrSite(t);
  const vendor = tsr.vendors.find(v => v.toLowerCase() === t.vendor.toLowerCase()) || t.vendor;
  const certTxt = str(fix.cert).replace(/\.0+$/, '');
  const cert = /^\d+$/.test(certTxt) ? Number(certTxt) : certTxt;
  const stream = _streamOverride.get(r.c.id) || r.c.stream;
  const listed = tsr.siteList.get(C.siteListKey(site));
  const cat = tsr.catalogue.get(t.tsrItem.name);
  const kFactor = listed ? C.distanceFactor(listed.distance) : null;
  return {
    C: t.owner, D: vendor, H: site, I: str(t.option), J: isBlank(fix.facing) ? '' : fix.facing,
    M: t.taskDate instanceof Date ? t.taskDate : (t.taskDate || ''), N: t.tsrItem.name, P: t.abs,
    S: stream, T: str(t.prq), U: cert, V: 'FAC', W: 'Submission #' + res.subNo,
    X: res.subDate || '', Y: r.folder,
    // preview of what the TSR formulas will show
    F: listed ? listed.region : '#N/A', G: listed ? listed.governorate : '#N/A',
    K: listed ? listed.distance : '#N/A', L: listed ? listed.physical : '#N/A',
    O: cat ? cat.unit : '#N/A', Q: cat ? cat.price : t.tsrItem.price,
    R: listed && cat ? round4(t.abs * cat.price * (kFactor ?? 1)) : '',
    inList: !!listed, vfValue: round4(t.abs * (cat ? cat.price : t.tsrItem.price) * (t.effFactor ?? 1))
  };
}

// ---------------------------------------------------------------------------
// Render
// ---------------------------------------------------------------------------
function renderResults(res) {
  const tsr = _tsr;
  const stat = (label, value, amount, cls) =>
    '<div class="poc-stat-box ' + (cls || '') + '"><span class="poc-stat-label">' + label + '</span>' +
    '<span class="poc-stat-value">' + value + '</span><span class="poc-stat-amount">' + amount + '</span></div>';
  const sum = a => a.reduce((s, c) => s + c.amount, 0);
  const pr = res.rows.map(r => pasteRow(r, res));
  const notListed = [...new Set(res.rows.filter((r, i) => !pr[i].inList).map(r => str(r.t.siteRaw)))];
  const fixes = res.rows.filter(r => r.t.fix && r.t.fix.fixes.length).length;
  const checks = res.planned.filter(c => c.streamCheck).length;
  const vfTotal = pr.reduce((s, p) => s + p.vfValue, 0);

  el('tsp-col-info').textContent = [
    'TSR ' + res.tsrNo, 'Submission #' + res.subNo, 'Paste at row ' + res.pasteRow,
    'Tracking header row ' + _trk.headerRow, 'Mails ' + _mails.length + (_ignored.length ? ' (' + _ignored.length + ' ignored)' : ''),
    'VF site list ' + tsr.siteList.size.toLocaleString() + ' sites'
  ].join('  ·  ');

  let banners = '';
  if (!res.rows.length) banners += '<div class="acc-banner acc-banner-fail">No combo can go into this submission — see Waiting and Not ready below.</div>';
  else banners += '<div class="acc-banner acc-banner-ok">Planned ' + res.planned.length + ' whole combos, ' + res.rows.length +
    ' rows in ' + res.folders.length + ' mail folders — tracking value ' + fmtEGP(res.total) +
    ', TSR value ' + fmtEGP(vfTotal) + ' (TSR prices × distance). No item exceeds its TSR Remaining.</div>';
  if (fixes) banners += '<div class="acc-banner tso-banner-fix">' + fixes + ' row' + (fixes > 1 ? 's' : '') +
    ' corrected from the acceptance mail (Certificate # / Facing) — the tracking should be updated the same way.</div>';
  if (notListed.length) banners += '<div class="acc-banner tso-banner-fix">' + notListed.length + ' site' + (notListed.length > 1 ? 's are' : ' is') +
    ' not in VF\'s site list in the TSR, so Distance will show #N/A until VF adds them: ' + escHtml(notListed.join(', ')) +
    '. The tracking distance is used for the quantity check.</div>';
  const distDiff = res.rows.filter((r, i) => pr[i].inList && r.t.factor != null && r.t.tsrFactor !== r.t.factor).length;
  if (distDiff) banners += '<div class="acc-banner tso-banner-fix">' + distDiff + ' row' + (distDiff > 1 ? 's have' : ' has') +
    ' a different distance in VF\'s site list than in the tracking. The TSR calculates with VF\'s distance, so quantities and the receiving template use it too (shown under "Distance in TSR").</div>';
  if (checks) banners += '<div class="acc-banner tso-banner-fix">' + checks + ' combo' + (checks > 1 ? 's need' : ' needs') +
    ' a General Stream check — marked ⚠ in the planned list; pick from the dropdown before exporting.</div>';

  el('tsp-summary').innerHTML =
    '<div class="poc-stats-row acc-stats-row">' +
      stat('Planned', res.planned.length + ' combos', fmtEGP(res.total), 'acc-stat-pass') +
      stat('Ready, waiting for qty', res.waiting.length, fmtEGP(sum(res.waiting))) +
      stat('Not ready', res.blocked.length, fmtEGP(sum(res.blocked)), res.blocked.length ? 'acc-stat-fail' : '') +
      stat('In another TSR', res.pending.length, fmtEGP(sum(res.pending))) +
      stat('TOC track (later)', res.tocTasks.length + ' tasks', fmtEGP(res.tocTasks.reduce((s, t) => s + t.total, 0))) +
    '</div>' + banners;

  const section = (title, count, extra, table, open) =>
    '<details class="tso-details"' + (open ? ' open' : '') + '><summary><span class="tso-sum-title">' + title + '</span>' +
    '<span class="tso-sum-count">' + count + '</span>' + (extra ? '<span class="tso-sum-extra">' + extra + '</span>' : '') +
    '</summary><div class="table-wrapper">' + table + '</div></details>';

  // Planned rows (stream select per combo, on its first row)
  const streams = validStreams();
  const seen = new Set();
  const planned = '<table class="acc-table"><thead><tr><th>#</th><th>Folder</th><th>Site ID</th><th>Job Code</th>' +
    '<th>Line Item (TSR name)</th><th>Qty</th><th>Acc. Week</th><th>General Stream</th><th>Distance in TSR</th>' +
    '<th>Tracking Row</th><th>Amount</th><th>Corrected from mail</th></tr></thead><tbody>' +
    res.rows.map((r, i) => {
      const p = pr[i];
      let streamCell = '<span class="acc-sub">same combo</span>';
      if (!seen.has(r.c.id)) {
        seen.add(r.c.id);
        const cur = _streamOverride.get(r.c.id) || r.c.stream;
        streamCell = '<select class="tsp-stream" data-combo="' + escHtml(r.c.id) + '">' +
          streams.map(s => '<option' + (s === cur ? ' selected' : '') + '>' + escHtml(s) + '</option>').join('') +
          (streams.includes(cur) ? '' : '<option selected>' + escHtml(cur) + '</option>') + '</select>' +
          (r.c.streamCheck && !_streamOverride.has(r.c.id) ? '<div class="acc-note acc-item-bad">⚠ ' + escHtml(r.c.streamCheck) + '</div>' : '');
      }
      return '<tr class="acc-row-pass"><td>' + (i + 1) + '</td><td class="tso-folder-col">' + r.folder + '</td>' +
        '<td class="acc-site">' + escHtml(str(r.t.siteRaw)) + '</td><td>' + escHtml(r.c.jc) + '</td>' +
        '<td class="acc-item-col">' + escHtml(p.N) + '</td><td class="num">' + fmtQty(p.P) + '</td>' +
        '<td>' + escHtml(r.t.weekRaw) + '</td><td>' + streamCell + '</td>' +
        '<td>' + (p.inList ? escHtml(p.K) + (r.t.factor != null && r.t.tsrFactor !== r.t.factor
            ? '<div class="acc-note">tracking: ' + escHtml(r.t.distanceRaw) + '</div>' : '')
          : '<span class="acc-item-bad">#N/A (tracking: ' + escHtml(r.t.distanceRaw) + ')</span>') + '</td>' +
        '<td>' + r.t.excelRow + '</td><td class="num">' + fmtEGP(r.t.total) + '</td>' +
        '<td class="acc-notes-col">' + (r.t.fix ? r.t.fix.fixes.map(f => '<div class="acc-note">' + escHtml(f.text) + '</div>').join('') : '') + '</td></tr>';
    }).join('') + '</tbody></table>';
  el('tsp-planned-wrap').innerHTML = res.rows.length
    ? section('Planned submission — in mail order', res.rows.length + ' rows', fmtEGP(res.total), planned, true) : '';
  el('tsp-planned-wrap').querySelectorAll('.tsp-stream').forEach(s => s.addEventListener('change', () => {
    _streamOverride.set(s.dataset.combo, s.value);
    const note = s.parentElement.querySelector('.acc-note'); if (note) note.remove();
  }));

  // Quantity check
  const usedQty = new Map();
  for (const c of res.planned) for (const [n, q] of c.need) usedQty.set(n, (usedQty.get(n) || 0) + q);
  const qrows = _tsr.items.filter(it => it.qty || it.remaining || usedQty.get(it.name) || res.pendingUse.get(it.name));
  const pendLabel = res.pendingSubs.length ? 'Sent, not received (#' + res.pendingSubs.join(', #') + ')' : 'Sent, not received';
  el('tsp-qty-wrap').innerHTML = section('Quantity check', qrows.length + ' items', 'Balance after this submission', '<table class="acc-table"><thead><tr>' +
    '<th>Item</th><th>TSR Qty</th><th>Remaining (Request Form)</th><th>' + pendLabel + '</th><th>This submission</th><th>Balance</th></tr></thead><tbody>' +
    qrows.map(it => {
      const u = usedQty.get(it.name) || 0, pnd = res.pendingUse.get(it.name) || 0, bal = it.remaining - pnd - u;
      return '<tr class="' + (u ? '' : 'acc-sub') + '"><td class="acc-item-col">' + escHtml(it.name) + '</td><td class="num">' + fmtQty(it.qty) +
        '</td><td class="num">' + fmtQty(it.remaining) + '</td><td class="num">' + (pnd ? fmtQty(pnd) : '—') + '</td><td class="num">' + (u ? fmtQty(u) : '—') + '</td><td class="num' +
        (bal < -0.005 ? ' acc-item-bad' : '') + '">' + fmtQty(bal) + '</td></tr>';
    }).join('') + '</tbody></table>');

  const comboTable = (arr, reasonOf) => '<table class="acc-table"><thead><tr><th>Job Code</th><th>Site ID</th><th>Acc. Week</th>' +
    '<th>Items</th><th>Amount</th><th>Tracking Rows</th><th>Reason</th></tr></thead><tbody>' +
    arr.map(c => '<tr><td>' + escHtml(c.jc) + '</td><td class="acc-site">' + escHtml(str(c.siteRaw)) + '</td><td>' + escHtml(c.weekLabel) +
      '</td><td class="acc-sub">' + escHtml(c.tasks.map(t => t.code).join(', ')) + '</td><td class="num">' + fmtEGP(c.amount) +
      '</td><td class="acc-sub">' + c.tasks.map(t => t.excelRow).join(', ') + '</td><td class="acc-notes-col">' +
      reasonOf(c).map(x => '<div class="acc-note">' + escHtml(x) + '</div>').join('') + '</td></tr>').join('') + '</tbody></table>';
  el('tsp-waiting-wrap').innerHTML = res.waiting.length
    ? section('Ready, waiting for TSR quantity', res.waiting.length + ' combos', fmtEGP(sum(res.waiting)), comboTable(res.waiting, c => [c.short])) : '';
  el('tsp-blocked-wrap').innerHTML = res.blocked.length
    ? section('Not ready — fix in the tracking', res.blocked.length + ' combos', fmtEGP(sum(res.blocked)),
      comboTable(res.blocked.slice(0, 400), c => c.reasons) +
      (res.blocked.length > 400 ? '<p class="acc-sub">First 400 shown — the full list is in the exported file.</p>' : '')) : '';

  el('tsp-btn-paste').disabled = !res.rows.length;
  el('tsp-btn-recv').disabled = !res.rows.length || !_tsr.receiving.length;
  el('tsp-btn-folders').disabled = !res.folders.length;
  setStatus('');
}

// ---------------------------------------------------------------------------
// Export 1 — paste-ready workbook
// ---------------------------------------------------------------------------
const THIN = { style: 'thin', color: { argb: 'FFBFBFBF' } };
const BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };
const fill = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
const BLOCKS = [['C', 'D'], ['H', 'J'], ['M', 'N'], ['P', 'P'], ['S', 'Y']];
const FORMULA_COLS = new Set(['E', 'F', 'G', 'K', 'L', 'O', 'Q', 'R']);
const BLOCK_FILL = ['FFDDEBF7', 'FFE2EFDA', 'FFFFF2CC', 'FFFCE4D6', 'FFEDE2F6'];

function fileTag(res) { return 'TSR_' + (res.tsrNo || 'TSR') + '_Sub-' + res.subNo; }

function headerSheet(wb, name, headers, widths) {
  const ws = wb.addWorksheet(name);
  headers.forEach((h, i) => { ws.getColumn(i + 1).width = widths[i] || 14; });
  const r = ws.addRow(headers);
  r.eachCell(c => { c.fill = fill('FF0070C0'); c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; c.border = BORDER;
                    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; });
  ws.views = [{ state: 'frozen', ySplit: 1 }];
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

async function exportPaste() {
  const res = _result; if (!res || !res.rows.length) return;
  const tsr = _tsr;
  const wb = new ExcelJS.Workbook(); wb.creator = 'LMP Invoicing System';
  const ws = wb.addWorksheet('TSR Paste', { properties: { tabColor: { argb: 'FF0070C0' } } });
  const pr = res.rows.map(r => pasteRow(r, res));
  const letters = []; for (let i = 0; i < 25; i++) letters.push(C.colLetter(i));  // A … Y
  const startRow = res.pasteRow, n = pr.length, first = 4, last = first + n - 1;

  ws.getCell('A1').value = 'TSR ' + res.tsrNo + ' — Submission #' + res.subNo + ' — ' + n + ' rows — paste into "PO Break Down- Contractor & Acc" starting at row ' + startRow;
  ws.getCell('A1').font = { bold: true, size: 13, color: { argb: 'FF1A3A5C' } };

  // Row 2: block labels; row 3: TSR headers (same letters as the TSR sheet)
  const blockOf = L => BLOCKS.findIndex(([a, b]) => L >= a && L <= b && L.length === 1);
  for (const L of letters) {
    const c2 = ws.getCell(L + '2'), c3 = ws.getCell(L + '3');
    const hi = C.colIndex(L);
    c3.value = L === 'A' ? '#' : L === 'B' ? 'Tracking Row' : (tsr.pbHeaders[hi] || '');
    const b = blockOf(L);
    if (b >= 0) {
      const [a, z] = BLOCKS[b];
      if (L === a) c2.value = 'Block ' + (b + 1) + ': copy ' + a + first + ':' + z + last + ' → paste at ' + a + startRow;
      c2.fill = fill(BLOCK_FILL[b]); c3.fill = fill(BLOCK_FILL[b]);
    } else {
      if (FORMULA_COLS.has(L)) c2.value = 'TSR formula — preview only';
      c2.fill = fill('FFD9D9D9'); c3.fill = fill('FFD9D9D9');
    }
    c2.font = { bold: true, size: 9 }; c2.alignment = { wrapText: true, vertical: 'top' };
    c3.font = { bold: true }; c3.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    c2.border = BORDER; c3.border = BORDER;
  }
  ws.getRow(2).height = 42; ws.getRow(3).height = 30;
  const widths = { A: 5, B: 9, C: 18, D: 11, E: 10, F: 10, G: 12, H: 12, I: 7, J: 11, K: 14, L: 10, M: 11, N: 48,
                   O: 10, P: 7, Q: 10, R: 12, S: 30, T: 18, U: 12, V: 9, W: 15, X: 12, Y: 10 };
  for (const [L, w] of Object.entries(widths)) ws.getColumn(L).width = w;

  pr.forEach((p, i) => {
    const r = first + i, row = res.rows[i];
    const vals = { A: i + 1, B: row.t.excelRow, ...p };
    for (const L of letters) {
      const c = ws.getCell(L + r);
      let v = vals[L];
      if (v === undefined || v === null) v = '';
      c.value = C.excelDate(v);
      if (v instanceof Date) c.numFmt = 'dd-mmm-yy';
      c.border = BORDER;
      if (FORMULA_COLS.has(L) || L === 'A' || L === 'B') {
        c.fill = fill('FFF2F2F2'); c.font = { italic: true, color: { argb: v === '#N/A' ? 'FFC00000' : 'FF7F7F7F' } };
      }
      if (row.t.fix && ((L === 'U' && row.t.fix.fixes.some(f => f.field === 'cert')) || (L === 'J' && row.t.fix.fixes.some(f => f.field === 'facing')))) {
        c.fill = fill('FFFFFF00');
      }
      if (L === 'S' && row.c.streamCheck && !_streamOverride.has(row.c.id)) c.fill = fill('FFFFC7CE');
    }
    ws.getCell('Q' + r).numFmt = '#,##0.00'; ws.getCell('R' + r).numFmt = '#,##0.00';
  });
  ws.views = [{ state: 'frozen', ySplit: 3, xSplit: 2 }];

  // How to paste
  const hp = wb.addWorksheet('How to paste');
  hp.getColumn(1).width = 110;
  [
    'The PO Break Down sheet is protected: Excel refuses a paste that touches a locked (formula) cell.',
    'So the rows come in 5 blocks with the SAME column letters as the TSR. Paste each block at the same TSR row:',
    ...BLOCKS.map(([a, z], b) => '   Block ' + (b + 1) + ':  in "TSR Paste" select ' + a + first + ':' + z + last + '  →  Copy  →  in the TSR click ' + a + startRow + '  →  Paste Values (Ctrl+Alt+V, V)'),
    'Grey columns (Region, Governorate, Distance, Physical Site, Unit, Unit Price, Total Value) are the TSR\'s own formulas — never paste them; they are shown here only as a preview.',
    'Red #N/A in Distance = the site is not in VF\'s site list inside the TSR; VF fills it after you send the file.',
    'Yellow = Certificate # / Facing corrected from the acceptance mail (fix the tracking the same way).',
    'Pink General Stream = rule could not decide — check it before pasting.',
    'Then: download the Receiving Template and the mail folders, and paste the Allocation sheet values into the tracking TSR Sub# column.'
  ].forEach((t, i) => { const c = hp.getCell('A' + (i + 1)); c.value = t; c.alignment = { wrapText: true }; if (i === 0) c.font = { bold: true }; });

  // Allocation — what to write in the tracking
  const al = headerSheet(wb, 'Allocation', ['Tracking Row', 'ID#', 'Job Code', 'Site ID', 'Line Item', 'TSR Sub# (write in tracking)', 'Folder', 'Certificate # (mail)', 'Facing (mail)'],
    [12, 22, 12, 12, 48, 24, 8, 18, 12]);
  addRows(al, res.rows.slice().sort((a, b) => a.t.excelRow - b.t.excelRow).map(r => [r.t.excelRow, r.t.id, r.c.jc, r.t.siteRaw, r.t.item,
    res.tsrNo + ' sub-' + res.subNo, r.folder, r.t.fix ? r.t.fix.cert : r.t.cert, r.t.fix ? r.t.fix.facing : r.t.facingRaw]));

  // Quantity check (formulas so it recalculates)
  const qs = headerSheet(wb, 'Qty Check', ['Item', 'TSR Qty', 'Remaining (Request Form)', 'Sent, not yet received', 'This submission', 'Balance'], [60, 10, 14, 14, 14, 12]);
  const usedQty = new Map();
  for (const c of res.planned) for (const [nm, q] of c.need) usedQty.set(nm, (usedQty.get(nm) || 0) + q);
  tsr.items.filter(it => it.qty || it.remaining || usedQty.get(it.name) || res.pendingUse.get(it.name)).forEach(it => {
    const r = qs.addRow([it.name, it.qty, round4(it.remaining), round4(res.pendingUse.get(it.name) || 0), round4(usedQty.get(it.name) || 0), null]);
    r.getCell(6).value = { formula: 'C' + r.number + '-D' + r.number + '-E' + r.number };
    r.eachCell({ includeEmpty: true }, c => { c.border = BORDER; });
    [3, 4, 5, 6].forEach(k => { r.getCell(k).numFmt = '0.00'; });
  });

  const comboRows = (arr, reasonOf) => arr.map(c => [c.jc, c.siteRaw, c.weekLabel, c.tasks.map(t => t.code).join(', '), c.amount, reasonOf(c), c.tasks.map(t => t.excelRow).join(', ')]);
  const cw = [12, 12, 18, 30, 14, 90, 30];
  const cH = ['Job Code', 'Site ID', 'Acceptance Week', 'Items', 'Amount (EGP)', 'Reason', 'Tracking Rows'];
  if (res.waiting.length) addRows(headerSheet(wb, 'Waiting for Qty', cH, cw), comboRows(res.waiting, c => c.short), [0, 0, 0, 0, '#,##0']);
  if (res.blocked.length) addRows(headerSheet(wb, 'Not Ready', cH, cw), comboRows(res.blocked, c => c.reasons.join(' | ')), [0, 0, 0, 0, '#,##0']);
  if (res.pending.length) addRows(headerSheet(wb, 'In Another TSR', cH, cw), comboRows(res.pending, c => c.reasons.join(' | ')), [0, 0, 0, 0, '#,##0']);
  if (res.tocTasks.length) {
    addRows(headerSheet(wb, 'TOC Track (later)', ['Tracking Row', 'Job Code', 'Site ID', 'Line Item', 'Status', 'Acceptance Status', 'Acceptance Week', 'Amount (EGP)'], [12, 12, 12, 48, 10, 16, 18, 14]),
      res.tocTasks.map(t => [t.excelRow, t.jc, t.siteRaw, t.item, t.statusRaw, t.accStatus, t.weekRaw, t.total]), [0, 0, 0, 0, 0, 0, 0, '#,##0']);
  }

  const name = fileTag(res) + '_Paste_' + new Date().toISOString().slice(0, 10) + '.xlsx';
  C.downloadBlob(new Blob([await wb.xlsx.writeBuffer()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), name);
  setStatus('✅ Downloaded ' + name + ' — ' + n + ' rows in 5 paste blocks, starting at TSR row ' + startRow + '.');
}

// ---------------------------------------------------------------------------
// Export 2 — receiving template for VF invoicing (same layout as LMP's file)
// ---------------------------------------------------------------------------
async function exportReceiving() {
  const res = _result; if (!res || !res.rows.length) return;
  const tsr = _tsr;
  const qty = new Map();   // receiving code → actual qty (TSR row qty × distance factor)
  for (const r of res.rows) {
    const code = C.itemCode(r.t.tsrItem.name);
    qty.set(code, (qty.get(code) || 0) + r.t.effQty);
  }
  const lines = tsr.receiving.filter(x => (qty.get(x.code) || 0) > 0);
  const missing = [...qty.keys()].filter(k => !tsr.receiving.some(x => x.code === k));
  if (missing.length) throw new Error('Not in the TSR Receiving Template sheet: ' + missing.join(', '));

  const wb = new ExcelJS.Workbook(); wb.creator = 'LMP Invoicing System';
  const ws = wb.addWorksheet('Receiving');
  const F = { name: 'Aptos Narrow', size: 11 };
  const H = ['#', 'Vendor Name', 'Vendor ID', 'Description', 'Unit Cost', 'PO#', 'Line Row', 'QTY', 'Total Value', 'PO unit price', 'EVO Mat. No.', 'Stream Name'];
  const W = [5.55, 24.66, 11.33, 35, 10.89, 11, 10.55, 8.78, 11.78, 11.66, 19.33, 44.33];
  W.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  const hdr = ws.getRow(2);
  H.forEach((h, i) => { const c = hdr.getCell(i + 1); c.value = h; c.font = { ...F, bold: true }; c.fill = fill('FFC0E6F5');
                        if (i === 0) c.alignment = { horizontal: 'center' }; });
  const po = /^\d+$/.test(res.poNo) ? Number(res.poNo) : res.poNo;
  lines.forEach((x, i) => {
    const r = ws.getRow(3 + i);
    const vals = [i + 1, x.vendor, Number(x.vendorId) || x.vendorId, x.desc, x.cost, po, 1, round4(qty.get(x.code)), null, 1, x.evo, x.stream];
    vals.forEach((v, j) => { const c = r.getCell(j + 1); c.value = v; c.font = F; });
    r.getCell(9).value = { formula: 'E' + r.number + '*H' + r.number };
    [1, 7, 8, 9, 10].forEach(k => { r.getCell(k).alignment = { horizontal: 'center' }; });
    r.getCell(4).alignment = { horizontal: 'left' };
    r.getCell(5).numFmt = '#,##0.00'; r.getCell(9).numFmt = '#,##0.00';
  });
  const tr = ws.getRow(3 + lines.length);
  for (let j = 1; j <= 12; j++) { const c = tr.getCell(j); c.font = { ...F, bold: true }; c.fill = fill('FFC0E6F5'); }
  tr.getCell(2).value = 'Grand Total';
  tr.getCell(9).value = { formula: 'SUM(I3:I' + (2 + lines.length) + ')' };
  tr.getCell(9).numFmt = '#,##0.00'; tr.getCell(9).alignment = { horizontal: 'center' };
  ws.autoFilter = 'A2:L' + (3 + lines.length);

  const name = 'Receiving Template LMP TSR ' + res.tsrNo + ' Sub-' + res.subNo + '.xlsx';
  C.downloadBlob(new Blob([await wb.xlsx.writeBuffer()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), name);
  const total = lines.reduce((s, x) => s + x.cost * qty.get(x.code), 0);
  setStatus('✅ Downloaded ' + name + ' — ' + lines.length + ' items, ' + fmtEGP(total) + (res.poNo ? '' : ' — PO# is blank, fill it in before sending.'));
}

// ---------------------------------------------------------------------------
// Export 3 — mail folders 1, 2, 3… (zipped)
// ---------------------------------------------------------------------------
async function exportFolders() {
  const res = _result; if (!res || !res.folders.length) return;
  const JSZip = await C.getJsZip();
  if (!JSZip) throw new Error('ZIP library could not be loaded. Check your internet connection.');
  setStatus('Preparing ' + res.folders.length + ' folders…');
  const zip = new JSZip();
  res.folders.forEach((m, i) => zip.folder(String(i + 1)).file(m.file.name, m.file));
  const blob = await zip.generateAsync({ type: 'blob' });
  const name = fileTag(res) + '_Mails_' + new Date().toISOString().slice(0, 10) + '.zip';
  C.downloadBlob(blob, name);
  setStatus('✅ Downloaded ' + name + ' — folders 1–' + res.folders.length + ', one mail each.');
}

// Shared with the New PO section (tsr-newpo-app.js): the loaded files
window.TSRPlanner = { get trk() { return _trk; }, get tsr() { return _tsr; } };

// Exposed for testing in a headless browser
window.__tspDebug = { get result() { return _result; }, get tsr() { return _tsr; }, get trk() { return _trk; }, pasteRow };
})();
