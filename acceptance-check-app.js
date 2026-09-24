// =============================================================================
// LMP Invoicing System — Acceptance Check logic
// Compares the line items of one acceptance week + area in the Tracking file
// against the "Total" sheet of the Acceptance sheet and lists every mismatch.
// Wrapped in IIFE to avoid global namespace collisions.
// =============================================================================
(function () {

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------
let _trk     = null; // parsed tracking file   (readTracking)
let _acc     = null; // parsed acceptance sheet (readAcceptance)
let _results = [];   // last comparison, re-rendered when "show matched" toggles

// Area selector → tracking week prefix + acceptance-sheet Area values
const AREAS = {
  delta: { label: 'Delta',            prefix: 'D', names: ['delta'] },
  cgu:   { label: 'Cairo-Giza-Upper', prefix: 'U', names: ['cairo', 'giza', 'upper'] },
  alex:  { label: 'Alex',             prefix: 'A', names: ['alex'] }
};

const ISSUE_LABELS = {
  match:         'Match',
  mismatch:      'Line item mismatch',
  trk_only:      'Missing in acceptance',
  acc_only:      'Missing in tracking',
  site_trk_only: 'Site not in acceptance',
  site_acc_only: 'Site not in tracking'
};

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
const el = id => document.getElementById(id);

function str(v) { return v == null ? '' : String(v).trim(); }

// Header text normaliser: "Item_Description" → "item description"
function normHeader(v) {
  return str(v).toLowerCase().replace(/[_\s]+/g, ' ').trim();
}

// Site ID key: "U3844" → "U3844", 3688 → "3688", "0568" → "568"
function siteKey(v) {
  if (v == null || v === '') return '';
  if (typeof v === 'number') return String(Math.round(v));
  const s = String(v).toUpperCase().replace(/\s+/g, '').replace(/\.0+$/, '');
  return /^\d+$/.test(s) ? String(parseInt(s, 10)) : s;
}

// Digits only — used to hint at near-identical Site IDs ("3844" vs "U3844")
function siteDigits(v) {
  return siteKey(v).replace(/\D/g, '').replace(/^0+/, '');
}

// Catalogue code of a line item: "TX03 - MW Link E-Band 1+0" → "TX03".
// Items without a code fall back to their normalised full text.
function itemKey(desc) {
  const s = str(desc);
  const m = s.match(/^([A-Za-z]{1,4})\s*[-.]?\s*(\d{1,3})\b/);
  if (m) return m[1].toUpperCase() + String(parseInt(m[2], 10)).padStart(2, '0');
  return s.toLowerCase().replace(/\s+/g, ' ');
}

// Letter family of a code ("TX03" → "TX") — used to pair mismatching items
function itemFamily(key) {
  const m = key.match(/^([A-Z]+)\d+$/);
  return m ? m[1] : null;
}

// Tracking Acceptance Week: "D-W36-W37-2026", "U-W09-2026/U-W15-2026"
// → [{ prefix: 'D', weeks: [36, 37], year: 2026 }, ...]
function parseTrackingWeek(raw) {
  const s = str(raw).toUpperCase();
  if (!s) return [];
  return s.split(/[\/,;&]+/).map(seg => {
    const p = seg.match(/^\s*([A-Z]+)\s*-\s*W\s*\d/);
    const y = seg.match(/(20\d{2})/);
    const weeks = [];
    const re = /W\s*(\d{1,2})/g;
    let m;
    while ((m = re.exec(seg)) !== null) weeks.push(parseInt(m[1], 10));
    return { prefix: p ? p[1] : '', weeks, year: y ? parseInt(y[1], 10) : null };
  }).filter(seg => seg.weeks.length > 0);
}

// Acceptance-sheet Week: 34, "36--37", "36-37", "36/37".
// Excel sometimes turns "3-4" into a date (3-Apr) — day and month are the two weeks.
function parseAccWeek(v) {
  if (v == null || v === '') return [];
  if (v instanceof Date) return [v.getDate(), v.getMonth() + 1];
  if (typeof v === 'number') return [Math.round(v)];
  return (String(v).match(/\d+/g) || []).map(Number).filter(n => n >= 1 && n <= 53);
}

function accWeekText(v) {
  if (v instanceof Date) return v.getDate() + '-' + (v.getMonth() + 1);
  return str(v);
}

function areaMatches(areaCfg, raw) {
  const a = str(raw).toLowerCase();
  return a !== '' && areaCfg.names.some(n => a.startsWith(n));
}

function findSheet(wb, name) {
  return wb.SheetNames.find(n => n.trim().toLowerCase() === name) || null;
}

function colLetter(idx) {
  let letter = '', n = idx + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}

function escHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function showError(msg) {
  const d = el('acc-error');
  d.textContent = msg; d.style.display = 'block';
  d.scrollIntoView({ behavior: 'smooth', block: 'center' });
}
function clearError() {
  const d = el('acc-error');
  d.textContent = ''; d.style.display = 'none';
}
function showProgress(id, show) {
  const e = el(id); if (e) e.style.display = show ? 'block' : 'none';
}
function checkReady() {
  el('acc-btn-check').disabled = !(_trk && _acc
    && el('acc-area').value !== '' && parseWeekInput(el('acc-week').value).length > 0);
}

// "36, 37" / "36-37" / "36 37" → [36, 37]
function parseWeekInput(text) {
  const nums = (String(text).match(/\d+/g) || []).map(Number).filter(n => n >= 1 && n <= 53);
  return [...new Set(nums)].sort((a, b) => a - b);
}

function selectedYear() {
  return parseInt(el('acc-year').value, 10) || null;
}

function trkSegInScope(seg, areaCfg, year) {
  return seg.prefix === areaCfg.prefix && (!year || seg.year === null || seg.year === year);
}

// ---------------------------------------------------------------------------
// File picking
// ---------------------------------------------------------------------------
function wireFilePicker(btnId, inputId, cardId, nameId, progressId, onLoad) {
  el(btnId).addEventListener('click', () => el(inputId).click());
  el(inputId).addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      clearError();
      showProgress(progressId, true);
      const data = await file.arrayBuffer();
      onLoad(XLSX.read(data, { type: 'array', cellDates: true }));
      el(nameId).textContent = file.name;
      el(cardId).classList.add('loaded');
    } catch (err) {
      showError('Failed to open ' + file.name + ': ' + err.message);
    } finally {
      showProgress(progressId, false);
      renderWeekChips();
      checkReady();
    }
  });
}

wireFilePicker('acc-btn-track', 'acc-input-track', 'acc-card-track',
  'acc-track-filename', 'acc-track-progress', wb => { _trk = readTracking(wb); });
wireFilePicker('acc-btn-acc', 'acc-input-acc', 'acc-card-acc',
  'acc-acc-filename', 'acc-acc-progress', wb => { _acc = readAcceptance(wb); });

el('acc-year').value = new Date().getFullYear();
['acc-area', 'acc-year', 'acc-week'].forEach(id => {
  el(id).addEventListener(id === 'acc-area' ? 'change' : 'input', () => {
    renderWeekChips();
    checkReady();
  });
});

// ---------------------------------------------------------------------------
// Week picker — clickable chips for every week found in either file for the
// selected area/year. Chips and the Weeks text box stay in sync both ways.
// ---------------------------------------------------------------------------
function availableWeeks() {
  const areaCfg = AREAS[el('acc-area').value];
  if (!areaCfg) return null;
  const year = selectedYear();
  const map  = new Map(); // week → { trk, acc } item counts
  const bump = (w, side) => {
    if (!map.has(w)) map.set(w, { trk: 0, acc: 0 });
    map.get(w)[side]++;
  };
  if (_trk) _trk.rows.forEach(r => {
    const ws = new Set();
    r.weekSegs.forEach(s => { if (trkSegInScope(s, areaCfg, year)) s.weeks.forEach(w => ws.add(w)); });
    ws.forEach(w => bump(w, 'trk'));
  });
  if (_acc) _acc.rows.forEach(r => {
    if (areaMatches(areaCfg, r.area)) new Set(r.weeks).forEach(w => bump(w, 'acc'));
  });
  return [...map.entries()].sort((a, b) => a[0] - b[0]);
}

function renderWeekChips() {
  const box   = el('acc-week-chips');
  const weeks = availableWeeks();
  if (!weeks || !(_trk || _acc)) {
    box.innerHTML = '<span class="acc-chips-hint">Load a file and select an area to pick weeks from here &mdash; or type them in the Weeks box.</span>';
    return;
  }
  if (!weeks.length) {
    box.innerHTML = '<span class="acc-chips-hint">No weeks found for this area and year.</span>';
    return;
  }
  const selected = new Set(parseWeekInput(el('acc-week').value));
  box.innerHTML = weeks.map(([w, c]) =>
    '<button type="button" class="acc-chip' + (selected.has(w) ? ' active' : '') +
      (c.trk && c.acc ? '' : ' acc-chip-partial') + '" data-week="' + w + '"' +
      ' title="Tracking: ' + c.trk + ' item' + (c.trk !== 1 ? 's' : '') +
      ' · Acceptance: ' + c.acc + ' item' + (c.acc !== 1 ? 's' : '') + '">' + w + '</button>'
  ).join('') +
  (weeks.some(([, c]) => !(c.trk && c.acc))
    ? '<span class="acc-chips-hint">Dashed = week found in only one of the two files.</span>' : '');
}

function setWeeks(list) {
  el('acc-week').value = [...new Set(list)].sort((a, b) => a - b).join(', ');
  renderWeekChips();
  checkReady();
}

el('acc-week-chips').addEventListener('click', e => {
  const chip = e.target.closest('.acc-chip');
  if (!chip) return;
  const w   = Number(chip.dataset.week);
  const cur = parseWeekInput(el('acc-week').value);
  setWeeks(cur.includes(w) ? cur.filter(x => x !== w) : cur.concat(w));
});
el('acc-weeks-all').addEventListener('click', () => setWeeks((availableWeeks() || []).map(([w]) => w)));
el('acc-weeks-clear').addEventListener('click', () => setWeeks([]));

renderWeekChips();

el('acc-week').addEventListener('keydown', e => {
  if (e.key === 'Enter' && !el('acc-btn-check').disabled) el('acc-btn-check').click();
});
el('acc-show-matched').addEventListener('change', () => renderTable(_results));

el('acc-btn-check').addEventListener('click', async () => {
  clearError();
  el('acc-results').style.display = 'none';
  el('acc-loading').style.display = 'flex';
  await new Promise(r => setTimeout(r, 30));
  try {
    const out = runCheck();
    _results = out.results;
    renderResults(out);
    el('acc-results').style.display = 'block';
    el('acc-results').scrollIntoView({ behavior: 'smooth' });
  } catch (err) {
    showError(err.message || 'An unexpected error occurred.');
  } finally {
    el('acc-loading').style.display = 'none';
  }
});

// ---------------------------------------------------------------------------
// Readers
// ---------------------------------------------------------------------------
function sheetRows(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: true });
  const r0   = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']).s.r : 0;
  return { rows, r0 };
}

// Tracking file — sheet "Invoicing Track"; header row located by name
function readTracking(wb) {
  const name = findSheet(wb, 'invoicing track');
  if (!name) {
    throw new Error('Tracking file: sheet "Invoicing Track" not found. Available: ' +
      wb.SheetNames.join(', '));
  }
  const { rows, r0 } = sheetRows(wb.Sheets[name]);

  let hdr = -1;
  for (let i = 0; i < Math.min(rows.length, 30); i++) {
    const h = (rows[i] || []).map(normHeader);
    if (h.some(t => t.includes('logical site')) && h.some(t => t.includes('acceptance week'))) {
      hdr = i; break;
    }
  }
  if (hdr < 0) {
    throw new Error('Tracking file: could not find the header row ' +
      '(a row containing both "Logical Site ID" and "Acceptance Week").');
  }

  const col = { site: -1, item: -1, week: -1, status: -1, facing: -1, job: -1 };
  (rows[hdr] || []).forEach((cell, c) => {
    const t = normHeader(cell);
    if (col.site   < 0 && t.includes('logical site'))             col.site   = c;
    if (col.item   < 0 && (t === 'line item' || t === 'line items')) col.item = c;
    if (col.week   < 0 && t.includes('acceptance week'))          col.week   = c;
    if (col.status < 0 && t === 'status')                         col.status = c;
    if (col.facing < 0 && t.includes('facing'))                   col.facing = c;
    if (col.job    < 0 && t.includes('job code'))                 col.job    = c;
  });
  if (col.item < 0) throw new Error('Tracking file: "Line Item" column not found.');

  const out = [];
  for (let i = hdr + 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const site = siteKey(row[col.site]);
    if (!site) continue;
    const status = col.status >= 0 ? str(row[col.status]).toLowerCase() : '';
    if (status === 'cancelled' || status === 'canceled') continue;
    const item = str(row[col.item]);
    out.push({
      excelRow: i + r0 + 1,
      site, siteRaw: str(row[col.site]),
      item, key: itemKey(item),
      facing: col.facing >= 0 ? siteKey(row[col.facing]) : '',
      jobCode: col.job >= 0 ? str(row[col.job]) : '',
      weekRaw: str(row[col.week]),
      weekSegs: parseTrackingWeek(row[col.week])
    });
  }
  return { rows: out, col, headerRow: hdr + r0 + 1, sheet: name };
}

// Acceptance sheet — tab "Total"; the table can start anywhere on the sheet
function readAcceptance(wb) {
  const name = findSheet(wb, 'total');
  if (!name) {
    throw new Error('Acceptance sheet: tab "Total" not found. Available: ' +
      wb.SheetNames.join(', '));
  }
  const { rows, r0 } = sheetRows(wb.Sheets[name]);

  let hdr = -1;
  for (let i = 0; i < rows.length; i++) {
    const h = (rows[i] || []).map(normHeader);
    if (h.some(t => t === 'site id' || t.startsWith('site id')) &&
        h.some(t => t.includes('item description'))) {
      hdr = i; break;
    }
  }
  if (hdr < 0) {
    throw new Error('Acceptance sheet: could not find the table header ' +
      '(a row containing both "Site_ID" and "Item_Description") in tab "' + name + '".');
  }

  const col = { site: -1, item: -1, week: -1, area: -1, req: -1, facing: -1 };
  (rows[hdr] || []).forEach((cell, c) => {
    const t = normHeader(cell);
    if (col.site   < 0 && (t === 'site id' || t.startsWith('site id'))) col.site = c;
    if (col.item   < 0 && t.includes('item description'))             col.item   = c;
    if (col.week   < 0 && /^week\b/.test(t))                           col.week   = c;
    if (col.area   < 0 && t === 'area')                                col.area   = c;
    if (col.req    < 0 && t.includes('request'))                       col.req    = c;
    if (col.facing < 0 && t.includes('facing'))                        col.facing = c;
  });
  const missing = [];
  if (col.week < 0) missing.push('"Week"');
  if (col.area < 0) missing.push('"Area"');
  if (missing.length) {
    throw new Error('Acceptance sheet: missing column(s) ' + missing.join(', ') +
      ' in the header row (row ' + (hdr + r0 + 1) + ').');
  }

  const out = [];
  for (let i = hdr + 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const site = siteKey(row[col.site]);
    const item = str(row[col.item]);
    if (!site || !item) continue;
    out.push({
      excelRow: i + r0 + 1,
      site, siteRaw: str(row[col.site]),
      item, key: itemKey(item),
      facing: col.facing >= 0 ? siteKey(row[col.facing]) : '',
      reqId: col.req >= 0 ? str(row[col.req]) : '',
      area: str(row[col.area]),
      weekRaw: accWeekText(row[col.week]),
      weeks: parseAccWeek(row[col.week])
    });
  }
  return { rows: out, col, headerRow: hdr + r0 + 1, sheet: name };
}

// ---------------------------------------------------------------------------
// Comparison
// ---------------------------------------------------------------------------
function runCheck() {
  const areaCfg = AREAS[el('acc-area').value];
  if (!areaCfg) throw new Error('Please select an area.');

  const weekList = parseWeekInput(el('acc-week').value);
  if (!weekList.length) throw new Error('Please select or enter at least one week (1–53).');
  const weekSet = new Set(weekList);

  const year = selectedYear();
  const trk  = _trk;
  const acc  = _acc;

  const trkInScope = r => r.weekSegs.some(s =>
    trkSegInScope(s, areaCfg, year) && s.weeks.some(w => weekSet.has(w)));
  const accInScope = r => areaMatches(areaCfg, r.area) && r.weeks.some(w => weekSet.has(w));

  const trkScope = trk.rows.filter(trkInScope);
  const accScope = acc.rows.filter(accInScope);

  const groupBy = list => {
    const m = new Map();
    list.forEach(r => { if (!m.has(r.site)) m.set(r.site, []); m.get(r.site).push(r); });
    return m;
  };
  const tMap = groupBy(trkScope);
  const aMap = groupBy(accScope);
  const trkAll = groupBy(trk.rows);
  const accAll = groupBy(acc.rows);

  // Where else does this site + item appear in the tracking file? (wrong week)
  function trkHints(a) {
    return (trkAll.get(a.site) || [])
      .filter(t => t.key === a.key && !trkInScope(t))
      .map(t => 'Tracking has this item under "' + (t.weekRaw || 'no week') + '" (row ' + t.excelRow + ')');
  }
  // Where else does this site + item appear in the acceptance sheet?
  function accHints(t) {
    return (accAll.get(t.site) || [])
      .filter(a => a.key === t.key && !accInScope(a))
      .map(a => 'Acceptance sheet has this item in week ' + (a.weekRaw || '?') +
                ' / ' + (a.area || '?') + ' (row ' + a.excelRow + ')');
  }
  function similarSite(site, map) {
    const d = siteDigits(site);
    if (!d) return null;
    for (const k of map.keys()) if (k !== site && siteDigits(k) === d) return k;
    return null;
  }

  const results = [];
  const sites = [...new Set([...tMap.keys(), ...aMap.keys()])]
    .sort((x, y) => x.localeCompare(y, undefined, { numeric: true }));

  for (const site of sites) {
    const T = tMap.get(site) || [];
    const A = aMap.get(site) || [];

    if (!A.length) {
      const sim = similarSite(site, aMap);
      T.forEach(t => results.push({
        site, type: 'site_trk_only', t, a: null,
        notes: (sim ? ['Acceptance sheet has a similar Site ID "' + sim + '" in this week'] : [])
          .concat(accHints(t))
      }));
      continue;
    }
    if (!T.length) {
      const sim = similarSite(site, tMap);
      A.forEach(a => results.push({
        site, type: 'site_acc_only', t: null, a,
        notes: (sim ? ['Tracking has a similar Site ID "' + sim + '" in this week'] : [])
          .concat(trkHints(a))
      }));
      continue;
    }

    // 1. Exact item matches (one acceptance row per tracking row)
    const aLeft = [...A];
    const tLeft = [];
    for (const t of T) {
      const idx = aLeft.findIndex(a => a.key === t.key);
      if (idx >= 0) results.push({ site, type: 'match', t, a: aLeft.splice(idx, 1)[0], notes: [] });
      else tLeft.push(t);
    }

    // 2. Pair the leftovers — prefer same code family (TX↔TX), then same facing
    for (const t of tLeft) {
      if (!aLeft.length) {
        results.push({ site, type: 'trk_only', t, a: null, notes: accHints(t) });
        continue;
      }
      let best = 0, bestScore = -1;
      aLeft.forEach((a, i) => {
        const fam = itemFamily(t.key);
        const score = (fam && fam === itemFamily(a.key) ? 2 : 0) +
                      (t.facing && t.facing === a.facing ? 1 : 0);
        if (score > bestScore) { bestScore = score; best = i; }
      });
      const a = aLeft.splice(best, 1)[0];
      results.push({ site, type: 'mismatch', t, a, notes: accHints(t).concat(trkHints(a)) });
    }
    aLeft.forEach(a => results.push({ site, type: 'acc_only', t: null, a, notes: trkHints(a) }));
  }

  return {
    results, trk, acc, trkScope, accScope,
    areaLabel: areaCfg.label, areaPrefix: areaCfg.prefix,
    weekLabel: weekList.join(', '), weekCount: weekList.length, year
  };
}

// ---------------------------------------------------------------------------
// Rendering
// ---------------------------------------------------------------------------
function renderResults(out) {
  const { results, trk, acc, trkScope, accScope } = out;
  const issues  = results.filter(r => r.type !== 'match');
  const matched = results.length - issues.length;
  const sites   = new Set(results.map(r => r.site)).size;
  const badSites = new Set(issues.map(r => r.site)).size;

  el('acc-col-info').textContent =
    'Tracking [' + trk.sheet + ', header row ' + trk.headerRow + ']: Site=' + colLetter(trk.col.site) +
    '  Line Item=' + colLetter(trk.col.item) + '  Acc. Week=' + colLetter(trk.col.week) +
    '   |   Acceptance [' + acc.sheet + ', header row ' + acc.headerRow + ']: Site_ID=' + colLetter(acc.col.site) +
    '  Item=' + colLetter(acc.col.item) + '  Area=' + colLetter(acc.col.area) +
    '  Week=' + colLetter(acc.col.week);

  let html =
    '<p class="acc-scope">Week' + (out.weekCount > 1 ? 's' : '') +
    ' <strong>' + escHtml(out.weekLabel) + '</strong>' +
    (out.year ? ' / ' + out.year : '') + ' &mdash; <strong>' + escHtml(out.areaLabel) +
    '</strong> (tracking weeks starting with <code>' + out.areaPrefix + '-</code>)</p>' +
    '<div class="poc-stats-row acc-stats-row">' +
      statBox('Sites Checked', sites, '') +
      statBox('Tracking Items', trkScope.length, '') +
      statBox('Acceptance Items', accScope.length, '') +
      statBox('Matched', matched, 'acc-stat-pass') +
      statBox('Issues', issues.length, issues.length ? 'acc-stat-fail' : 'acc-stat-pass') +
    '</div>';

  if (!trkScope.length && !accScope.length) {
    html += '<div class="acc-banner acc-banner-fail">No rows found for this week and area in either file. ' +
            'Check the week number, year and area.</div>';
  } else if (!issues.length) {
    html += '<div class="acc-banner acc-banner-ok">&#10003; All ' + matched +
            ' line items match between the tracking file and the acceptance sheet.</div>';
  } else {
    html += '<div class="acc-banner acc-banner-fail">&#9888; ' + issues.length + ' issue' +
            (issues.length !== 1 ? 's' : '') + ' found across ' + badSites + ' site' +
            (badSites !== 1 ? 's' : '') + '. Use the tracking row numbers to correct the master tracking.</div>';
  }
  el('acc-summary').innerHTML = html;
  renderTable(results);
}

function statBox(label, value, cls) {
  return '<div class="poc-stat-box ' + cls + '">' +
    '<span class="poc-stat-label">' + label + '</span>' +
    '<span class="poc-stat-value">' + value + '</span></div>';
}

function renderTable(results) {
  const showMatched = el('acc-show-matched').checked;
  const rows = results.filter(r => showMatched || r.type !== 'match');
  const wrap = el('acc-table-wrap');

  if (!rows.length) {
    wrap.innerHTML = results.length
      ? '<p class="card-hint">No issues to show. Tick "Show matched rows" to see every line item.</p>'
      : '';
    return;
  }

  let html =
    '<table class="acc-table"><thead><tr>' +
      '<th>Site ID</th><th>Status</th>' +
      '<th>Tracking Row</th><th>Tracking Line Item</th><th>Tracking Week</th>' +
      '<th>Acceptance Row</th><th>Acceptance Line Item</th><th>Acceptance Week</th>' +
      '<th>Notes</th>' +
    '</tr></thead><tbody>';

  for (const r of rows) {
    const ok = r.type === 'match';
    const t = r.t, a = r.a;
    html +=
      '<tr class="' + (ok ? 'acc-row-pass' : 'acc-row-fail') + '">' +
        '<td class="acc-site">' + escHtml(t ? t.siteRaw : a.siteRaw) + '</td>' +
        '<td><span class="acc-badge acc-badge-' + r.type + '">' + ISSUE_LABELS[r.type] + '</span></td>' +
        '<td class="num">' + (t ? t.excelRow : '&mdash;') + '</td>' +
        '<td class="acc-item-col' + (t && !ok ? ' acc-item-bad' : '') + '">' +
          (t ? escHtml(t.item || '(blank)') : '&mdash;') + '</td>' +
        '<td>' + (t ? escHtml(t.weekRaw) : '&mdash;') + '</td>' +
        '<td class="num">' + (a ? a.excelRow + (a.reqId ? '<div class="acc-sub">Req ' + escHtml(a.reqId) + '</div>' : '') : '&mdash;') + '</td>' +
        '<td class="acc-item-col">' + (a ? escHtml(a.item) : '&mdash;') + '</td>' +
        '<td>' + (a ? escHtml(a.weekRaw) + (a.area ? '<div class="acc-sub">' + escHtml(a.area) + '</div>' : '') : '&mdash;') + '</td>' +
        '<td class="acc-notes-col">' + r.notes.map(n => '<div class="acc-note">' + escHtml(n) + '</div>').join('') + '</td>' +
      '</tr>';
  }
  html += '</tbody></table>';
  wrap.innerHTML = html;
}

})(); // end IIFE
