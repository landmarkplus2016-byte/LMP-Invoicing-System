// =============================================================================
// LMP Invoicing System — TSR Sub Order logic
// Collects the tracking tasks of one TSR Sub# in the order they appear in the
// weekly acceptance mails (Upper-Cairo / Alex / Delta folders), caps every line
// item by the TSR remaining quantity and picks tasks up to a target amount.
// Wrapped in IIFE to avoid global namespace collisions.
// =============================================================================
(function () {

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------
let _trk       = null;       // readTracking result
let _tsr       = null;       // readTsr result
let _mails     = [];         // mail files with area / weeks parsed from their names
let _ignored   = [];         // mail-folder files whose area or week could not be read
let _mailCache = new Map();  // path → parsed attachment { rows, … } or { error }
let _result    = null;       // last run, used by the export
let _MsgReader = null;       // cached @kenjiuno/msgreader

const AREA_LABEL = { A: 'Alex', D: 'Delta', U: 'Upper-Cairo' };
const AREA_ORDER = { A: 0, D: 1, U: 2 };

// Distance band → quantity multiplier (same bands as TSR Sub Prep)
const DISTANCE_MULTIPLIERS = {
  '0km-100km':   1.0,
  '100km-400km': 1.1,
  '400km-800km': 1.2,
  '>800km':      1.25
};

const MAIL_EXT  = /\.(msg|eml|xlsx|xls|xlsm|xlsb)$/i;
const EXCEL_EXT = /\.(xlsx|xls|xlsm|xlsb)$/i;

// Output columns: [header, tracking column key, width]
const OUT_COLS = [
  ['ID#',               'id',        12],
  ['VF Task Owner',     'owner',     18],
  ['Vendor',            'vendor',    14],
  ['Logical Site ID',   'site',      16],
  ['Site Option',       'option',    13],
  ['Facing',            'facing',    12],
  ['Task Date',         'taskDate',  13],
  ['Line Item',         'item',      55],
  ['Absolute Quantity', 'absQty',    18],
  ['PRQ',               'prq',       12],
  ['Certificate #',     'cert',      16],
  ['Acceptance Status', 'accStatus', 20]
];

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

// Catalogue code of a line item: "TX09 - IDU Upgrade" → "TX09", "EX6" → "EX06".
// Items without a code fall back to their normalised full text.
function itemKey(desc) {
  const s = str(desc);
  const m = s.match(/^([A-Za-z]{1,4})\s*[-.]?\s*(\d{1,3})\b/);
  if (m) return m[1].toUpperCase() + String(parseInt(m[2], 10)).padStart(2, '0');
  return s.toLowerCase().replace(/\s+/g, ' ');
}

// TSR Sub# cell → comparable key: 2, "2", "2.0" → "2"; text is lower-cased
function subKey(v) {
  if (v == null || v instanceof Date) return '';
  if (typeof v === 'number') return String(v);
  const s = String(v).trim().replace(/\s+/g, ' ');
  return /^\d+\.0+$/.test(s) ? String(parseInt(s, 10)) : s.toLowerCase();
}

// Tracking Acceptance Week: "D-W38-W39-2026", "U-W09-2026/U-W15-2026"
// → [{ prefix: 'D', weeks: [38, 39], year: 2026 }, …]
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

// Mail name → weeks + year. Every number after a "W" is a week:
//   "Landmark -- G-Cairo  Upper -- Telecom Activities Status W34-2026" → [34], 2026
//   "RE Delta Telecom Activities  Status  ll LM+ ll 26W03-W04"         → [3, 4], 2026
//   "Alex Telecom Activities  Status  ll LM+ ll 26W39"                 → [39], 2026
function parseMailWeeks(name) {
  const s = name.toUpperCase();
  const weeks = [];
  const re = /(?:^|[^A-Z])W\s*(\d{1,2})(?!\d)/g;
  let m;
  while ((m = re.exec(s)) !== null) {
    const w = parseInt(m[1], 10);
    if (w >= 1 && w <= 53 && !weeks.includes(w)) weeks.push(w);
  }
  let year = null;
  const y4 = s.match(/(?:^|\D)(20\d{2})(?!\d)/);
  const y2 = s.match(/(?:^|\D)(\d{2})\s*W\s*\d/);
  if (y4) year = parseInt(y4[1], 10);
  else if (y2) year = 2000 + parseInt(y2[1], 10);
  return { weeks: weeks.sort((a, b) => a - b), year };
}

// Area code from a folder or file name
function detectArea(text) {
  const t = str(text).toLowerCase();
  if (t.includes('alex'))                 return 'A';
  if (t.includes('delta'))                return 'D';
  if (/upper|cairo|giza/.test(t))         return 'U';
  return null;
}

function distanceFactor(raw) {
  const k = str(raw).toLowerCase().replace(/\s+/g, '');
  return DISTANCE_MULTIPLIERS[k] ?? 1.0;
}

// "500K", "1.5M", "500,000", "EGP 500 000" → number; blank → null; junk → NaN
function parseAmount(v) {
  const s = str(v).toUpperCase().replace(/EGP|[,\s]/g, '');
  if (!s) return null;
  const m = s.match(/^(\d+(?:\.\d+)?)([KM]?)$/);
  if (!m) return NaN;
  return Number(m[1]) * (m[2] === 'K' ? 1e3 : m[2] === 'M' ? 1e6 : 1);
}

function fmtEGP(n) {
  return 'EGP ' + Number(n || 0).toLocaleString('en-EG', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function fmtQty(n) {
  return (Math.round(n * 100) / 100).toLocaleString('en-EG');
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

// First column index whose header passes any test, trying the tests in order
function findCol(headers, tests) {
  for (const test of tests) {
    const i = headers.findIndex(test);
    if (i >= 0) return i;
  }
  return -1;
}

function escHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function showError(msg) {
  const d = el('tso-error');
  d.textContent = msg; d.style.display = 'block';
  d.scrollIntoView({ behavior: 'smooth', block: 'center' });
}
function clearError() {
  const d = el('tso-error');
  d.textContent = ''; d.style.display = 'none';
}
function showProgress(id, show) {
  const e = el(id); if (e) e.style.display = show ? 'block' : 'none';
}
function setLoadingText(txt) {
  const p = el('tso-loading').querySelector('p');
  if (p) p.textContent = txt;
}

function checkReady() {
  const amount = parseAmount(el('tso-amount').value);
  el('tso-btn-run').disabled = !(_trk && _tsr && _mails.length
    && el('tso-sub').value !== '' && !Number.isNaN(amount));
}

// ---------------------------------------------------------------------------
// Mail attachment extraction (.msg via msgreader, .eml via plain MIME parsing)
// ---------------------------------------------------------------------------
async function getMsgReader() {
  if (_MsgReader) return _MsgReader;
  try {
    const mod = await import('https://esm.sh/@kenjiuno/msgreader');
    _MsgReader = mod.default ?? mod.MsgReader ?? mod;
    return _MsgReader;
  } catch { return null; }
}

async function extractFromMsg(file) {
  const Cls = await getMsgReader();
  if (!Cls) throw new Error('Outlook mail reader could not be loaded. Check your internet connection.');
  const rdr  = new Cls(await file.arrayBuffer());
  const info = rdr.getFileData();
  for (const att of (info.attachments || [])) {
    const ad   = rdr.getAttachment(att);
    const name = ad.fileName || att.fileName || '';
    if (EXCEL_EXT.test(name)) return { fileName: name, data: ad.content };
  }
  return null;
}

async function extractFromEml(file) {
  const text = await file.text();
  const bm   = text.match(/boundary=["']?([^\s"'\r\n;]+)["']?/i);
  if (!bm) return null;
  const parts = text.split(new RegExp('--' + bm[1].replace(/[.*+?^${}()|[\]\\]/g, '\\$&')));
  for (const part of parts) {
    const nm = part.match(/filename\*?=["']?(?:UTF-8'')?([^"'\r\n;]+)/i);
    if (!nm) continue;
    const fileName = decodeURIComponent(nm[1].trim());
    if (!EXCEL_EXT.test(fileName)) continue;
    const body = part.match(/\r?\n\r?\n([\s\S]*)/);
    const b64  = body ? body[1].replace(/[\r\n\s]/g, '') : '';
    if (!b64) continue;
    const bin  = atob(b64);
    const data = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) data[i] = bin.charCodeAt(i);
    return { fileName, data };
  }
  return null;
}

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

  const h = (rows[hdr] || []).map(normHeader);
  const col = {
    id:        findCol(h, [t => t === 'id#' || t === 'id #']),
    owner:     findCol(h, [t => t.includes('task owner')]),
    vendor:    findCol(h, [t => t.includes('vendor')]),
    site:      findCol(h, [t => t.includes('logical site')]),
    option:    findCol(h, [t => t.includes('site option')]),
    facing:    findCol(h, [t => t.includes('facing')]),
    taskDate:  findCol(h, [t => t.includes('task date')]),
    item:      findCol(h, [t => t === 'line item' || t === 'line items']),
    absQty:    findCol(h, [t => t.includes('absolute') && (t.includes('qty') || t.includes('quant'))]),
    prq:       findCol(h, [t => t.includes('prq')]),
    cert:      findCol(h, [t => t.includes('certificate')]),
    accStatus: findCol(h, [t => t.includes('acceptance status')]),
    week:      findCol(h, [t => t.includes('acceptance week')]),
    sub:       findCol(h, [
                 t => t.includes('tsr') && t.includes('sub'),
                 t => !t.includes('date') && (t.includes('submission') ||
                      (t.startsWith('sub') && (t.includes('#') || t.includes('no') || t.includes('num'))))
               ]),
    total:     findCol(h, [t => t.includes('new total')]),
    status:    findCol(h, [t => t === 'status']),
    distance:  findCol(h, [t => t.includes('distance')])
  };

  const missing = [];
  if (col.site  < 0) missing.push('"Logical Site ID"');
  if (col.item  < 0) missing.push('"Line Item"');
  if (col.week  < 0) missing.push('"Acceptance Week"');
  if (col.sub   < 0) missing.push('"TSR Sub#"');
  if (col.total < 0) missing.push('"New Total"');
  if (missing.length) {
    const found = (rows[hdr] || []).map((c, i) => (c ? colLetter(i) + ':"' + str(c) + '"' : null))
      .filter(Boolean).join(' | ');
    throw new Error('Tracking file: missing column(s) ' + missing.join(', ') +
      '. Headers found in row ' + (hdr + r0 + 1) + ': [ ' + found + ' ]');
  }

  const tasks = [];
  for (let i = hdr + 1; i < rows.length; i++) {
    const row  = rows[i] || [];
    const site = siteKey(row[col.site]);
    const item = str(row[col.item]);
    if (!site || !item) continue;
    const status = col.status >= 0 ? str(row[col.status]).toLowerCase() : '';
    if (status === 'cancelled' || status === 'canceled') continue;

    const absRaw = col.absQty >= 0 ? row[col.absQty] : null;
    const absNum = Number(absRaw);
    const qty    = (absRaw == null || absRaw === '' || isNaN(absNum) ? 1 : absNum) *
                   distanceFactor(col.distance >= 0 ? row[col.distance] : '');
    const amount = Number(row[col.total]) || 0;

    const values = {};
    for (const [, key] of OUT_COLS) values[key] = col[key] >= 0 ? row[col[key]] : null;

    tasks.push({
      excelRow: i + r0 + 1,
      sub:      subKey(row[col.sub]),
      subRaw:   str(row[col.sub]),
      site, siteRaw: str(row[col.site]),
      item, key: itemKey(item),
      facing:   col.facing >= 0 ? siteKey(row[col.facing]) : '',
      weekRaw:  str(row[col.week]),
      weekSegs: parseTrackingWeek(row[col.week]),
      qty, amount, values
    });
  }
  return { tasks, col, headerRow: hdr + r0 + 1, sheet: name };
}

// TSR file — sheet "Request Form - VF" (same detection as TSR Sub Prep)
function readTsr(wb) {
  const ws = wb.Sheets['Request Form - VF'];
  if (!ws) {
    throw new Error('TSR file: sheet "Request Form - VF" not found. Available: ' +
      wb.SheetNames.join(', '));
  }
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  let hdr = -1;
  for (let i = 0; i < rows.length && hdr < 0; i++) {
    if ((rows[i] || []).some(c => c && c.toString().toLowerCase().includes('item description'))) hdr = i;
  }
  if (hdr < 0) throw new Error('TSR file: could not find the "Item Description" header row.');

  let cItem = -1, cPrice = -1, cRem = -1;
  (rows[hdr] || []).forEach((cell, c) => {
    const t = str(cell).toLowerCase();
    if (cItem  < 0 && t.includes('item description'))                  cItem  = c;
    if (cPrice < 0 && t.includes('unit price'))                        cPrice = c;
    if (cRem   < 0 && t.includes('remaining') && !t.includes('after')) cRem   = c;
  });
  if (cRem < 0) cRem = 50;

  const items = new Map(); // item description → remaining qty
  for (let i = hdr + 1; i < rows.length; i++) {
    const row  = rows[i] || [];
    const desc = str(row[cItem]);
    if (!desc) continue;
    items.set(desc, (items.get(desc) ?? 0) + (Number(row[cRem]) || 0));
  }
  return { items, col: { item: cItem, price: cPrice, remaining: cRem }, headerRow: hdr + 1 };
}

// Canonical TSR item for a tracking line item: exact name, else one contains the other
function tsrKey(lineItem) {
  if (_tsr.items.has(lineItem)) return lineItem;
  for (const key of _tsr.items.keys()) {
    if (key.includes(lineItem) || lineItem.includes(key)) return key;
  }
  return null;
}

// Mail attachment workbook → task rows in sheet order. The sheet holding a
// Site ID + Line Item header is used (the one with the most rows if several).
function readMailWorkbook(wb) {
  let best = null;
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    if (!ws || !ws['!ref']) continue;
    const { rows, r0 } = sheetRows(ws);
    for (let i = 0; i < Math.min(rows.length, 60); i++) {
      const h = (rows[i] || []).map(normHeader);
      const cSite = findCol(h, [
        t => t === 'site id' || t === 'siteid',
        t => t.includes('logical site'),
        t => t.includes('site id') && !t.includes('facing'),
        t => t === 'site' || t === 'site code'
      ]);
      const cItem = findCol(h, [
        t => t.includes('item description'),
        t => t === 'line item' || t === 'line items',
        t => t.includes('line item'),
        t => t.includes('activity description'),
        t => t === 'item' || t === 'description' || t === 'activity'
      ]);
      if (cSite < 0 || cItem < 0) continue;
      const cFacing = findCol(h, [t => t.includes('facing')]);

      const out = [];
      for (let j = i + 1; j < rows.length; j++) {
        const r    = rows[j] || [];
        const site = siteKey(r[cSite]);
        const item = str(r[cItem]);
        if (!site || !item) continue;
        out.push({
          pos: out.length, excelRow: j + r0 + 1,
          site, item, key: itemKey(item),
          facing: cFacing >= 0 ? siteKey(r[cFacing]) : ''
        });
      }
      if (!best || out.length > best.rows.length) {
        best = { sheet: name, headerRow: i + r0 + 1,
                 col: { site: cSite, item: cItem, facing: cFacing }, rows: out };
      }
      break;
    }
  }
  if (!best) throw new Error('No sheet with a Site ID + Line Item header in the attachment');
  return best;
}

async function parseMail(mail) {
  if (_mailCache.has(mail.path)) return _mailCache.get(mail.path);
  let res;
  try {
    let att;
    if (/\.msg$/i.test(mail.file.name))      att = await extractFromMsg(mail.file);
    else if (/\.eml$/i.test(mail.file.name)) att = await extractFromEml(mail.file);
    else att = { fileName: mail.file.name, data: new Uint8Array(await mail.file.arrayBuffer()) };
    if (!att) throw new Error('No Excel attachment found in the mail');
    const wb = XLSX.read(att.data, { type: 'array', cellDates: true });
    res = { ...readMailWorkbook(wb), attachment: att.fileName };
  } catch (err) {
    res = { error: err.message || String(err), rows: [] };
  }
  _mailCache.set(mail.path, res);
  return res;
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
      checkReady();
    }
  });
}

wireFilePicker('tso-btn-track', 'tso-input-track', 'tso-card-track',
  'tso-track-filename', 'tso-track-progress', wb => { _trk = null; _trk = readTracking(wb); populateSubs(); });
wireFilePicker('tso-btn-tsr', 'tso-input-tsr', 'tso-card-tsr',
  'tso-tsr-filename', 'tso-tsr-progress', wb => { _tsr = null; _tsr = readTsr(wb); });

el('tso-btn-folder').addEventListener('click', () => el('tso-input-folder').click());
el('tso-input-folder').addEventListener('change', e => loadFolderFiles(Array.from(e.target.files)));
// A folder dragged onto the card arrives from the shared drag-drop handler in index.html
el('tso-input-folder').addEventListener('folderdrop', e => loadFolderFiles(e.detail));

function loadFolderFiles(files) {
  if (!files.length) return;
  clearError();
  _mails = []; _ignored = []; _mailCache = new Map();

  for (const file of files) {
    if (!MAIL_EXT.test(file.name) || file.name.startsWith('~$')) continue;
    const path  = file.webkitRelativePath || file.name;
    const parts = path.split('/');
    const base  = file.name.replace(/\.[^.]+$/, '');
    // Area from the deepest folder that names one, else from the mail name
    let area = null;
    for (let i = parts.length - 2; i >= 0 && !area; i--) area = detectArea(parts[i]);
    if (!area) area = detectArea(base);
    const { weeks, year } = parseMailWeeks(base);
    if (!area || !weeks.length) {
      _ignored.push({ path, reason: !area ? 'area not recognised' : 'no week number after "W"' });
      continue;
    }
    _mails.push({ file, path, name: base, area, weeks, year });
  }

  // Global mail order: oldest week first, then area
  _mails.sort((a, b) =>
    (a.year ?? 0) - (b.year ?? 0) || a.weeks[0] - b.weeks[0] ||
    AREA_ORDER[a.area] - AREA_ORDER[b.area] || a.name.localeCompare(b.name));
  _mails.forEach((m, i) => { m.rank = i; });

  const root   = (files[0].webkitRelativePath || '').split('/')[0] || 'Folder';
  const counts = ['U', 'A', 'D'].map(a => AREA_LABEL[a] + ' ' + _mails.filter(m => m.area === a).length);
  el('tso-folder-filename').textContent = root + '  (' + counts.join(' · ') + ' mails)';
  el('tso-card-folder').classList.toggle('loaded', _mails.length > 0);
  if (!_mails.length) {
    showError('No mails found in the folder. Expected Upper-Cairo / Alex / Delta folders holding ' +
      'mails named like "… Status W34-2026" or "… 26W03-W04".');
  }
  checkReady();
}

// TSR Sub# filter: every distinct value in the tracking column, with task counts
function populateSubs() {
  const sel  = el('tso-sub');
  const prev = sel.value;
  const map  = new Map(); // key → { label, count }
  for (const t of _trk.tasks) {
    if (!t.sub) continue;
    if (!map.has(t.sub)) map.set(t.sub, { label: t.subRaw, count: 0 });
    map.get(t.sub).count++;
  }
  const keys = [...map.keys()].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  sel.innerHTML = '<option value="">' + (keys.length ? '-- Select TSR Sub# --' : 'No values in TSR Sub# column') +
    '</option>' + keys.map(k =>
      '<option value="' + escHtml(k) + '">' + escHtml(map.get(k).label) +
      ' (' + map.get(k).count + ' task' + (map.get(k).count !== 1 ? 's' : '') + ')</option>').join('');
  if (map.has(prev)) sel.value = prev;
}

el('tso-sub').addEventListener('change', checkReady);
el('tso-amount').addEventListener('input', () => {
  const v    = parseAmount(el('tso-amount').value);
  const hint = el('tso-amount-hint');
  if (v === null)          hint.textContent = 'Blank = take every task found in the mails.';
  else if (Number.isNaN(v)) hint.textContent = 'Not a valid amount — use e.g. 500000, 500,000, 500K or 1.5M.';
  else                     hint.textContent = '= ' + fmtEGP(v);
  hint.classList.toggle('tso-amount-bad', Number.isNaN(v));
  checkReady();
});
el('tso-amount').addEventListener('keydown', e => {
  if (e.key === 'Enter' && !el('tso-btn-run').disabled) el('tso-btn-run').click();
});

el('tso-btn-run').addEventListener('click', async () => {
  clearError();
  el('tso-results').style.display = 'none';
  el('tso-loading').style.display = 'flex';
  await new Promise(r => setTimeout(r, 30));
  try {
    _result = await runOrder();
    renderResults(_result);
    el('tso-results').style.display = 'block';
    el('tso-results').scrollIntoView({ behavior: 'smooth' });
  } catch (err) {
    showError(err.message || 'An unexpected error occurred.');
  } finally {
    el('tso-loading').style.display = 'none';
  }
});

el('tso-btn-export').addEventListener('click', () => {
  exportOrder().catch(err => showError('Export failed: ' + err.message));
});
el('tso-btn-folders').addEventListener('click', () => {
  exportFolders().catch(err => showError('Folder download failed: ' + err.message));
});

// ---------------------------------------------------------------------------
// Core logic
// ---------------------------------------------------------------------------
// Mails that can hold a tracking week segment, best candidate first:
// same area, same year (when both are known), sharing a week number.
// A mail covering exactly the same weeks beats a partial overlap; ties go to
// the most recently saved file (a later "RE:" mail carries the latest sheet).
function mailsFor(seg) {
  const same = m => m.weeks.length === seg.weeks.length && m.weeks.every(w => seg.weeks.includes(w));
  return _mails
    .filter(m => m.area === seg.prefix &&
      (seg.year == null || m.year == null || m.year === seg.year) &&
      m.weeks.some(w => seg.weeks.includes(w)))
    .sort((a, b) => (same(b) - same(a)) ||
      (b.file.lastModified - a.file.lastModified) || a.name.localeCompare(b.name));
}

async function runOrder() {
  const sub    = el('tso-sub').value;
  const target = parseAmount(el('tso-amount').value);
  const scope  = _trk.tasks.filter(t => t.sub === sub);
  if (!scope.length) throw new Error('No tasks found for this TSR Sub#.');

  // ── 1. Find each task in its acceptance mail ─────────────────────────────
  const used     = new Set();   // "mail path|row pos" already matched
  const found    = [];
  const notFound = [];

  for (const t of scope) {
    if (!t.weekSegs.length) {
      notFound.push({ ...t, reason: t.weekRaw ? 'Acceptance Week "' + t.weekRaw + '" could not be read'
                                              : 'No Acceptance Week' });
      continue;
    }
    const cands = [];
    for (const seg of t.weekSegs) {
      for (const m of mailsFor(seg)) if (!cands.includes(m)) cands.push(m);
    }
    if (!cands.length) {
      notFound.push({ ...t, reason: 'No mail found for ' + t.weekRaw });
      continue;
    }

    let hit = null, siteSeen = false;
    const errors = [];
    for (const m of cands) {
      setLoadingText('Reading mail: ' + m.name + '…');
      const parsed = await parseMail(m);
      if (parsed.error) { errors.push(m.name + ': ' + parsed.error); continue; }
      const free = parsed.rows.filter(r => r.site === t.site && !used.has(m.path + '|' + r.pos));
      if (free.length) siteSeen = true;
      const same = free.filter(r => r.key === t.key);
      const row  = same.find(r => t.facing && r.facing === t.facing) || same[0];
      if (row) { hit = { mail: m, row }; break; }
    }

    if (hit) {
      used.add(hit.mail.path + '|' + hit.row.pos);
      found.push({ ...t, mail: hit.mail, mailRow: hit.row });
    } else {
      const names = cands.map(m => '"' + m.name + '"').join(', ');
      let reason;
      if (errors.length === cands.length) reason = 'Mail could not be read — ' + errors.join('; ');
      else if (siteSeen) reason = 'Site is in the mail but not with this line item (' + names + ')';
      else               reason = 'Site + line item not in the mail (' + names + ')';
      notFound.push({ ...t, reason });
    }
  }

  // ── 2. Order: mail order (oldest week first), then row order inside the mail
  found.sort((a, b) => a.mail.rank - b.mail.rank || a.mailRow.pos - b.mailRow.pos);

  // ── 3. Pick tasks in that order up to the target amount ──────────────────
  // Every task must fit in the TSR remaining quantity of its line item.
  // A task that would overshoot is taken only if that lands closer to the
  // target than stopping before it; otherwise it is skipped and smaller tasks
  // further down the list are still tried.
  const avail = new Map(_tsr.items);
  const selected = [];
  const skipped  = [];
  let total = 0;

  for (const t of found) {
    if (target != null && total >= target) {
      skipped.push({ ...t, reason: 'Target amount reached' }); continue;
    }
    const k = tsrKey(t.item);
    if (k === null) { skipped.push({ ...t, reason: 'Line item not in the TSR' }); continue; }
    const left = avail.get(k);
    if (t.qty > left + 0.005) {
      skipped.push({ ...t, reason: 'Not enough TSR quantity — needs ' + fmtQty(t.qty) +
        ', ' + fmtQty(Math.max(left, 0)) + ' left' });
      continue;
    }
    if (target != null && total + t.amount > target &&
        total + t.amount - target >= target - total) {
      skipped.push({ ...t, reason: 'Would move the total further from the target' }); continue;
    }
    avail.set(k, left - t.qty);
    total += t.amount;
    selected.push({ ...t, cum: total });
  }

  // ── 4. Number the mails holding selected tasks 1, 2, 3… in mail order ────
  const folders = [...new Set(selected.map(t => t.mail))].sort((a, b) => a.rank - b.rank);
  const folderOf = new Map(folders.map((m, i) => [m, i + 1]));
  for (const t of selected) t.folder = folderOf.get(t.mail);

  return {
    sub, subLabel: scope[0].subRaw, target, scope, selected, skipped, notFound,
    total, folders, folderOf,
    scopeAmount: scope.reduce((s, t) => s + t.amount, 0),
    foundAmount: found.reduce((s, t) => s + t.amount, 0)
  };
}

// ---------------------------------------------------------------------------
// Render
// ---------------------------------------------------------------------------
function renderResults(r) {
  const c = _trk.col;
  el('tso-col-info').textContent = [
    'Tracking header row ' + _trk.headerRow,
    'Site ' + colLetter(c.site), 'Line Item ' + colLetter(c.item),
    'Acc. Week ' + colLetter(c.week), 'TSR Sub# ' + colLetter(c.sub),
    'New Total ' + colLetter(c.total),
    'TSR Remaining ' + colLetter(_tsr.col.remaining),
    'Mails: ' + _mails.length + (_ignored.length ? ' (' + _ignored.length + ' files ignored)' : '')
  ].join('  ·  ');

  const stat = (label, value, amount, cls) =>
    '<div class="poc-stat-box ' + (cls || '') + '">' +
      '<span class="poc-stat-label">' + label + '</span>' +
      '<span class="poc-stat-value">' + value + '</span>' +
      '<span class="poc-stat-amount">' + amount + '</span>' +
    '</div>';

  const diff = r.target != null ? r.total - r.target : null;
  let banner;
  if (!r.selected.length) {
    banner = '<div class="acc-banner acc-banner-fail">No task could be selected — see the reasons below.</div>';
  } else if (r.target == null) {
    banner = '<div class="acc-banner acc-banner-ok">All ' + r.selected.length +
      ' tasks found in the mails that fit the TSR are selected: ' + fmtEGP(r.total) + '.</div>';
  } else {
    banner = '<div class="acc-banner ' + (Math.abs(diff) <= r.target * 0.05 ? 'acc-banner-ok' : 'acc-banner-fail') + '">' +
      'Selected ' + fmtEGP(r.total) + ' against a target of ' + fmtEGP(r.target) +
      ' (' + (diff >= 0 ? '+' : '−') + fmtEGP(Math.abs(diff)) + ').</div>';
  }

  el('tso-summary').innerHTML =
    '<div class="poc-stats-row acc-stats-row">' +
      stat('TSR Sub# ' + escHtml(r.subLabel), r.scope.length, fmtEGP(r.scopeAmount)) +
      stat('Found in mails', r.scope.length - r.notFound.length, fmtEGP(r.foundAmount)) +
      stat('Selected', r.selected.length, fmtEGP(r.total), 'acc-stat-pass') +
      stat('Not found in mails', r.notFound.length,
        fmtEGP(r.notFound.reduce((s, t) => s + t.amount, 0)), r.notFound.length ? 'acc-stat-fail' : '') +
    '</div>' + banner;

  // Collapsible section: the lists run to hundreds of rows, so they start closed
  const section = (title, count, extra, table) =>
    '<details class="tso-details">' +
      '<summary><span class="tso-sum-title">' + title + '</span>' +
      '<span class="tso-sum-count">' + count + '</span>' +
      (extra ? '<span class="tso-sum-extra">' + extra + '</span>' : '') + '</summary>' +
      '<div class="table-wrapper">' + table + '</div>' +
    '</details>';

  // Selected tasks, in mail order
  el('tso-selected-wrap').innerHTML = r.selected.length
    ? section('Selected tasks — in mail order', r.selected.length, fmtEGP(r.total),
      '<table class="acc-table"><thead><tr>' +
        '<th>#</th><th>Folder</th><th>Site ID</th><th>Line Item</th><th>Acc. Week</th><th>Mail (row)</th>' +
        '<th>Tracking Row</th><th>Amount</th><th>Running Total</th>' +
      '</tr></thead><tbody>' +
      r.selected.map((t, i) =>
        '<tr class="acc-row-pass">' +
          '<td>' + (i + 1) + '</td>' +
          '<td class="tso-folder-col">' + t.folder + '</td>' +
          '<td class="acc-site">' + escHtml(t.siteRaw) + '</td>' +
          '<td class="acc-item-col">' + escHtml(t.item) + '</td>' +
          '<td>' + escHtml(t.weekRaw) + '</td>' +
          '<td class="acc-sub">' + escHtml(AREA_LABEL[t.mail.area] + ' W' + t.mail.weeks.join('-W')) +
            ' (row ' + t.mailRow.excelRow + ')</td>' +
          '<td>' + t.excelRow + '</td>' +
          '<td class="num">' + fmtEGP(t.amount) + '</td>' +
          '<td class="num">' + fmtEGP(t.cum) + '</td>' +
        '</tr>').join('') +
      '</tbody></table>')
    : '';

  // Tasks left out, with the reason
  const left = r.skipped.concat(r.notFound);
  el('tso-skipped-wrap').innerHTML = left.length
    ? section('Not included', left.length, fmtEGP(left.reduce((s, t) => s + t.amount, 0)),
      '<table class="acc-table"><thead><tr>' +
        '<th>Tracking Row</th><th>Site ID</th><th>Line Item</th><th>Acc. Week</th>' +
        '<th>Amount</th><th>Reason</th>' +
      '</tr></thead><tbody>' +
      left.map(t =>
        '<tr class="' + (t.mail ? '' : 'acc-row-fail') + '">' +
          '<td>' + t.excelRow + '</td>' +
          '<td class="acc-site">' + escHtml(t.siteRaw) + '</td>' +
          '<td class="acc-item-col">' + escHtml(t.item) + '</td>' +
          '<td>' + escHtml(t.weekRaw) + '</td>' +
          '<td class="num">' + fmtEGP(t.amount) + '</td>' +
          '<td class="acc-notes-col"><span class="acc-note">' + escHtml(t.reason) + '</span></td>' +
        '</tr>').join('') +
      '</tbody></table>')
    : '';

  // Mails that were opened, so a wrong pick or an unreadable sheet is visible
  const opened = _mails.filter(m => _mailCache.has(m.path))
    .sort((a, b) => (r.folderOf.get(a) ?? Infinity) - (r.folderOf.get(b) ?? Infinity) || a.rank - b.rank);
  el('tso-mails-wrap').innerHTML = opened.length || _ignored.length
    ? section('Mails used', r.folders.length + ' folder' + (r.folders.length !== 1 ? 's' : ''),
        opened.length + ' mail' + (opened.length !== 1 ? 's' : '') + ' opened',
      '<table class="acc-table"><thead><tr>' +
        '<th>Folder</th><th>Area</th><th>Weeks</th><th>Mail</th><th>Attachment / Sheet</th><th>Tasks taken</th>' +
      '</tr></thead><tbody>' +
      opened.map(m => {
        const p = _mailCache.get(m.path);
        const n = r.selected.filter(t => t.mail === m).length;
        return '<tr class="' + (p.error ? 'acc-row-fail' : '') + '">' +
          '<td class="tso-folder-col">' + (r.folderOf.get(m) ?? '—') + '</td>' +
          '<td>' + AREA_LABEL[m.area] + '</td>' +
          '<td>' + m.weeks.join(', ') + (m.year ? ' / ' + m.year : '') + '</td>' +
          '<td class="acc-item-col">' + escHtml(m.path) + '</td>' +
          '<td class="acc-sub">' + (p.error ? '<span class="acc-item-bad">' + escHtml(p.error) + '</span>'
            : escHtml(p.attachment) + ' — sheet "' + escHtml(p.sheet) + '", header row ' + p.headerRow +
              ', ' + p.rows.length + ' rows') + '</td>' +
          '<td>' + n + '</td>' +
        '</tr>';
      }).join('') +
      _ignored.map(f =>
        '<tr><td colspan="3" class="acc-sub">Ignored</td><td class="acc-item-col">' + escHtml(f.path) +
        '</td><td colspan="2" class="acc-sub">' + escHtml(f.reason) + '</td></tr>').join('') +
      '</tbody></table>')
    : '';

  el('tso-btn-export').disabled = !r.selected.length;
  el('tso-btn-folders').disabled = !r.folders.length;
  el('tso-export-status').textContent = '';
}

// ---------------------------------------------------------------------------
// Export
// ---------------------------------------------------------------------------
async function exportOrder() {
  const r = _result;
  if (!r || !r.selected.length) return;

  const THIN     = { style: 'thin', color: { argb: 'FF000000' } };
  const ALL_THIN = { top: THIN, left: THIN, bottom: THIN, right: THIN };

  function addSheet(wb, name, tasks, extraCols) {
    const ws   = wb.addWorksheet(name);
    const cols = OUT_COLS.concat(extraCols);
    cols.forEach(([, , w], i) => { ws.getColumn(i + 1).width = w; });

    const hdr = ws.addRow(cols.map(([h]) => h));
    hdr.height = 22;
    hdr.eachCell(c => {
      c.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0070C0' } };
      c.font      = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
      c.alignment = { horizontal: 'center', vertical: 'middle' };
      c.border    = ALL_THIN;
    });
    ws.views = [{ state: 'frozen', ySplit: 1 }];

    for (const t of tasks) {
      const row = ws.addRow(cols.map(([, key]) => {
        const v = key in t.values ? t.values[key] : t[key];
        return v == null ? '' : v;
      }));
      row.eachCell({ includeEmpty: true }, c => {
        let v = c.value;
        if (v instanceof Date) {
          // SheetJS dates are local midnight; shift so UTC midnight = local date
          c.value  = new Date(v.getTime() - v.getTimezoneOffset() * 60000);
          c.numFmt = 'dd-mmm-yy';
        }
        c.border    = ALL_THIN;
        c.font      = { size: 11 };
        c.alignment = { vertical: 'middle' };
      });
    }
  }

  const wb = new ExcelJS.Workbook();
  wb.creator = 'LMP Invoicing System';
  addSheet(wb, 'Submission', r.selected, [['Folder', 'folder', 10]]);
  const left = r.skipped.concat(r.notFound);
  if (left.length) {
    addSheet(wb, 'Not Included', left, [['Acceptance Week', 'weekRaw', 20], ['Reason', 'reason', 60]]);
  }

  const filename = fileBase(r) + '_Order_' + new Date().toISOString().slice(0, 10) + '.xlsx';
  downloadBlob(new Blob([await wb.xlsx.writeBuffer()],
    { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), filename);

  el('tso-export-status').textContent = '✅ Downloaded ' + filename + ' — ' +
    r.selected.length + ' task' + (r.selected.length !== 1 ? 's' : '') + ', ' + fmtEGP(r.total) + '.';
}

// Mail folders: every mail holding a selected task goes into its numbered
// folder (1, 2, 3… — the Folder column of the export), zipped for download.
async function exportFolders() {
  const r = _result;
  if (!r || !r.folders.length) return;
  const JSZip = await getJsZip();
  if (!JSZip) throw new Error('ZIP library could not be loaded. Check your internet connection.');

  el('tso-export-status').textContent = 'Preparing ' + r.folders.length + ' folders…';
  const zip = new JSZip();
  r.folders.forEach((m, i) => zip.folder(String(i + 1)).file(m.file.name, m.file));
  const blob = await zip.generateAsync({ type: 'blob' });

  const filename = fileBase(r) + '_Mails_' + new Date().toISOString().slice(0, 10) + '.zip';
  downloadBlob(blob, filename);
  el('tso-export-status').textContent = '✅ Downloaded ' + filename + ' — folders 1–' +
    r.folders.length + ', one mail each.';
}

let _JSZip = null;
async function getJsZip() {
  if (_JSZip) return _JSZip;
  try {
    await new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
      s.onload = res; s.onerror = rej;
      document.head.appendChild(s);
    });
    _JSZip = window.JSZip || null;
    return _JSZip;
  } catch { return null; }
}

function fileBase(r) {
  return 'TSR_Sub_' + (r.subLabel.replace(/[\/:*?"<>|#]+/g, '').trim().replace(/\s+/g, '_') || 'Sub');
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a   = Object.assign(document.createElement('a'), { href: url, download: filename });
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

})(); // end IIFE
