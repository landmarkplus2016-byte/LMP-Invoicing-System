// =============================================================================
// LMP Invoicing System — tsr-core.js
// Shared TSR rules: one place for the matching keys, distance bands, week
// parsing, mail reading and the readers for the tracking and TSR files.
// Exposed as window.TSRCore. Used by the TSR Planner tab (tsr-planner-app.js).
// =============================================================================
(function () {
'use strict';

// ---------------------------------------------------------------------------
// Text helpers
// ---------------------------------------------------------------------------
function str(v) { return v == null ? '' : String(v).trim(); }

// Header text normaliser: "Item_Description" → "item description"
function normHeader(v) { return str(v).toLowerCase().replace(/[_\s]+/g, ' ').trim(); }

// Placeholder text that means "no value": "", "NA", "N/A", "-"
function isBlank(v) { return /^(|n\/?a|-+)$/i.test(str(v)); }

function escHtml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function fmtEGP(n) {
  return 'EGP ' + Number(n || 0).toLocaleString('en-EG', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function fmtQty(n) { return (Math.round((n || 0) * 100) / 100).toLocaleString('en-EG'); }
function round4(n) { return Math.round((n || 0) * 10000) / 10000; }

function colLetter(idx) {
  let letter = '', n = idx + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}
function colIndex(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

// First column index whose header passes any test, trying the tests in order
function findCol(headers, tests) {
  for (const test of tests) {
    const i = headers.findIndex(test);
    if (i >= 0) return i;
  }
  return -1;
}

function findSheet(wb, name) {
  return wb.SheetNames.find(n => n.trim().toLowerCase() === name.toLowerCase()) || null;
}

function sheetRows(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: true });
  const r0   = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']).s.r : 0;
  return { rows, r0 };
}

// Value of one cell (0-based row / col) without converting the whole sheet —
// the TSR PO Break Down sheet carries a 100 000-row hidden site list.
function cellV(ws, r, c) {
  const cell = ws[XLSX.utils.encode_cell({ r, c })];
  return cell ? cell.v : null;
}

// ---------------------------------------------------------------------------
// Matching keys
// ---------------------------------------------------------------------------
// Site ID key: "U3844" → "U3844", 3688 → "3688", "0568" → "568", "D6840 " → "D6840"
function siteKey(v) {
  if (v == null || v === '') return '';
  if (typeof v === 'number') return String(Math.round(v));
  const s = String(v).toUpperCase().replace(/\s+/g, '').replace(/\.0+$/, '');
  return /^\d+$/.test(s) ? String(parseInt(s, 10)) : s;
}

// Codes renamed between the old and new price lists — the same task.
// Key = alternative code, value = the code it is treated as.
const ITEM_ALIASES = {
  DS07: 'DS09'   // "DS07 - Dismantle & Return Microwave Link or Dish" = "DS09 - Dismantle & Return Microwave Link"
};

// Items merged by description between the old and new price lists
const isMacro = t => /macro\s*\/\s*mega\s+site/.test(t);
const hasFive = t => /\bfive\b|\(5\)/.test(t);
const ITEM_GROUPS = [
  { key: 'RF-MACRO-5PLUS', test: t => isMacro(t) && hasFive(t) },
  { key: 'RF-MACRO-UPTO4', test: t => isMacro(t) && /\b(three|four|up to four)\b|\([34]\)/.test(t) }
];

// Raw catalogue code: "TX09 - IDU Upgrade" → "TX09", "EX6" → "EX06", "DS07_Dismantle…" → "DS07"
function itemCode(desc) {
  const m = str(desc).match(/^([A-Za-z]{1,4})\s*[-_.]?\s*(\d{1,3})(?!\d)/);
  return m ? m[1].toUpperCase() + String(parseInt(m[2], 10)).padStart(2, '0') : '';
}

// Matching key of a line item: merged items → group key, renamed codes → alias
function itemKey(desc) {
  const s = str(desc);
  const t = s.toLowerCase();
  const group = ITEM_GROUPS.find(g => g.test(t));
  if (group) return group.key;
  const code = itemCode(s);
  if (code) return ITEM_ALIASES[code] || code;
  return t.replace(/\s+/g, ' ');
}

// TSR Sub# cell → comparable key
function subKey(v) {
  if (v == null || v instanceof Date) return '';
  if (typeof v === 'number') return String(v);
  const s = String(v).trim().replace(/\s+/g, ' ');
  return /^\d+\.0+$/.test(s) ? String(parseInt(s, 10)) : s.toLowerCase();
}

// ---------------------------------------------------------------------------
// Status, distance, weeks
// ---------------------------------------------------------------------------
// Tracking Status → 'dead' (not invoiced at all), 'done' or 'open'
const DEAD_STATUS = /^(cancel+ed|cancal+ed|cancel|withdraw|withdrawn|repeated|duplicated?|closed|rej|rejected)$/i;
function statusClass(raw) {
  const s = str(raw).toLowerCase();
  if (DEAD_STATUS.test(s)) return 'dead';
  if (s === 'done') return 'done';
  return 'open';
}

// Distance band → quantity factor. Unknown / blank → factor null.
const DISTANCE_BANDS = [
  { key: '0km-100km',   factor: 1.0  },
  { key: '100km-400km', factor: 1.1  },
  { key: '400km-800km', factor: 1.2  },
  { key: '>800km',      factor: 1.25 }
];
function distanceFactor(raw) {
  const k = str(raw).toLowerCase().replace(/\s+/g, '');
  const b = DISTANCE_BANDS.find(x => x.key === k);
  return b ? b.factor : null;
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
    // "A-W43-2026" or the full word "Alex-W43-2026" / "Delta-…" / "Upper-…"
    const pre = p ? p[1] : '';
    const prefix = pre.length > 1 ? (detectArea(pre) || pre) : pre;
    return { prefix, weeks, year: y ? parseInt(y[1], 10) : null };
  }).filter(seg => seg.weeks.length > 0);
}

// Sort key of an acceptance week (year*100 + lowest week); Infinity when unreadable
function weekSortKey(raw) {
  let best = Infinity;
  for (const seg of parseTrackingWeek(raw)) {
    if (seg.year == null) continue;
    best = Math.min(best, seg.year * 100 + Math.min(...seg.weeks));
  }
  return best;
}

// Mail name → weeks + year
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

function detectArea(text) {
  const t = str(text).toLowerCase();
  if (t.includes('alex'))         return 'A';
  if (t.includes('delta'))        return 'D';
  if (/upper|cairo|giza/.test(t)) return 'U';
  return null;
}
const AREA_LABEL = { A: 'Alex', D: 'Delta', U: 'Upper-Cairo' };
const AREA_ORDER = { A: 0, D: 1, U: 2 };

// ---------------------------------------------------------------------------
// Mail files
// ---------------------------------------------------------------------------
const MAIL_EXT  = /\.(msg|eml|xlsx|xls|xlsm|xlsb)$/i;
const EXCEL_EXT = /\.(xlsx|xls|xlsm|xlsb)$/i;

// Folder upload → mail list (area from the deepest folder that names one,
// weeks from the mail name), sorted oldest week first, then area
function buildMailList(files) {
  const mails = [], ignored = [];
  for (const file of files) {
    if (!MAIL_EXT.test(file.name) || file.name.startsWith('~$')) continue;
    const path  = file.webkitRelativePath || file.name;
    const parts = path.split('/');
    const base  = file.name.replace(/\.[^.]+$/, '');
    let area = null;
    for (let i = parts.length - 2; i >= 0 && !area; i--) area = detectArea(parts[i]);
    if (!area) area = detectArea(base);
    const { weeks, year } = parseMailWeeks(base);
    if (!area || !weeks.length) {
      ignored.push({ path, reason: !area ? 'area not recognised' : 'no week number after "W"' });
      continue;
    }
    mails.push({ file, path, name: base, area, weeks, year });
  }
  mails.sort((a, b) =>
    (a.year ?? 0) - (b.year ?? 0) || a.weeks[0] - b.weeks[0] ||
    AREA_ORDER[a.area] - AREA_ORDER[b.area] || a.name.localeCompare(b.name));
  mails.forEach((m, i) => { m.rank = i; });
  return { mails, ignored };
}

// Mails that can hold a tracking week segment, best candidate first
function mailsFor(mails, seg) {
  const same = m => m.weeks.length === seg.weeks.length && m.weeks.every(w => seg.weeks.includes(w));
  return mails
    .filter(m => m.area === seg.prefix &&
      (seg.year == null || m.year == null || m.year === seg.year) &&
      m.weeks.some(w => seg.weeks.includes(w)))
    .sort((a, b) => (same(b) - same(a)) ||
      (b.file.lastModified - a.file.lastModified) || a.name.localeCompare(b.name));
}

let _MsgReader = null;
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

// Mail attachment workbook → rows in sheet order (sheet with Site ID + Line Item)
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
      const cReq    = findCol(h, [t => t.includes('request')]);
      const out = [];
      for (let j = i + 1; j < rows.length; j++) {
        const r    = rows[j] || [];
        const site = siteKey(r[cSite]);
        const item = str(r[cItem]);
        if (!site || !item) continue;
        out.push({
          pos: out.length, excelRow: j + r0 + 1,
          site, item, key: itemKey(item),
          facing: cFacing >= 0 ? siteKey(r[cFacing]) : '',
          facingRaw: cFacing >= 0 ? r[cFacing] : null,
          hasFacingCol: cFacing >= 0,
          request: cReq >= 0 ? r[cReq] : null
        });
      }
      if (!best || out.length > best.rows.length) {
        best = { sheet: name, headerRow: i + r0 + 1, rows: out };
      }
      break;
    }
  }
  if (!best) throw new Error('No sheet with a Site ID + Line Item header in the attachment');
  return best;
}

// One mail → parsed attachment rows (cached by path by the caller)
async function parseMailFile(file) {
  try {
    let att;
    if (/\.msg$/i.test(file.name))      att = await extractFromMsg(file);
    else if (/\.eml$/i.test(file.name)) att = await extractFromEml(file);
    else att = { fileName: file.name, data: new Uint8Array(await file.arrayBuffer()) };
    if (!att) throw new Error('No Excel attachment found in the mail');
    const wb = XLSX.read(att.data, { type: 'array', cellDates: true });
    return { ...readMailWorkbook(wb), attachment: att.fileName };
  } catch (err) {
    return { error: err.message || String(err), rows: [] };
  }
}

// Mail row is the reference: Certificate # must equal the mail Request #,
// Facing must equal the mail Facing (blank mail facing clears it).
// Returns { cert, facing, fixes }
function mailFixes(task, mr) {
  let cert = task.cert, facing = task.facingRaw;
  const fixes = [];
  if (mr && !isBlank(mr.request)) {
    const same = str(cert).replace(/\.0+$/, '').toLowerCase() === str(mr.request).replace(/\.0+$/, '').toLowerCase();
    if (!same) {
      fixes.push({ field: 'cert', text: isBlank(cert)
        ? 'Certificate # filled from mail: ' + str(mr.request)
        : 'Certificate # ' + str(cert) + ' → ' + str(mr.request) + ' (mail Request #)' });
      cert = mr.request;
    }
  }
  if (mr && mr.hasFacingCol && isBlank(mr.facingRaw)) {
    if (str(facing) !== '' && !isBlank(facing)) {
      fixes.push({ field: 'facing', text: 'Facing ' + str(facing) + ' removed (blank in mail)' });
    }
    facing = '';
  } else if (mr && !isBlank(mr.facingRaw)) {
    if (siteKey(isBlank(facing) ? '' : facing) !== mr.facing) {
      fixes.push({ field: 'facing', text: isBlank(facing)
        ? 'Facing filled from mail: ' + str(mr.facingRaw)
        : 'Facing ' + str(facing) + ' → ' + str(mr.facingRaw) + ' (mail)' });
      facing = mr.facingRaw;
    }
  }
  if (isBlank(facing)) facing = '';
  return { cert, facing, fixes };
}

// ---------------------------------------------------------------------------
// Tracking file — sheet "Invoicing Track", every column found by header name
// ---------------------------------------------------------------------------
function readTracking(wb) {
  const name = findSheet(wb, 'invoicing track');
  if (!name) throw new Error('Tracking file: sheet "Invoicing Track" not found. Available: ' + wb.SheetNames.join(', '));
  const { rows, r0 } = sheetRows(wb.Sheets[name]);

  let hdr = -1;
  for (let i = 0; i < Math.min(rows.length, 30); i++) {
    const h = (rows[i] || []).map(normHeader);
    if (h.some(t => t.includes('logical site')) && h.some(t => t.includes('acceptance week'))) { hdr = i; break; }
  }
  if (hdr < 0) throw new Error('Tracking file: header row not found (needs "Logical Site ID" and "Acceptance Week").');

  const h = (rows[hdr] || []).map(normHeader);
  const col = {
    id:        findCol(h, [t => t === 'id#' || t === 'id #']),
    jc:        findCol(h, [t => t.includes('job code')]),
    vendor:    findCol(h, [t => t === 'vendor', t => t.includes('vendor')]),
    site:      findCol(h, [t => t.includes('logical site')]),
    option:    findCol(h, [t => t.includes('site option')]),
    facing:    findCol(h, [t => t.includes('facing')]),
    distance:  findCol(h, [t => t.includes('distance')]),
    absQty:    findCol(h, [t => t.includes('absolute') && (t.includes('qty') || t.includes('quant'))]),
    item:      findCol(h, [t => t === 'line item' || t === 'line items']),
    price:     findCol(h, [t => t === 'new price']),
    total:     findCol(h, [t => t.includes('new total')]),
    comments:  findCol(h, [t => t === 'comments']),
    status:    findCol(h, [t => t === 'status']),
    taskDate:  findCol(h, [t => t.includes('task date')]),
    owner:     findCol(h, [t => t.includes('task owner')]),
    prq:       findCol(h, [t => t === 'prq', t => t.includes('prq')]),
    accStatus: findCol(h, [t => t.includes('acceptance status')]),
    fac:       findCol(h, [t => t.includes('fac date'), t => t === 'fac']),
    cert:      findCol(h, [t => t.includes('certificate')]),
    week:      findCol(h, [t => t.includes('acceptance week')]),
    sub:       findCol(h, [t => t.includes('tsr') && t.includes('sub')]),
    poStatus:  findCol(h, [t => t === 'po status']),
    poNumber:  findCol(h, [t => t === 'po number'])
  };
  const required = { jc: 'Job Code', site: 'Logical Site ID', item: 'Line Item', status: 'Status',
                     week: 'Acceptance Week', total: 'New Total', absQty: 'Absolute Quantity' };
  const missing = Object.entries(required).filter(([k]) => col[k] < 0).map(([, v]) => '"' + v + '"');
  if (missing.length) throw new Error('Tracking file: missing column(s) ' + missing.join(', ') + '.');

  const get = (row, k) => (col[k] >= 0 ? row[col[k]] : null);
  const tasks = [];
  for (let i = hdr + 1; i < rows.length; i++) {
    const row  = rows[i] || [];
    const site = siteKey(get(row, 'site'));
    const item = str(get(row, 'item'));
    if (!site || !item) continue;
    const absRaw = get(row, 'absQty');
    const absNum = Number(absRaw);
    const abs    = absRaw == null || absRaw === '' || isNaN(absNum) ? 1 : absNum;
    const factor = distanceFactor(get(row, 'distance'));
    const sc     = statusClass(get(row, 'status'));
    tasks.push({
      excelRow: i + r0 + 1,
      id: get(row, 'id'), jc: str(get(row, 'jc')), jcKey: str(get(row, 'jc')).toUpperCase(),
      vendor: str(get(row, 'vendor')), owner: str(get(row, 'owner')),
      siteRaw: get(row, 'site'), site, option: get(row, 'option'),
      facingRaw: get(row, 'facing'), facing: siteKey(isBlank(get(row, 'facing')) ? '' : get(row, 'facing')),
      distanceRaw: str(get(row, 'distance')), factor,
      abs, qty: abs * (factor ?? 1),
      item, key: itemKey(item), code: itemCode(item),
      total: Number(get(row, 'total')) || 0, price: Number(get(row, 'price')) || 0,
      comments: str(get(row, 'comments')),
      statusRaw: str(get(row, 'status')), status: sc,
      taskDate: get(row, 'taskDate'), prq: get(row, 'prq'),
      accStatus: str(get(row, 'accStatus')), fac: get(row, 'fac'),
      cert: get(row, 'cert'),
      weekRaw: str(get(row, 'week')), weekSegs: parseTrackingWeek(get(row, 'week')),
      weekKey: weekSortKey(get(row, 'week')),
      subRaw: str(get(row, 'sub')), sub: subKey(get(row, 'sub')),
      poStatus: str(get(row, 'poStatus')), poNumber: str(get(row, 'poNumber'))
    });
  }
  return { tasks, col, headerRow: hdr + r0 + 1, sheet: name };
}

// ---------------------------------------------------------------------------
// TSR file (.xlsb) — Request Form, PO Break Down and Receiving Template
// ---------------------------------------------------------------------------
const PB_SHEET = 'PO Break Down- Contractor & Acc';

function readTsrFile(wb, fileName) {
  const out = { warnings: [] };

  // ── Request Form - VF: items, qty, price, remaining, Req Form # ──────────
  const rfName = findSheet(wb, 'Request Form - VF');
  if (!rfName) throw new Error('TSR file: sheet "Request Form - VF" not found. Available: ' + wb.SheetNames.join(', '));
  const rfRows = XLSX.utils.sheet_to_json(wb.Sheets[rfName], { header: 1, defval: null, raw: true });
  let hdr = -1;
  for (let i = 0; i < rfRows.length && hdr < 0; i++) {
    if ((rfRows[i] || []).some(c => str(c).toLowerCase().includes('item description'))) hdr = i;
  }
  if (hdr < 0) throw new Error('TSR file: "Item Description" header not found on Request Form - VF.');
  let cItem = -1, cQty = -1, cPrice = -1, cRem = -1, cUnit = -1, cMat = -1, cEvo = -1, cPa = -1, cPaItem = -1;
  (rfRows[hdr] || []).forEach((cell, c) => {
    const t = str(cell).toLowerCase();
    if (cItem  < 0 && t.includes('item description'))                  cItem  = c;
    if (cQty   < 0 && t === 'qty')                                     cQty   = c;
    if (cPrice < 0 && t.includes('unit price'))                        cPrice = c;
    if (cRem   < 0 && t.includes('remaining') && !t.includes('after')) cRem   = c;
    if (cUnit  < 0 && t === 'unit')                                    cUnit  = c;
    if (cMat   < 0 && t.includes('material group'))                    cMat   = c;
    if (cEvo   < 0 && t.includes('evo'))                               cEvo   = c;
    if (cPa    < 0 && t.includes('price agreement'))                   cPa    = c;
    if (cPaItem < 0 && t.includes('pa item'))                          cPaItem = c;
  });
  out.rfCol = { item: cItem, qty: cQty, price: cPrice, remaining: cRem, unit: cUnit };
  out.rfHeaderRow = hdr + 1;
  if (cRem < 0) { cRem = 50; out.warnings.push('Request Form: "Remaining" header not found — column AY used.'); }
  out.items = [];
  for (let i = hdr + 1; i < rfRows.length; i++) {
    const row = rfRows[i] || [];
    const raw = row[cItem];
    if (!str(raw) || !itemCode(raw)) continue;
    out.items.push({
      name: String(raw),            // exact text — the PO Break Down price lookup needs it verbatim
      key: itemKey(raw), code: itemCode(raw),
      qty: Number(row[cQty]) || 0,
      price: Number(row[cPrice]) || 0,
      remaining: Number(row[cRem]) || 0,
      rfRow: i + 1,                 // Excel row on Request Form - VF
      unit: cUnit >= 0 ? str(row[cUnit]) : '',
      matGroup: cMat >= 0 ? str(row[cMat]) : '',
      evo: cEvo >= 0 ? str(row[cEvo]).replace(/\.0+$/, '') : '',
      pa: cPa >= 0 ? str(row[cPa]).replace(/\.0+$/, '') : '',
      paItem: cPaItem >= 0 ? str(row[cPaItem]).replace(/\.0+$/, '') : ''
    });
  }
  // Req Form # — cell right of a "Req Form#" label
  out.tsrNo = '';
  for (let i = 0; i < Math.min(rfRows.length, 20) && !out.tsrNo; i++) {
    const row = rfRows[i] || [];
    row.forEach((c, j) => { if (!out.tsrNo && /req\s*form/i.test(str(c))) out.tsrNo = str(row[j + 1]).replace(/\.0+$/, ''); });
  }
  if (!out.tsrNo) { const m = str(fileName).match(/(\d{3,5})/); if (m) out.tsrNo = m[1]; }

  // ── PO Break Down: header, existing rows, hidden lists ───────────────────
  const pbName = findSheet(wb, PB_SHEET);
  if (!pbName) throw new Error('TSR file: sheet "' + PB_SHEET + '" not found.');
  const ws  = wb.Sheets[pbName];
  const rng = XLSX.utils.decode_range(ws['!ref']);
  let pbHdr = -1;
  for (let r = 0; r <= Math.min(rng.e.r, 40) && pbHdr < 0; r++) {
    let hasItem = false, hasSite = false;
    for (let c = 0; c <= 30; c++) {
      const t = normHeader(cellV(ws, r, c));
      if (t.includes('item description')) hasItem = true;
      if (t === 'site id') hasSite = true;
    }
    if (hasItem && hasSite) pbHdr = r;
  }
  if (pbHdr < 0) throw new Error('TSR file: PO Break Down header row (Site ID + Item Description) not found.');
  const hdrs = [], hdrsRaw = [];
  for (let c = 0; c <= 30; c++) { hdrsRaw.push(str(cellV(ws, pbHdr, c))); hdrs.push(normHeader(cellV(ws, pbHdr, c))); }
  const pc = {
    owner:   findCol(hdrs, [t => t.includes('task owner')]),
    vendor:  findCol(hdrs, [t => t.includes('hardware vendor'), t => t.includes('vendor')]),
    site:    findCol(hdrs, [t => t === 'site id']),
    option:  findCol(hdrs, [t => t === 'option']),
    facing:  findCol(hdrs, [t => t.includes('facing')]),
    distance:findCol(hdrs, [t => t === 'distance']),
    date:    findCol(hdrs, [t => t.includes('activity date')]),
    item:    findCol(hdrs, [t => t.includes('item description')]),
    qty:     findCol(hdrs, [t => t === 'qty']),
    stream:  findCol(hdrs, [t => t.includes('general stream')]),
    prq:     findCol(hdrs, [t => t.includes('crq') || t.includes('prq')]),
    cert:    findCol(hdrs, [t => t.includes('certificate')]),
    status:  findCol(hdrs, [t => t.includes('status')]),
    sub:     findCol(hdrs, [t => t.includes('submission') && !t.includes('date')]),
    subDate: findCol(hdrs, [t => t.includes('submission date')]),
    comment: findCol(hdrs, [t => t.includes('comment')])
  };
  const needPc = ['site', 'item', 'qty', 'sub', 'stream', 'cert', 'status', 'comment', 'date', 'owner', 'vendor', 'option', 'facing', 'prq'];
  const missPc = needPc.filter(k => pc[k] < 0);
  if (missPc.length) throw new Error('TSR file: PO Break Down column(s) not found: ' + missPc.join(', '));
  out.pbHeaders = hdrsRaw;
  out.pbCol = pc;
  out.pbHeaderRow = pbHdr + 1;

  // Existing submission rows (data starts 2 rows below the header in VF's template)
  out.existing = [];
  out.lastRow = pbHdr + 1;           // Excel row of the last used Site ID / item
  out.submissions = new Set();
  for (let r = pbHdr + 1; r <= Math.min(rng.e.r, pbHdr + 6000); r++) {
    const item = str(cellV(ws, r, pc.item));
    const site = str(cellV(ws, r, pc.site));
    if (!item && !site) continue;
    out.lastRow = r + 1;
    if (!item) continue;
    const subTxt = str(cellV(ws, r, pc.sub));
    const m = subTxt.match(/(\d+)/);
    const subNo = m ? parseInt(m[1], 10) : null;
    if (subNo) out.submissions.add(subNo);
    const dCell = ws[XLSX.utils.encode_cell({ r, c: pc.distance })];
    out.existing.push({
      excelRow: r + 1, subNo,
      site: siteKey(cellV(ws, r, pc.site)), key: itemKey(item), code: itemCode(item), itemName: item,
      cert: str(cellV(ws, r, pc.cert)).replace(/\.0+$/, ''),
      facing: siteKey(cellV(ws, r, pc.facing)),
      qty: Number(cellV(ws, r, pc.qty)) || 0,
      factor: dCell && dCell.t !== 'e' ? distanceFactor(dCell.v) : null
    });
  }
  out.firstDataRow = pbHdr + 3;      // VF template: header row 8, data from row 10
  out.nextRow = Math.max(out.lastRow + 1, out.firstDataRow);
  out.nextSub = out.submissions.size ? Math.max(...out.submissions) + 1 : 1;

  // Hidden site list: header "Site Id | Physical | Region | Governorate | Site Type | Distance"
  out.siteList = new Map();
  let slCol = -1, slHdr = -1;
  for (let r = 0; r <= 30 && slCol < 0; r++) {
    for (let c = 30; c <= Math.min(rng.e.c, 200); c++) {
      if (normHeader(cellV(ws, r, c)) === 'site id' && normHeader(cellV(ws, r, c + 1)) === 'physical') { slCol = c; slHdr = r; break; }
    }
  }
  if (slCol >= 0) {
    for (let r = slHdr + 1; r <= rng.e.r; r++) {
      const id = cellV(ws, r, slCol);
      if (id == null || id === '') continue;
      const k = siteListKey(id);
      if (!out.siteList.has(k)) {
        out.siteList.set(k, {
          physical: cellV(ws, r, slCol + 1), region: cellV(ws, r, slCol + 2),
          governorate: cellV(ws, r, slCol + 3), distance: str(cellV(ws, r, slCol + 5))
        });
      }
    }
  } else {
    out.warnings.push('Hidden site list (Site Id / Physical / … / Distance) not found on PO Break Down — Distance preview unavailable.');
  }

  // Dropdown lists used by VF's template (fixed addresses, checked before use)
  const listAt = (letters, r1, r2) => {
    const c = colIndex(letters), vals = [];
    for (let r = r1 - 1; r <= r2 - 1; r++) { const v = str(cellV(ws, r, c)); if (v) vals.push(v); }
    return vals;
  };
  out.vendors = listAt('CO', 58, 61);
  if (!out.vendors.some(v => /ericsson|huawei/i.test(v))) { out.vendors = []; out.warnings.push('Vendor dropdown list not found (CO58:CO61).'); }
  out.streams = listAt('CO', 63, 155);
  if (out.streams.length < 10) { out.streams = []; out.warnings.push('General Stream dropdown list not found (CO63:CO155).'); }
  out.catalogue = new Map();          // exact item name → { unit, price } (price lookup CQ:CS)
  const cq = colIndex('CQ');
  for (let r = 0; r <= Math.min(rng.e.r, 400); r++) {
    const n = cellV(ws, r, cq);
    if (typeof n === 'string' && itemCode(n)) out.catalogue.set(n, { unit: cellV(ws, r, cq + 1), price: Number(cellV(ws, r, cq + 2)) || 0 });
  }

  // ── Receiving: which submissions VF has received (Remaining already net of them)
  out.receivedSubs = new Set();
  const rcName = findSheet(wb, 'Receiving');
  if (rcName) {
    const rc = XLSX.utils.sheet_to_json(wb.Sheets[rcName], { header: 1, defval: null, raw: true });
    for (let i = 0; i < Math.min(rc.length, 10); i++) {
      const row = rc[i] || [];
      const subs = row.map((c, j) => [j, str(c).match(/^submission\s*#\s*(\d+)/i)]).filter(x => x[1]);
      if (!subs.length) continue;
      const totals = rc[i + 1] || [];
      for (const [j, m] of subs) if ((Number(totals[j]) || 0) > 0) out.receivedSubs.add(parseInt(m[1], 10));
      break;
    }
  } else {
    out.warnings.push('Receiving sheet not found — every submission in the file is treated as not yet received.');
  }

  // ── Receiving Template: VF catalogue for the receiving file ──────────────
  out.receiving = [];
  const rtName = findSheet(wb, 'Receiving Template');
  if (rtName) {
    const rt = XLSX.utils.sheet_to_json(wb.Sheets[rtName], { header: 1, defval: null, raw: true });
    let rh = -1;
    for (let i = 0; i < Math.min(rt.length, 20) && rh < 0; i++) {
      const hh = (rt[i] || []).map(normHeader);
      if (hh.includes('description') && hh.some(t => t.includes('evo'))) rh = i;
    }
    if (rh >= 0) {
      const hh = (rt[rh] || []).map(normHeader);
      const rc = {
        vendor: hh.indexOf('vendor name'), vendorId: hh.indexOf('vendor id'),
        desc: hh.indexOf('description'), cost: hh.indexOf('unit cost'),
        line: hh.indexOf('line row'), evo: hh.findIndex(t => t.includes('evo')),
        stream: hh.indexOf('stream name')
      };
      for (let i = rh + 1; i < rt.length; i++) {
        const row = rt[i] || [];
        const desc = str(row[rc.desc]);
        if (!desc || !itemCode(desc)) continue;
        out.receiving.push({
          vendor: str(row[rc.vendor]), vendorId: row[rc.vendorId],
          desc, code: itemCode(desc), cost: Number(row[rc.cost]) || 0,
          line: str(row[rc.line]), evo: str(row[rc.evo]).replace(/\.0+$/, ''), stream: str(row[rc.stream])
        });
      }
    }
  }
  if (!out.receiving.length) out.warnings.push('Receiving Template sheet not found or empty — receiving file unavailable.');
  return out;
}

// Key for the hidden site list, matching how VLOOKUP(…, FALSE) compares:
// case-insensitive but space-sensitive ("D6840 " ≠ "D6840"). The planner writes
// trimmed Site IDs, so a lookup with the trimmed value predicts the TSR result.
function siteListKey(v) {
  if (v == null) return '';
  if (typeof v === 'number') return 'n:' + String(Math.round(v * 1000) / 1000);   // numbers only match numbers
  return 't:' + String(v).toUpperCase();
}

// Request Form item for a tracking line item: same code first, then same key
function tsrItemFor(tsr, task) {
  return tsr.items.find(it => it.code === task.code) ||
         tsr.items.find(it => it.key === task.key) || null;
}

// Site ID as written into the TSR. Numeric IDs must be numbers: VLOOKUP never
// matches text "3618" to the number 3618 in VF's site list.
function tsrSite(t) {
  const s = str(t.siteRaw);
  return typeof t.siteRaw === 'number' ? t.siteRaw : (/^\d+$/.test(s) ? Number(s) : s);
}

// Test "is this task tagged with TSR <tag>" on the tracking TSR Sub# column:
// "1326", "1326 sub-2", "check 1326" → true; "11326" → false.
function tagTest(tag) {
  const t = str(tag).toLowerCase();
  if (!t) return () => false;
  const re = new RegExp('(^|\\D)' + t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(\\D|$)');
  return task => re.test(task.sub);
}

// Rows already in the TSR file (PO Break Down) → tracking tasks.
// Site + item key; Certificate # preferred, then a task tagged with this TSR,
// then Done, same Facing, with a FAC date.
// Returns { inTsr: Map(task → submission #), pairs: [{ ex, task }], unmatched: [ex] }
function matchExisting(tsr, tasks, tag) {
  const inTsr = new Map(), pairs = [], unmatched = [];
  const bySiteItem = new Map();
  for (const t of tasks) {
    const k = t.site + '|' + t.key;
    if (!bySiteItem.has(k)) bySiteItem.set(k, []);
    bySiteItem.get(k).push(t);
  }
  const certOf = v => str(v).replace(/\.0+$/, '');
  const tg = str(tag).toLowerCase();
  for (const ex of tsr.existing) {
    const cands = (bySiteItem.get(ex.site + '|' + ex.key) || []).filter(t => !inTsr.has(t));
    if (!cands.length) { unmatched.push(ex); continue; }
    // Certificate #, then this TSR's tag, then a Done task, same Facing, with a FAC date
    // (an old Assigned duplicate of the same Site + item must not take the row)
    const score = t => [certOf(t.cert) === ex.cert, !!tg && t.sub.startsWith(tg), t.status === 'done',
                        !!ex.facing && t.facing === ex.facing, !!t.fac];
    cands.sort((a, b) => { const sa = score(a), sb = score(b);
      for (let i = 0; i < sa.length; i++) if (sa[i] !== sb[i]) return sb[i] - sa[i];
      return 0; });
    inTsr.set(cands[0], ex.subNo);
    pairs.push({ ex, task: cands[0] });
  }
  return { inTsr, pairs, unmatched };
}

// Quantity still free on the Request Form. Remaining only drops when VF
// receives a submission, so rows of submissions in the file that are not
// received yet are deducted (Qty × the Distance the TSR shows, else the
// matched task's tracking distance, else 1).
// Returns { avail: Map(item name → qty), pendingUse: Map, pendingSubs: [n] }
function tsrAvailable(tsr, pairs) {
  const taskOf = new Map(pairs.map(p => [p.ex, p.task]));
  const pendingUse = new Map(), pendingSubs = new Set();
  for (const ex of tsr.existing) {
    if (!ex.subNo || tsr.receivedSubs.has(ex.subNo)) continue;
    const item = tsr.items.find(it => it.code === ex.code) || tsr.items.find(it => it.key === ex.key);
    if (!item) continue;
    const task = taskOf.get(ex);
    const f = ex.factor ?? (task && task.factor) ?? 1;
    pendingUse.set(item.name, (pendingUse.get(item.name) || 0) + ex.qty * f);
    pendingSubs.add(ex.subNo);
  }
  const avail = new Map(tsr.items.map(it => [it.name, it.remaining - (pendingUse.get(it.name) || 0)]));
  return { avail, pendingUse, pendingSubs: [...pendingSubs].sort((a, b) => a - b) };
}

// ---------------------------------------------------------------------------
// Downloads
// ---------------------------------------------------------------------------
let _JSZip = null;
async function getJsZip() {
  if (_JSZip) return _JSZip;
  if (window.JSZip) return (_JSZip = window.JSZip);
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

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a   = Object.assign(document.createElement('a'), { href: url, download: filename });
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// SheetJS dates are local midnight — sometimes a few seconds before it (its
// serial → Date conversion drifts) — and ExcelJS writes UTC. Round to the
// nearest minute, take that calendar day, and write it as UTC midnight so
// Excel shows the same date as the tracking (16-Aug stays 16-Aug).
function excelDate(v) {
  if (!(v instanceof Date) || isNaN(v)) return v;
  const r = new Date(Math.round(v.getTime() / 60000) * 60000);
  return new Date(Date.UTC(r.getFullYear(), r.getMonth(), r.getDate()));
}

window.TSRCore = {
  str, normHeader, isBlank, escHtml, fmtEGP, fmtQty, round4, colLetter, colIndex, findCol, findSheet,
  sheetRows, cellV, siteKey, itemCode, itemKey, subKey, ITEM_ALIASES, ITEM_GROUPS,
  statusClass, distanceFactor, DISTANCE_BANDS, parseTrackingWeek, weekSortKey, parseMailWeeks,
  detectArea, AREA_LABEL, AREA_ORDER, buildMailList, mailsFor, parseMailFile, mailFixes,
  readTracking, readTsrFile, tsrItemFor, siteListKey, getJsZip, downloadBlob, excelDate,
  tsrSite, tagTest, matchExisting, tsrAvailable
};
})();
