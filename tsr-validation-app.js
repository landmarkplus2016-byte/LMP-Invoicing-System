// =============================================================================
// LMP Invoicing System — TSR Sub Validation logic
// Validates a TSR submission against Excel (regular folders) and PDF (TOC folders)
// attachments extracted from mail files.
// Wrapped in IIFE to avoid global namespace collisions.
// =============================================================================
(function () {

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------
let _tsrWorkbook  = null;
let _folderFiles  = [];
let _MsgReader    = null;   // cached @kenjiuno/msgreader
let _pdfJs        = null;   // cached PDF.js (pdfjsLib global)

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
const el = id => document.getElementById(id);

function normalizeStr(v) {
  if (v == null) return '';
  return String(v).trim();
}
function normalizeLower(v) { return normalizeStr(v).toLowerCase(); }

function colLetterToIndex(letter) {
  const s = letter.trim().toUpperCase();
  if (!/^[A-Z]+$/.test(s)) return -1;
  let idx = 0;
  for (let i = 0; i < s.length; i++) idx = idx * 26 + (s.charCodeAt(i) - 64);
  return idx - 1;
}
function colIndexToLetter(idx) {
  let letter = '', n = idx + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}

function extractNumber(v) {
  if (v == null || v === '' || v instanceof Date) return null;
  if (typeof v === 'number') return Math.round(v);
  const m = String(v).match(/\d+/);
  return m ? parseInt(m[0], 10) : null;
}

function normalizeFolderNumber(v) {
  if (v == null || v === '' || v instanceof Date) return null;
  if (typeof v === 'number') return Math.round(v);
  const s = String(v).trim();
  if (/^TOC/i.test(s)) return null;     // TOC folders handled by normalizeTocKey
  const m = s.match(/(\d+)/);           // first number anywhere in the string
  return m ? parseInt(m[1], 10) : null;
}

// Normalise a TOC folder identifier: "TOC 1", "TOC1", "toc-2", "TOC 2" → "TOC 1", "TOC 2"
function normalizeTocKey(s) {
  if (!s) return null;
  const str = String(s).trim();
  const m   = str.match(/TOC\s*[-\s]?\s*(\d+)/i);
  if (m) return 'TOC ' + parseInt(m[1], 10);
  if (/^TOC$/i.test(str)) return 'TOC 1';
  return null;
}

// Extract catalogue-code prefix for item matching: "EX06 - ..." → "EX06"
function itemMatchKey(desc) {
  const s = normalizeStr(desc);
  const m = s.match(/^([A-Za-z]{1,4}\d{1,3})\b/);
  return m ? m[1].toUpperCase() : normalizeLower(s);
}

// Normalise an activity code so all punctuation variants map to the same key:
// "EX.01" → "EX01",  "EX-01" → "EX01",  "EX 01" → "EX01",  "EX01" → "EX01"
function normalizeActivityCode(code) {
  return String(code || '').replace(/[\s.\-_]/g, '').toUpperCase();
}

function escHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function showError(msg) {
  const d = el('tsrval-error');
  d.textContent = msg; d.style.display = 'block';
}
function clearError() {
  const d = el('tsrval-error');
  d.textContent = ''; d.style.display = 'none';
}
function showProgress(id, show) {
  const e = el(id); if (e) e.style.display = show ? 'block' : 'none';
}
function checkReady() {
  const ready = _tsrWorkbook !== null
    && _folderFiles.length > 0
    && el('tsrval-sub-number').value.trim() !== '';
  el('tsrval-btn-analyze').disabled = !ready;
}
function setLoadingText(txt) {
  const p = el('tsrval-loading').querySelector('p');
  if (p) p.textContent = txt;
}

// ---------------------------------------------------------------------------
// CDN Library loaders (on-demand)
// ---------------------------------------------------------------------------
async function getMsgReader() {
  if (_MsgReader) return _MsgReader;
  try {
    const mod = await import('https://esm.sh/@kenjiuno/msgreader');
    _MsgReader = mod.default ?? mod.MsgReader ?? mod;
    return _MsgReader;
  } catch { return null; }
}

async function getPdfJs() {
  if (_pdfJs) return _pdfJs;
  try {
    await new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
      s.onload = res; s.onerror = rej;
      document.head.appendChild(s);
    });
    const lib = window.pdfjsLib;
    if (!lib) return null;
    // Point worker to same CDN (has CORS headers)
    lib.GlobalWorkerOptions.workerSrc =
      'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    _pdfJs = lib;
    return _pdfJs;
  } catch { return null; }
}

// ---------------------------------------------------------------------------
// Attachment extractors — MSG / EML (Excel or PDF variants)
// ---------------------------------------------------------------------------
async function extractAttachmentFromMsg(file, extRegex) {
  const Cls = await getMsgReader();
  if (!Cls) return null;
  try {
    const buf  = await file.arrayBuffer();
    const rdr  = new Cls(buf);
    const info = rdr.getFileData();
    for (const att of (info.attachments || [])) {
      const ad   = rdr.getAttachment(att);
      const name = (ad.fileName || att.fileName || '').toLowerCase();
      if (extRegex.test(name))
        return { fileName: ad.fileName || att.fileName || 'attachment', data: ad.content };
    }
  } catch (e) { console.warn('MSG extract error:', e); }
  return null;
}

async function extractAttachmentFromEml(file, extRegex) {
  try {
    const text = await file.text();
    const bm   = text.match(/boundary=["']?([^\s"'\r\n;]+)["']?/i);
    if (!bm) return null;
    const boundary = bm[1];
    const parts = text.split(
      new RegExp('--' + boundary.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))
    );
    for (const part of parts) {
      if (!extRegex.test(part)) continue;
      const nm = part.match(/filename\*?=["']?(?:UTF-8'')?([^"'\r\n;]+)/i);
      if (!nm) continue;
      const fileName = decodeURIComponent(nm[1].trim());
      const b64m = part.match(/\r?\n\r?\n([\s\S]*)/);
      if (!b64m) continue;
      const b64 = b64m[1].replace(/[\r\n\s]/g, '');
      if (!b64) continue;
      const bin = atob(b64);
      const data = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) data[i] = bin.charCodeAt(i);
      return { fileName, data };
    }
  } catch (e) { console.warn('EML extract error:', e); }
  return null;
}

const EXCEL_EXT = /\.(xlsx|xls|xlsb)$/i;
const PDF_EXT   = /\.pdf$/i;

async function extractFromMsg(file, extRegex) {
  return extractAttachmentFromMsg(file, extRegex);
}
async function extractFromEml(file, extRegex) {
  return extractAttachmentFromEml(file, extRegex);
}

// ---------------------------------------------------------------------------
// PDF table parser
// Strategy: find the header row by keywords, then for every data-row cell
// find its nearest header item by x-distance and classify by that header's
// text.  Columns with unrecognised headers (Activity code, Quantity, Area…)
// are silently ignored — no hardcoded column positions needed.
// ---------------------------------------------------------------------------
async function parsePdfToDataRows(pdfData) {
  const lib = await getPdfJs();
  if (!lib) throw new Error('PDF.js could not be loaded. Check your internet connection.');

  const pdf = await lib.getDocument({ data: pdfData }).promise;

  // Collect text items from all pages
  const allItems = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page    = await pdf.getPage(p);
    const content = await page.getTextContent();
    const vp      = page.getViewport({ scale: 1 });
    for (const item of content.items) {
      if (!item.str?.trim()) continue;
      allItems.push({
        str:  item.str.trim(),
        x:    item.transform[4],
        y:    vp.height - item.transform[5],  // PDF y is bottom-up, flip
        page: p
      });
    }
  }
  if (!allItems.length) return [];

  // Group into row buckets by y-coordinate (tolerance ±3 pt, per page)
  const Y_TOL   = 3;
  const buckets = [];
  for (const item of allItems) {
    const b = buckets.find(b => b.page === item.page && Math.abs(b.avgY - item.y) <= Y_TOL);
    if (b) {
      b.items.push(item);
      b.avgY = b.items.reduce((s, i) => s + i.y, 0) / b.items.length;
    } else {
      buckets.push({ avgY: item.y, page: item.page, items: [item] });
    }
  }
  buckets.sort((a, b) => a.page !== b.page ? a.page - b.page : a.avgY - b.avgY);
  const rows = buckets.map(b => b.items.sort((a, c) => a.x - c.x));

  // Find header row — must contain "site" AND at least one of:
  //   "activity"  (TOC PDF:     Site ID | Activity code | Activity Description | Quantity)
  //   "item"      (regular PDF: Request # | Site_ID | Facing # | Item_Description | …)
  //   "request"   (regular PDF — same)
  let hdrIdx = -1;
  for (let i = 0; i < rows.length; i++) {
    const txt = rows[i].map(x => x.str.toLowerCase()).join(' ');
    if (txt.includes('site') &&
        (txt.includes('activity') || txt.includes('item') || txt.includes('request'))) {
      hdrIdx = i; break;
    }
  }
  if (hdrIdx < 0) throw new Error(
    'Could not find table header in PDF. ' +
    'Expected a row containing "Site" and "Activity" (or "Item"/"Request").'
  );

  // Classify each header cell into a logical column key.
  // "Activity code" → 'req' (reused for the activity code in TOC PDFs)
  // "Activity Description" / "Item Description" / "Item" → 'item'
  // Unrecognised headers (Quantity, Area, …) → null (ignored)
  const classifyHdr = (t) => {
    if (t.includes('site'))                                              return 'site';
    if (t.includes('activity') && t.includes('code'))                   return 'req';
    if (t.includes('activity') || t.includes('descript') ||
        t.includes('item'))                                              return 'item';
    if (t.includes('request'))                                          return 'req';
    if (t.includes('facing'))                                           return 'facing';
    return null;
  };

  const headerMap = rows[hdrIdx].map(h => ({ x: h.x, key: classifyHdr(h.str.toLowerCase()) }));

  // Extract data rows: each cell is assigned to its nearest header by x-distance,
  // then kept only if that header has a recognised key.
  const dataRows = [];
  for (let i = hdrIdx + 1; i < rows.length; i++) {
    if (!rows[i].length) continue;
    const cells = { req: [], site: [], facing: [], item: [] };

    for (const cell of rows[i]) {
      const nearestHdr = headerMap.reduce((best, h) =>
        Math.abs(h.x - cell.x) < Math.abs(best.x - cell.x) ? h : best
      );
      if (nearestHdr.key) cells[nearestHdr.key].push(cell.str);
    }

    const row = {
      pos:     dataRows.length,
      request: normalizeStr(cells.req.join(' ')),
      siteId:  normalizeStr(cells.site.join(' ')),
      facing:  normalizeStr(cells.facing.join(' ')),
      item:    normalizeStr(cells.item.join(' '))
    };
    if (row.siteId || row.request) dataRows.push(row);
  }
  return dataRows;
}

// ---------------------------------------------------------------------------
// Build folder data map
// Regular sub-folders (numeric name)  → Excel attachment
// TOC sub-folders (name starts "TOC") → PDF attachment
// Returns: Map<key, { type:'excel'|'pdf', fileName, data:Uint8Array }>
// ---------------------------------------------------------------------------
async function buildFolderDataMap(allFiles) {
  // Group uploaded files by sub-folder name
  const bySubFolder = new Map();
  for (const file of allFiles) {
    const parts = file.webkitRelativePath.split('/');
    if (parts.length < 2) continue;
    const subName = parts[1];
    if (!bySubFolder.has(subName)) bySubFolder.set(subName, []);
    bySubFolder.get(subName).push(file);
  }

  const dataMap = new Map();

  for (const [subName, files] of bySubFolder) {
    const isTocFolder = /^TOC/i.test(subName);

    if (isTocFolder) {
      // ── TOC folder → want PDF attachment ──────────────────────────────────
      const tocKey = normalizeTocKey(subName);
      if (!tocKey || dataMap.has(tocKey)) continue;

      // Direct PDF file in folder
      const pdfFile = files.find(f => PDF_EXT.test(f.name));
      if (pdfFile) {
        const buf = await pdfFile.arrayBuffer();
        dataMap.set(tocKey, { type: 'pdf', fileName: pdfFile.name, data: new Uint8Array(buf) });
        continue;
      }
      // .msg → extract PDF attachment
      const msgFile = files.find(f => /\.msg$/i.test(f.name));
      if (msgFile) {
        const r = await extractFromMsg(msgFile, PDF_EXT);
        if (r) { dataMap.set(tocKey, { type: 'pdf', ...r }); continue; }
      }
      // .eml → extract PDF attachment
      const emlFile = files.find(f => /\.eml$/i.test(f.name));
      if (emlFile) {
        const r = await extractFromEml(emlFile, PDF_EXT);
        if (r) { dataMap.set(tocKey, { type: 'pdf', ...r }); continue; }
      }

    } else {
      // ── Regular (numbered) folder → want Excel attachment ─────────────────
      const fn = normalizeFolderNumber(subName);
      if (fn == null || dataMap.has(fn)) continue;

      // Direct Excel file
      const xlFile = files.find(f => EXCEL_EXT.test(f.name));
      if (xlFile) {
        const buf = await xlFile.arrayBuffer();
        dataMap.set(fn, { type: 'excel', fileName: xlFile.name, data: new Uint8Array(buf) });
        continue;
      }
      // .msg → extract Excel attachment
      const msgFile = files.find(f => /\.msg$/i.test(f.name));
      if (msgFile) {
        const r = await extractFromMsg(msgFile, EXCEL_EXT);
        if (r) { dataMap.set(fn, { type: 'excel', ...r }); continue; }
      }
      // .eml → extract Excel attachment
      const emlFile = files.find(f => /\.eml$/i.test(f.name));
      if (emlFile) {
        const r = await extractFromEml(emlFile, EXCEL_EXT);
        if (r) { dataMap.set(fn, { type: 'excel', ...r }); continue; }
      }
    }
  }

  return dataMap;
}

// ---------------------------------------------------------------------------
// File Picking — TSR File
// ---------------------------------------------------------------------------
el('tsrval-btn-tsr').addEventListener('click', () => el('tsrval-input-tsr').click());

el('tsrval-input-tsr').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  clearError();
  showProgress('tsrval-tsr-progress', true);
  el('tsrval-tsr-filename').textContent = 'Reading…';
  try {
    const data = await file.arrayBuffer();
    _tsrWorkbook = XLSX.read(data, { type: 'array', cellDates: true });
    el('tsrval-tsr-filename').textContent = file.name;
    el('tsrval-card-tsr').classList.add('loaded');
  } catch (err) {
    showError('Failed to open TSR file: ' + err.message);
    _tsrWorkbook = null;
    el('tsrval-tsr-filename').textContent = 'No file selected';
    el('tsrval-card-tsr').classList.remove('loaded');
  }
  showProgress('tsrval-tsr-progress', false);
  checkReady();
});

// ---------------------------------------------------------------------------
// File Picking — TSR Mails Folder
// ---------------------------------------------------------------------------
el('tsrval-btn-folder').addEventListener('click', () => el('tsrval-input-folder').click());

el('tsrval-input-folder').addEventListener('change', (e) => {
  const files = Array.from(e.target.files);
  if (!files.length) return;
  clearError();
  showProgress('tsrval-folder-progress', true);
  setTimeout(() => {
    _folderFiles = files;
    const rootFolder  = files[0].webkitRelativePath.split('/')[0];
    const interesting = files.filter(f => /\.(xlsx|xls|xlsb|msg|eml|pdf)$/i.test(f.name)).length;
    el('tsrval-folder-filename').textContent =
      rootFolder + '  (' + interesting + ' mail/Excel/PDF file' +
      (interesting !== 1 ? 's' : '') + ' across ' + files.length + ' total)';
    el('tsrval-card-folder').classList.add('loaded');
    showProgress('tsrval-folder-progress', false);
    checkReady();
  }, 50);
});

// ---------------------------------------------------------------------------
// Inputs
// ---------------------------------------------------------------------------
el('tsrval-sub-number').addEventListener('input', checkReady);

// ---------------------------------------------------------------------------
// Validate Button
// ---------------------------------------------------------------------------
el('tsrval-btn-analyze').addEventListener('click', async () => {
  clearError();
  el('tsrval-loading').style.display = 'flex';
  el('tsrval-results').style.display = 'none';
  await new Promise(r => setTimeout(r, 50));
  try {
    await runValidation();
  } catch (err) {
    showError(err.message || 'An unexpected error occurred.');
  } finally {
    el('tsrval-loading').style.display = 'none';
  }
});

// ---------------------------------------------------------------------------
// Core Validation
// ---------------------------------------------------------------------------
async function runValidation() {
  const subNum = parseInt(el('tsrval-sub-number').value.trim(), 10);
  if (isNaN(subNum)) throw new Error('Please enter a valid submission number.');


  // ── 1. Parse TSR sheet ─────────────────────────────────────────────────────
  const tsrSheet = _tsrWorkbook.Sheets['PO Break Down- Contractor & Acc'];
  if (!tsrSheet) {
    const available = Object.keys(_tsrWorkbook.Sheets).join(', ');
    throw new Error('Sheet "PO Break Down- Contractor & Acc" not found. Available: ' + available);
  }

  const tsrRows = XLSX.utils.sheet_to_json(tsrSheet, { header: 1, defval: null });

  // Find column header row: must contain BOTH "item description" AND "site id"
  let hdrIdx = -1;
  for (let i = 0; i < tsrRows.length; i++) {
    const row = tsrRows[i];
    if (!row) continue;
    let hasItem = false, hasSite = false;
    for (const c of row) {
      if (!c || c instanceof Date) continue;
      const h = c.toString().toLowerCase().trim();
      if (h.includes('item description')) hasItem = true;
      if (h.includes('site id') || h === 'site_id' || h.includes('site_id')) hasSite = true;
    }
    if (hasItem && hasSite) { hdrIdx = i; break; }
  }
  if (hdrIdx < 0) throw new Error(
    'Cannot find column header row (needs both "Item Description" and "Site ID" in same row).'
  );

  const tsrHeader = tsrRows[hdrIdx];

  // Detect columns
  let colSiteId = -1, colFacing = -1, colItemDesc = -1, colCert = -1;
  let colSub = -1, colComments = -1, colConStatus = -1;

  for (let c = 0; c < tsrHeader.length; c++) {
    const raw = tsrHeader[c];
    if (raw == null || raw instanceof Date) continue;
    const h = raw.toString().trim().toLowerCase();
    if (colSiteId    < 0 && (h === 'site id' || h === 'site_id' || h.includes('site id')))  colSiteId    = c;
    if (colFacing    < 0 && h.includes('facing'))                                            colFacing    = c;
    if (colItemDesc  < 0 && h.includes('item description'))                                  colItemDesc  = c;
    if (colCert      < 0 && h.includes('certificate'))                                       colCert      = c;
    if (colComments  < 0 && h.includes('contractor') && h.includes('comment'))              colComments  = c;
    if (colConStatus < 0 && h.includes('contractor') && h.includes('status'))               colConStatus = c;
    if (colSub < 0 && !h.includes('date') && (
        h.includes('submission') ||
        (h.includes('sub') && (h.includes('#') || h.includes('no') || h.includes('num')))
    )) colSub = c;
  }

  if (colSiteId    < 0) colSiteId    = 7;
  if (colFacing    < 0) colFacing    = 9;
  if (colItemDesc  < 0) colItemDesc  = 13;
  if (colCert      < 0) colCert      = 20;
  if (colConStatus < 0) colConStatus = 21;  // V
  if (colSub       < 0) colSub       = 22;  // W
  if (colComments  < 0) colComments  = 24;  // Y

  const colInfo = [
    'Header row: '  + (hdrIdx + 1),
    'Site ID: '     + colIndexToLetter(colSiteId),
    'Facing #: '    + colIndexToLetter(colFacing),
    'Item Desc: '   + colIndexToLetter(colItemDesc),
    'Cert #: '      + colIndexToLetter(colCert),
    'Con.Status: '  + colIndexToLetter(colConStatus),
    'Sub #: '       + colIndexToLetter(colSub),
    'Comments: '    + colIndexToLetter(colComments)
  ].join(' | ');

  // ── 2. Filter rows for this submission ─────────────────────────────────────
  const filteredRows = [];
  for (let i = hdrIdx + 1; i < tsrRows.length; i++) {
    const row = tsrRows[i];
    if (!row || row.every(c => c == null || c === '' || (c instanceof Date && isNaN(c)))) continue;
    if (extractNumber(row[colSub]) !== subNum) continue;

    const rawComment  = normalizeStr(row[colComments]);
    const conStatus   = normalizeStr(row[colConStatus]).toUpperCase().trim();
    const isToc       = conStatus === 'TOC';

    // For TOC rows: folder key is "TOC N" string.
    // For regular rows: folder key is an integer (digits at the START of comment only,
    // so "TOC 1" correctly gives null here, not 1).
    const folderKey = isToc
      ? normalizeTocKey(rawComment)
      : normalizeFolderNumber(rawComment);

    filteredRows.push({
      tsrExcelRow: i + 1,
      siteId:      normalizeStr(row[colSiteId]),
      facing:      normalizeStr(row[colFacing]),
      itemDesc:    normalizeStr(row[colItemDesc]),
      cert:        normalizeStr(row[colCert]),
      isToc,
      folderKey,
      rawComment
    });
  }

  if (filteredRows.length === 0) {
    const sample = [];
    for (let i = hdrIdx + 1; i < Math.min(hdrIdx + 50, tsrRows.length); i++) {
      const v = (tsrRows[i] || [])[colSub];
      if (v != null && !(v instanceof Date) && v !== '') {
        const s = String(v);
        if (!sample.includes(s)) { sample.push(s); if (sample.length >= 5) break; }
      }
    }
    const hint = sample.length
      ? 'Column ' + colIndexToLetter(colSub) + ' sample: ' + sample.join(', ') + '.'
      : 'Column ' + colIndexToLetter(colSub) + ' appears empty near the header.';
    throw new Error(
      'No rows found for Submission #' + subNum + '. ' + hint +
      ''
    );
  }

  // ── 3. Extract attachments from mail files ─────────────────────────────────
  setLoadingText('Extracting attachments from mail files…');
  const folderDataMap = await buildFolderDataMap(_folderFiles);

  // ── 4. Group TSR rows by folder key ───────────────────────────────────────
  const groups = new Map();
  for (const row of filteredRows) {
    const key = row.folderKey;  // integer for regular, "TOC N" for TOC, null if unknown
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  }

  // ── 5. Validate each group ─────────────────────────────────────────────────
  setLoadingText('Validating submission…');
  const folderResults = [];

  const sortedKeys = [...groups.keys()].sort((a, b) => {
    if (a == null) return 1; if (b == null) return -1;
    // Sort: regular folders first (integers), then TOC folders (strings)
    if (typeof a === 'number' && typeof b === 'number') return a - b;
    if (typeof a === 'number') return -1;
    if (typeof b === 'number') return 1;
    return String(a).localeCompare(String(b));
  });

  for (const folderKey of sortedKeys) {
    const groupRows = groups.get(folderKey);
    const isTocGroup = typeof folderKey === 'string' && folderKey.startsWith('TOC');

    // Handle unknown folder key (null)
    if (folderKey == null) {
      folderResults.push({
        folderKey: null, isToc: false, fileName: null,
        globalError: 'Folder number could not be determined from Contractor Comments',
        rows: groupRows.map(r => ({
          ...r, status: 'error',
          issues: ['No folder number in Contractor Comments: "' + r.rawComment + '"'],
          excelPos: null
        }))
      });
      continue;
    }

    // Retrieve folder data
    const folderData = folderDataMap.get(folderKey);
    if (!folderData) {
      const inFolder = _folderFiles.filter(f => {
        const sub = f.webkitRelativePath.split('/')[1] ?? '';
        return isTocGroup
          ? normalizeTocKey(sub) === folderKey
          : normalizeFolderNumber(sub) === folderKey;
      });
      const found = inFolder.length
        ? 'Files found: ' + inFolder.map(f => f.name).join(', ')
        : 'No files found in that sub-folder.';
      folderResults.push({
        folderKey, isToc: isTocGroup, fileName: null,
        globalError:
          'Could not extract ' + (isTocGroup ? 'PDF' : 'Excel') +
          ' from folder "' + folderKey + '". ' + found +
          ' (supports direct file, Outlook .msg, or .eml)',
        rows: groupRows.map(r => ({
          ...r, status: 'error',
          issues: ['No ' + (isTocGroup ? 'PDF' : 'Excel') + ' data for folder "' + folderKey + '"'],
          excelPos: null
        }))
      });
      continue;
    }

    // Parse the attachment into data rows
    let dataRows;
    try {
      if (folderData.type === 'pdf') {
        setLoadingText('Parsing PDF for ' + folderKey + '…');
        dataRows = await parsePdfToDataRows(folderData.data);
      } else {
        // Excel
        const wb  = XLSX.read(folderData.data, { type: 'array', cellDates: false });
        const ws  = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
        let exHdrIdx = 0;
        for (let i = 0; i < Math.min(raw.length, 20); i++) {
          if (raw[i].some(c => c && c.toString().toLowerCase().includes('request'))) {
            exHdrIdx = i; break;
          }
        }
        const exHdr = raw[exHdrIdx];
        let cReq = -1, cSite = -1, cFacing = -1, cItem = -1;
        for (let c = 0; c < exHdr.length; c++) {
          const h = (exHdr[c] ?? '').toString().trim().toLowerCase();
          if (cReq    < 0 && h.includes('request'))                                                cReq    = c;
          if (cSite   < 0 && (h.includes('site_id') || h.includes('site id') || h === 'site_id')) cSite   = c;
          if (cFacing < 0 && h.includes('facing'))                                                 cFacing = c;
          if (cItem   < 0 && h.includes('item') && h.includes('desc'))                            cItem   = c;
        }
        if (cReq    < 0) cReq    = 0;
        if (cSite   < 0) cSite   = 1;
        if (cFacing < 0) cFacing = 2;
        if (cItem   < 0) cItem   = 4;
        dataRows = raw.slice(exHdrIdx + 1)
          .map((row, idx) => ({
            pos:     idx,
            request: normalizeStr(row[cReq]),
            siteId:  normalizeStr(row[cSite]),
            facing:  normalizeStr(row[cFacing]),
            item:    normalizeStr(row[cItem])
          }))
          .filter(r => r.request !== '' || r.siteId !== '');
      }
    } catch (err) {
      folderResults.push({
        folderKey, isToc: isTocGroup, fileName: folderData.fileName,
        globalError: 'Failed to parse ' + folderData.type.toUpperCase() + ': ' + err.message,
        rows: groupRows.map(r => ({
          ...r, status: 'error',
          issues: ['Parse error: ' + err.message], excelPos: null
        }))
      });
      continue;
    }

    // ── Phase 1: Match each row ───────────────────────────────────────────────
    //
    // Regular folders: key = Site ID + Facing # + Item code prefix
    //   e.g.  "4710 || HD4770 || TX04"
    //
    // TOC folders: key = Site ID + stripped description text (no code prefix,
    //   no Facing #, no Certificate # — the PDF only has Site ID + Activity Description)
    //   TSR "EX01 - Site Visit" → stripped → "site visit"
    //   PDF "Site Visit"        → lower    → "site visit"  ← same key

    const buildComboKey = isTocGroup
      ? r => normalizeLower(r.siteId) + '||' + normalizeActivityCode(r.request)
      : r => normalizeLower(r.siteId) + '||' + normalizeLower(r.facing) + '||' + itemMatchKey(r.item);

    const byCombo = new Map();
    for (const r of dataRows) {
      const k = buildComboKey(r);
      if (!byCombo.has(k)) byCombo.set(k, r);
    }

    const validatedRows = [];
    for (const tsrRow of groupRows) {
      const issues = [];
      let status   = 'pass';
      let matchPos = null;

      // Build the lookup key for this TSR row.
      // TOC:     Site ID  +  normalised activity code
      //   TSR "EX01 - Site Visit" → itemMatchKey → "EX01"
      //   PDF "EX.01"             → normalizeActivityCode(r.req) → "EX01"  ← same key
      // Regular: Site ID + Facing # + Item prefix
      const key = isTocGroup
        ? normalizeLower(tsrRow.siteId) + '||' + itemMatchKey(tsrRow.itemDesc)
        : normalizeLower(tsrRow.siteId) + '||' + normalizeLower(tsrRow.facing) + '||' + itemMatchKey(tsrRow.itemDesc);

      const matched = byCombo.get(key);

      if (!matched) {
        if (isTocGroup) {
          // TOC fallback: same site, try loose activity-code contains-match
          const codeKey = itemMatchKey(tsrRow.itemDesc); // e.g. "EX01"
          const partial = dataRows.find(r =>
            normalizeLower(r.siteId) === normalizeLower(tsrRow.siteId) &&
            normalizeActivityCode(r.request).includes(codeKey)
          );
          if (partial) {
            issues.push(
              'Activity Code mismatch — TSR: "' + codeKey +
              '", PDF: "' + partial.request + '" (normalised: "' +
              normalizeActivityCode(partial.request) + '")'
            );
            matchPos = partial.pos;
          } else {
            issues.push(
              'Row not found in PDF — no match for ' +
              'Site ID "' + tsrRow.siteId + '" + ' +
              'Activity Code "' + codeKey + '" ' +
              '(from TSR item: "' + tsrRow.itemDesc + '")'
            );
          }
        } else {
          // Regular fallback: try by Site ID + Item prefix, ignore Facing
          const partial = dataRows.find(r =>
            normalizeLower(r.siteId) === normalizeLower(tsrRow.siteId) &&
            itemMatchKey(r.item)     === itemMatchKey(tsrRow.itemDesc)
          );
          if (partial) {
            issues.push('Facing # mismatch — TSR: "' + tsrRow.facing + '", file: "' + partial.facing + '"');
            if (normalizeLower(partial.request) !== normalizeLower(tsrRow.cert))
              issues.push('Certificate # mismatch — TSR: "' + tsrRow.cert + '", file (Request #): "' + partial.request + '"');
            matchPos = partial.pos;
          } else {
            issues.push(
              'Row not found — no match for ' +
              'Site ID "' + tsrRow.siteId + '" + ' +
              'Facing "' + (tsrRow.facing || '(blank)') + '" + ' +
              'Item "' + tsrRow.itemDesc + '"'
            );
          }
        }
        status = 'fail';
      } else {
        matchPos = matched.pos;
        if (!isTocGroup) {
          // Certificate # only checked for regular (Excel) folders — PDF has no Request #
          if (normalizeLower(matched.request) !== normalizeLower(tsrRow.cert)) {
            issues.push('Certificate # mismatch — TSR: "' + tsrRow.cert +
              '", file (Request #): "' + matched.request + '"');
            status = 'fail';
          }
        }
      }
      validatedRows.push({ ...tsrRow, status, issues, excelPos: matchPos });
    }

    // ── Phase 2: Bidirectional order check ────────────────────────────────────
    for (let i = 0; i < validatedRows.length; i++) {
      const posI = validatedRows[i].excelPos;
      if (posI === null) continue;

      for (let j = i + 1; j < validatedRows.length; j++) {
        const posJ = validatedRows[j].excelPos;
        if (posJ === null) continue;
        if (posJ < posI) {
          const ref = validatedRows[j].siteId + ' / ' + itemMatchKey(validatedRows[j].itemDesc);
          if (!validatedRows[i].issues.some(x => x.startsWith('Order error'))) {
            validatedRows[i].issues.push(
              'Order error — in the file this item (pos ' + (posI + 1) + ') comes ' +
              'after "' + ref + '" (pos ' + (posJ + 1) + '), but in the TSR it appears before it.'
            );
            validatedRows[i].status = 'fail';
          }
          break;
        }
      }
      for (let j = i - 1; j >= 0; j--) {
        const posJ = validatedRows[j].excelPos;
        if (posJ === null) continue;
        if (posJ > posI) {
          const ref = validatedRows[j].siteId + ' / ' + itemMatchKey(validatedRows[j].itemDesc);
          if (!validatedRows[i].issues.some(x => x.startsWith('Order error'))) {
            validatedRows[i].issues.push(
              'Order error — in the file this item (pos ' + (posI + 1) + ') comes ' +
              'before "' + ref + '" (pos ' + (posJ + 1) + '), but in the TSR it appears after it.'
            );
            validatedRows[i].status = 'fail';
          }
          break;
        }
      }
    }

    folderResults.push({
      folderKey, isToc: isTocGroup,
      fileName: folderData.fileName,
      globalError: null, rows: validatedRows
    });
  }

  // ── 6. Render ──────────────────────────────────────────────────────────────
  renderResults(subNum, filteredRows.length, folderResults, colInfo);
}

// ---------------------------------------------------------------------------
// Render Results
// ---------------------------------------------------------------------------
function renderResults(subNum, totalRows, folderResults, colInfo) {
  const allRows   = folderResults.flatMap(f => f.rows);
  const passCount = allRows.filter(r => r.status === 'pass').length;
  const failCount = allRows.filter(r => r.status !== 'pass').length;
  const allOk     = failCount === 0;

  const summaryHtml = `
    <div class="tsrval-col-info">&#8505;&nbsp;${escHtml(colInfo)}</div>
    <div class="poc-stats-row tsrval-stats-row">
      <div class="poc-stat-box fin-stat-red">
        <span class="poc-stat-label">Submission #</span>
        <span class="poc-stat-value">${subNum}</span>
      </div>
      <div class="poc-stat-box poc-stat-total">
        <span class="poc-stat-label">Total Rows</span>
        <span class="poc-stat-value">${totalRows}</span>
      </div>
      <div class="poc-stat-box tsrval-stat-pass">
        <span class="poc-stat-label">Passed</span>
        <span class="poc-stat-value">${passCount}</span>
      </div>
      <div class="poc-stat-box tsrval-stat-fail">
        <span class="poc-stat-label">Failed</span>
        <span class="poc-stat-value">${failCount}</span>
      </div>
    </div>
    ${allOk
      ? '<div class="tsrval-all-ok">&#10003; All rows validated successfully.</div>'
      : '<div class="tsrval-has-errors">&#9888; ' + failCount + ' issue' + (failCount > 1 ? 's' : '') +
        ' found — review the details below before submitting.</div>'
    }`;

  let foldersHtml = '';
  for (const folder of folderResults) {
    const title    = folder.folderKey != null ? String(folder.folderKey) : 'Unknown Folder';
    const typeTag  = folder.isToc ? ' <span class="tsrval-toc-tag">TOC</span>' : '';
    const filePart = folder.fileName ? ' &mdash; <em>' + escHtml(folder.fileName) + '</em>' : '';
    const passN    = folder.rows.filter(r => r.status === 'pass').length;
    const failN    = folder.rows.length - passN;
    const ok       = failN === 0;

    foldersHtml += `
      <div class="tsrval-folder-section${ok ? ' tsrval-folder-ok' : ' tsrval-folder-fail'}${folder.isToc ? ' tsrval-folder-toc' : ''}">
        <div class="tsrval-folder-header">
          <span class="tsrval-folder-title">&#128193; ${escHtml(title)}${typeTag}${filePart}</span>
          <span class="tsrval-folder-badge ${ok ? 'tsrval-badge-pass' : 'tsrval-badge-fail'}">
            ${ok ? '&#10003; All Pass' : failN + ' issue' + (failN > 1 ? 's' : '')}
          </span>
        </div>`;

    if (folder.globalError)
      foldersHtml += `<div class="tsrval-global-error">&#9888; ${escHtml(folder.globalError)}</div>`;

    foldersHtml += `
        <div class="table-wrapper">
          <table class="tsrval-table">
            <thead><tr>
              <th>#</th><th>TSR Row</th><th>Site ID</th><th>Facing #</th>
              <th>Item Description</th><th>Certificate #</th><th>Status</th><th>Issues</th>
            </tr></thead>
            <tbody>`;

    folder.rows.forEach((row, i) => {
      const cls   = row.status === 'pass' ? 'tsrval-row-pass' : 'tsrval-row-fail';
      const badge = row.status === 'pass'
        ? '<span class="tsrval-badge tsrval-badge-pass">&#10003; Pass</span>'
        : '<span class="tsrval-badge tsrval-badge-fail">&#10007; Fail</span>';
      const issHtml = (row.issues || [])
        .map(s => `<div class="tsrval-issue-item">&bull; ${escHtml(s)}</div>`).join('');
      foldersHtml += `
              <tr class="${cls}">
                <td style="text-align:center">${i + 1}</td>
                <td style="text-align:center">${row.tsrExcelRow}</td>
                <td>${escHtml(row.siteId)}</td>
                <td>${escHtml(row.facing)}</td>
                <td class="tsrval-item-col">${escHtml(row.itemDesc)}</td>
                <td>${escHtml(row.cert)}</td>
                <td style="text-align:center">${badge}</td>
                <td class="tsrval-issues-col">${issHtml}</td>
              </tr>`;
    });

    foldersHtml += `</tbody></table></div></div>`;
  }

  el('tsrval-summary').innerHTML = summaryHtml;
  el('tsrval-folder-results').innerHTML = foldersHtml;
  el('tsrval-results').style.display = 'block';
  el('tsrval-results').scrollIntoView({ behavior: 'smooth' });
}

})();
