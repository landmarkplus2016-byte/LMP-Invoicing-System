/* ============================================================
   LMP Invoicing System — Contractor Invoices logic
   Wrapped in IIFE to avoid naming conflicts with other apps.
   ============================================================ */
(function () {

// ---------------------------------------------------------------------------
// Canonical contractor list (sourced from list.xlsx)
// ---------------------------------------------------------------------------
const CONTRACTOR_LIST = ['Connect', 'DAM Tel', 'El-Khayal', 'New Plan', 'Upper Telecom'];

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------
let _wb     = null;   // loaded workbook
let _groups = null;   // Map<contractorName, row[]>

// ---------------------------------------------------------------------------
// Name normalisation — Levenshtein-based fuzzy match to canonical list
// ---------------------------------------------------------------------------
function levenshtein(a, b) {
  const m = a.length, n = b.length;
  const dp = Array.from({ length: m + 1 }, (_, i) =>
    Array.from({ length: n + 1 }, (_, j) => (i === 0 ? j : j === 0 ? i : 0))
  );
  for (let i = 1; i <= m; i++)
    for (let j = 1; j <= n; j++)
      dp[i][j] = a[i - 1] === b[j - 1]
        ? dp[i - 1][j - 1]
        : 1 + Math.min(dp[i - 1][j - 1], dp[i - 1][j], dp[i][j - 1]);
  return dp[m][n];
}

function normalizeContractor(raw) {
  const s  = String(raw ?? '').trim();
  if (!s) return null;
  const sl = s.toLowerCase().replace(/[-_.]/g, ' ').replace(/\s+/g, ' ').trim();
  // Detect In-House before fuzzy matching
  if (/^in[\s-]?house$/i.test(sl) || sl === 'inhouse') return 'In-House';
  let best = null, bestDist = Infinity;
  for (const name of CONTRACTOR_LIST) {
    const d = levenshtein(sl, name.toLowerCase().replace(/[-_.]/g, ' '));
    if (d < bestDist) { bestDist = d; best = name; }
  }
  // Accept match if within threshold (40% of canonical name length, min 3)
  const threshold = Math.max(3, Math.floor((best?.length ?? 4) * 0.4));
  return bestDist <= threshold ? best : s;
}

// ---------------------------------------------------------------------------
// File picker
// ---------------------------------------------------------------------------
document.getElementById('btn-pick-con-track').addEventListener('click', () => {
  document.getElementById('con-track-input').click();
});

document.getElementById('con-track-input').addEventListener('change', async (e) => {
  clearConError();
  const file = e.target.files[0];
  if (!file) return;
  conFileProgress(true);
  try {
    const buf = await file.arrayBuffer();
    _wb = XLSX.read(buf, { type: 'array', cellDates: true });
    document.getElementById('con-track-filename').textContent = file.name;
    document.getElementById('card-con-track').classList.add('loaded');
    conFileProgress(false);
    await triggerAnalysis();
  } catch (err) {
    conFileProgress(false);
    showConError('Failed to open file: ' + err.message);
  }
});

// ---------------------------------------------------------------------------
// Analysis
// ---------------------------------------------------------------------------
async function triggerAnalysis() {
  clearConError();
  conLoading(true);
  document.getElementById('con-results').style.display = 'none';
  await new Promise(r => setTimeout(r, 50));
  try {
    _groups = analyzeContractors();
    renderSummary(_groups);
    conLoading(false);
    document.getElementById('con-results').style.display = 'block';
    document.getElementById('con-results').scrollIntoView({ behavior: 'smooth' });
  } catch (err) {
    conLoading(false);
    showConError(err.message || 'Unexpected error during analysis.');
  }
}

function analyzeContractors() {
  const sheet = _wb.Sheets['Invoicing Track'];
  if (!sheet) throw new Error('Wrong file — expected a sheet named "Invoicing Track".');

  const allRows  = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
  const header   = allRows[3] ?? [];   // row 4 (0-indexed 3) is the header
  const dataRows = allRows.slice(4);   // data from row 5 onward

  // Detect columns by header name
  let cJob = -1, cSite = -1, cFacing = -1, cLineItem = -1, cPrice = -1;
  let cTaskDate = -1, cAcceptWeek = -1, cConInv = -1, cContractor = -1;

  header.forEach((h, i) => {
    const t = String(h ?? '').trim().toLowerCase();
    if (t.includes('job code'))                                                   cJob        = i;
    if (t.includes('logical site'))                                               cSite       = i;
    if ((t === 'facing' || t.includes('facing')) && !t.includes('re'))           cFacing     = i;
    if (t.includes('line item'))                                                  cLineItem   = i;
    if (t.includes('new total'))                                                  cPrice      = i;
    if (t.includes('task date'))                                                  cTaskDate   = i;
    if (t.includes('acceptance week'))                                            cAcceptWeek = i;
    if (t.includes('contractor invoice'))                                         cConInv     = i;
    if (t === 'contractor' || (t.includes('contractor') && !t.includes('invoice') && cContractor === -1))
                                                                                  cContractor = i;
  });

  cLineItem = 18; // always use col 18 — hardcoded same as TSR app for Invoicing Track
  if (cTaskDate   < 0) throw new Error('Column "Task Date" not found in the Invoicing Track header row.');
  if (cAcceptWeek < 0) throw new Error('Column "Acceptance Week" not found.');
  if (cConInv     < 0) throw new Error('Column "Contractor Invoice #" (or similar) not found.');
  if (cContractor < 0) throw new Error('Column "Contractor" not found.');

  function rowYear(val) {
    if (val == null) return 0;
    if (val instanceof Date) return val.getFullYear();
    if (typeof val === 'number') {
      // Excel date serial → JS date (Excel epoch offset 25569 days)
      return new Date(Math.round((val - 25569) * 86400 * 1000)).getFullYear();
    }
    const d = new Date(String(val));
    return isNaN(d) ? 0 : d.getFullYear();
  }

  function notBlank(v) { return v != null && String(v).trim() !== ''; }

  const groups = new Map();

  for (const row of dataRows) {
    if (rowYear(row[cTaskDate]) < 2026) continue;
    if (!notBlank(row[cAcceptWeek]))    continue;
    if (notBlank(row[cConInv]))         continue;  // skip rows where invoice # is already filled

    const norm = normalizeContractor(row[cContractor]);
    if (!norm || norm === 'In-House')   continue;

    if (!groups.has(norm)) groups.set(norm, []);
    groups.get(norm).push({
      jobCode:  cJob      >= 0 ? String(row[cJob]      ?? '').trim() : '',
      siteId:   cSite     >= 0 ? String(row[cSite]     ?? '').trim() : '',
      facing:   cFacing   >= 0 ? String(row[cFacing]   ?? '').trim() : '',
      lineItem: cLineItem >= 0 ? String(row[cLineItem] ?? '').trim() : '',
      price:    cPrice    >= 0 ? toContractorAmount(row[cPrice])      : '',
    });
  }

  if (groups.size === 0) {
    throw new Error(
      'No rows matched the criteria (Task Date ≥ 2026, Acceptance Week filled, ' +
      'Contractor Invoice # blank) — or all matched rows belong to In-House.'
    );
  }
  return groups;
}

// ---------------------------------------------------------------------------
// Render summary table
// ---------------------------------------------------------------------------
function renderSummary(groups) {
  let html = '', grandTotal = 0;
  for (const [name, rows] of groups) {
    const amt = sumPrices(rows);
    grandTotal += amt;
    html += `<tr><td>${esc(name)}</td><td class="num">${fmtPrice(amt)}</td></tr>`;
  }
  html += `<tr class="totals-row"><td><strong>Total</strong></td><td class="num"><strong>${fmtPrice(grandTotal)}</strong></td></tr>`;
  document.getElementById('con-summary-tbody').innerHTML = html;
  const statusEl = document.getElementById('con-export-status');
  statusEl.textContent = '';
  statusEl.className = '';
  document.getElementById('btn-export-con').disabled = false;
}

// ---------------------------------------------------------------------------
// Export button
// ---------------------------------------------------------------------------
document.getElementById('btn-export-con').addEventListener('click', async () => {
  if (!_groups) return;
  const btn      = document.getElementById('btn-export-con');
  const statusEl = document.getElementById('con-export-status');
  btn.disabled       = true;
  statusEl.textContent = 'Generating files…';
  statusEl.className   = '';
  try {
    let done = 0;
    for (const [name, rows] of _groups) {
      await buildAndDownload(name, rows);
      done++;
      statusEl.textContent = `Exporting… ${done} / ${_groups.size}`;
      await new Promise(r => setTimeout(r, 500)); // brief pause between downloads
    }
    statusEl.textContent = `\u2705 Done — ${done} file(s) downloaded.`;
    statusEl.className   = 'export-status-success';
  } catch (err) {
    statusEl.textContent = `\u274c Export failed: ${err.message}`;
    statusEl.className   = 'export-status-error';
  } finally {
    btn.disabled = false;
  }
});

// ---------------------------------------------------------------------------
// Excel colour / style constants (ExcelJS ARGB format)
// ---------------------------------------------------------------------------
const FILL_BLUE  = fill('FF9BC2E6');  // header / title — light steel blue
const FILL_PEACH = fill('FFF8CBAD');  // first row of each Job Code group
const FILL_WHITE = fill('FFFFFFFF');  // subsequent rows in a group
const FILL_GREEN = fill('FF00FF00');  // total row
const THIN       = { style: 'thin', color: { argb: 'FF000000' } };
const ALL_THIN   = { top: THIN, left: THIN, bottom: THIN, right: THIN };

function fill(argb) {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}

function fmtPrice(val) {
  if (val == null || val === '') return '';
  const n = typeof val === 'number' ? val
    : parseFloat(String(val).replace(/[^0-9.-]/g, ''));
  if (isNaN(n)) return String(val);
  return 'EGP ' + n.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
}

function toContractorAmount(raw) {
  const n = typeof raw === 'number' ? raw
    : parseFloat(String(raw ?? '').replace(/[^0-9.-]/g, ''));
  return isNaN(n) ? '' : Math.round(n * 0.7 * 100) / 100;
}

function sumPrices(rows) {
  return rows.reduce((acc, r) => {
    const n = typeof r.price === 'number' ? r.price
      : parseFloat(String(r.price ?? '').replace(/[^0-9.-]/g, ''));
    return acc + (isNaN(n) ? 0 : n);
  }, 0);
}

// ---------------------------------------------------------------------------
// Build workbook and trigger browser download
// ---------------------------------------------------------------------------
async function buildAndDownload(name, rows) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'LMP Invoicing System';
  addDraftSheet(wb, name, rows);
  addDeductionSheet(wb, rows);

  const buf  = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  const url = URL.createObjectURL(blob);
  const a   = Object.assign(document.createElement('a'), {
    href: url, download: name + ' Draft.xlsx'
  });
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ---------------------------------------------------------------------------
// "Draft" sheet — formatted to match Upper Telecom Draft.xlsx
// Layout: data in cols B–F, title merged E3:F4, headers row 5, data row 6+
// ---------------------------------------------------------------------------
function addDraftSheet(wb, contractorName, rows) {
  const ws = wb.addWorksheet('Draft');

  // Column widths  (A = narrow spacer so data begins at B)
  ws.getColumn(1).width = 2;
  ws.getColumn(2).width = 10;   // B — Job Code
  ws.getColumn(3).width = 10;   // C — Site ID
  ws.getColumn(4).width = 12;   // D — Facing
  ws.getColumn(5).width = 55;   // E — Line Item  (wide — long descriptions)
  ws.getColumn(6).width = 18;   // F — Price

  // --- Title: merged E3:F4 — left blank for manual entry ---
  ws.mergeCells('E3:F4');
  const titleCell = ws.getCell('E3');
  titleCell.value     = '';
  titleCell.fill      = FILL_BLUE;
  titleCell.font      = { bold: true, size: 11 };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  titleCell.border    = ALL_THIN;
  // Right border on D3/D4 to visually separate empty cols from title
  ws.getCell('D3').border = { right: THIN };
  ws.getCell('D4').border = { right: THIN };
  ws.getRow(3).height = 18;
  ws.getRow(4).height = 18;

  // --- Header row 5 ---
  styleHeaderRow(ws, 5, ['Job Code', ' Site ID', 'Facing', 'Line Item', 'Amount']);

  // --- Data rows starting at row 6 ---
  let prevJobCode = null;
  rows.forEach((r, i) => {
    const rn      = 6 + i;
    const isFirst = r.jobCode !== prevJobCode;
    if (isFirst) prevJobCode = r.jobCode;
    const rowFill = isFirst ? FILL_PEACH : FILL_WHITE;

    styleDataRow(ws, rn, [r.jobCode, r.siteId, r.facing, r.lineItem, fmtPrice(r.price)], rowFill);
  });

  // --- Total row ---
  const totalRowNum = 6 + rows.length;
  const total       = sumPrices(rows);
  styleTotalRow(ws, totalRowNum, fmtPrice(total));
}

// ---------------------------------------------------------------------------
// "Deduction" sheet — 4 cols (B–E), no Price column, for manual deduction entry
// ---------------------------------------------------------------------------
const DED_COLS = ['B', 'C', 'D', 'E'];

function addDeductionSheet(wb, rows) {
  const ws = wb.addWorksheet('Deduction');

  ws.getColumn(1).width = 2;
  ws.getColumn(2).width = 10;   // B — Job Code
  ws.getColumn(3).width = 10;   // C — Site ID
  ws.getColumn(4).width = 12;   // D — Facing
  ws.getColumn(5).width = 30;   // E — Deduction Amount

  styleHeaderRow(ws, 1, ['Job Code', ' Site ID', 'Facing', 'Deduction Amount'], DED_COLS);

  let prevJobCode = null;
  rows.forEach((r, i) => {
    const rn      = 2 + i;
    const isFirst = r.jobCode !== prevJobCode;
    if (isFirst) prevJobCode = r.jobCode;
    styleDataRow(ws, rn, [r.jobCode, r.siteId, r.facing, ''], isFirst ? FILL_PEACH : FILL_WHITE, DED_COLS);
  });
}

// ---------------------------------------------------------------------------
// Sheet-building helpers
// ---------------------------------------------------------------------------
const DATA_COLS = ['B', 'C', 'D', 'E', 'F'];

function styleHeaderRow(ws, rowNum, labels, cols) {
  cols = cols || DATA_COLS;
  labels.forEach((label, i) => {
    const cell     = ws.getCell(`${cols[i]}${rowNum}`);
    cell.value     = label;
    cell.fill      = FILL_BLUE;
    cell.font      = { bold: true, size: 11 };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border    = ALL_THIN;
  });
  ws.getRow(rowNum).height = 18;
}

function styleDataRow(ws, rowNum, values, rowFill, cols) {
  cols = cols || DATA_COLS;
  values.forEach((val, i) => {
    const cell     = ws.getCell(`${cols[i]}${rowNum}`);
    cell.value     = val;
    cell.fill      = rowFill;
    cell.font      = { size: 11 };
    cell.border    = ALL_THIN;
    cell.alignment = { vertical: 'middle', wrapText: i === 3 }; // wrap Line Item / Deduction
  });
  ws.getRow(rowNum).height = 14.4;
}

function styleTotalRow(ws, rowNum, totalStr) {
  DATA_COLS.forEach((col, i) => {
    const cell     = ws.getCell(`${col}${rowNum}`);
    cell.fill      = FILL_GREEN;
    cell.font      = { bold: true, size: 11 };
    cell.border    = ALL_THIN;
    cell.alignment = { vertical: 'middle' };
    if (i === 0) { cell.value = 'Total'; cell.alignment.horizontal = 'left';  }
    if (i === 4) { cell.value = totalStr; cell.alignment.horizontal = 'right'; }
  });
  ws.getRow(rowNum).height = 14.4;
}

// ---------------------------------------------------------------------------
// UI helpers
// ---------------------------------------------------------------------------
function showConError(msg) {
  const el = document.getElementById('con-error');
  el.textContent  = msg;
  el.style.display = 'block';
  el.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

function clearConError() {
  const el = document.getElementById('con-error');
  el.textContent  = '';
  el.style.display = 'none';
}

function conLoading(on) {
  document.getElementById('con-loading').style.display = on ? 'flex' : 'none';
}

function conFileProgress(on) {
  const el = document.getElementById('con-track-progress');
  if (el) el.style.display = on ? 'block' : 'none';
}

function esc(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

})(); // end IIFE
