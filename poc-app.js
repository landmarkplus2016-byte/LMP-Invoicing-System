/* ============================================================
   LMP Invoicing System — POC Invoice Prep logic
   Wrapped in IIFE to avoid naming conflicts with TSR app.
   ============================================================ */
(function () {

// ---------------------------------------------------------------------------
// Column patterns
// ---------------------------------------------------------------------------
const COL_PATTERNS = {
  jobCode:               ['job code', 'jobcode', 'job_code'],
  siteId:                ['site id', 'siteid', 'site_id'],
  area:                  ['area'],
  vfOwner:               ['vf owner', 'vfowner', 'vf_owner'],
  installationStatus:    ['installation status', 'install status', 'inst. status', 'inst status'],
  installationDate:      ['installation date', 'install date', 'inst. date', 'inst date'],
  installInvoicingDate:  ['installation invoicing date', 'instalation invoicing date', 'install invoicing date',
                          'invoicing date ins', 'invoicing date (ins)', 'ins invoicing date',
                          'inst invoicing date', 'inst. invoicing'],
  migrationStatus:       ['migration status', 'migr. status', 'mig status', 'mig. status'],
  migrationDate:         ['migration date', 'migr. date', 'migr date', 'mig date', 'mig. date'],
  acceptanceStatus:      ['acceptance status', 'accept status', 'fac status'],
  certificate:           ['certificate', 'cert'],
  facDate:               ['fac date', 'fac_date', 'facd ate'],
  migInvoicingDate:      ['migration invoicing date', 'migr invoicing date',
                          'invoicing date mig', 'invoicing date (mig)', 'mig invoicing date',
                          'migr. invoicing'],
  lineItem:              ['line item', 'lineitem', 'line_item'],
  price:                 ['price', 'unit price'],
  totalAmount:           ['total amount', 'total', 'amount'],
  instContractor:        ['inst contractor', 'installation contractor', 'install contractor'],
  migContractor:         ['migr contractor', 'migration contractor', 'mig contractor'],
};

// Headers containing these words are skipped for the given key — keeps
// "INST Contractor Invoice#" / "Contractor Portion ins" from being taken
// as the contractor name column.
const COL_EXCLUDE = {
  instContractor: ['invoice', 'portion'],
  migContractor:  ['invoice', 'portion'],
};

const OUTPUT_COLUMNS = [
  { label: 'Job Code',            key: 'jobCode'           },
  { label: 'Site ID',             key: 'siteId'            },
  { label: 'Area',                key: 'area'              },
  { label: 'VF Owner',            key: 'vfOwner'           },
  { label: 'Contractor',          key: 'contractor'        },
  { label: 'Installation Status', key: 'installationStatus'},
  { label: 'Installation Date',   key: 'installationDate'  },
  { label: 'Migration Status',    key: 'migrationStatus'   },
  { label: 'Migration Date',      key: 'migrationDate'     },
  { label: 'Acceptance Status',   key: 'acceptanceStatus'  },
  { label: 'Certificate',         key: 'certificate'       },
  { label: 'FAC Date',            key: 'facDate'           },
  { label: 'Line Item',           key: 'lineItem'          },
  { label: 'Price',               key: 'totalAmount'       },
  { label: 'Invoice Amount',      key: 'invoiceAmount'     },
];

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
function findColumn(headers, patterns, exclude = []) {
  const lower = headers.map(h => String(h ?? '').toLowerCase().trim());
  for (const pattern of patterns) {
    const idx = lower.findIndex(h => h.includes(pattern) && !exclude.some(x => h.includes(x)));
    if (idx !== -1) return idx;
  }
  return -1;
}

function buildColumnMap(headers) {
  const map = {};
  for (const [key, patterns] of Object.entries(COL_PATTERNS)) {
    map[key] = findColumn(headers, patterns, COL_EXCLUDE[key]);
  }
  return map;
}

function cell(row, idx) {
  if (idx === -1 || idx >= row.length) return '';
  const v = row[idx];
  return (v === null || v === undefined) ? '' : v;
}

function isBlank(val) {
  if (val === '' || val === null || val === undefined) return true;
  if (typeof val === 'string' && val.trim() === '') return true;
  return false;
}

function formatDate(d) {
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const day = String(d.getDate()).padStart(2, '0');
  return `${day}-${months[d.getMonth()]}-${String(d.getFullYear()).slice(-2)}`;
}

// ---------------------------------------------------------------------------
// Header row detection
// ---------------------------------------------------------------------------
const MAX_SCAN_ROWS = 30;

function scoreRow(row) {
  const lower = row.map(c => String(c ?? '').toLowerCase().trim());
  let score = 0;
  for (const patterns of Object.values(COL_PATTERNS)) {
    for (const pattern of patterns) {
      if (lower.some(h => h.includes(pattern))) { score++; break; }
    }
  }
  return score;
}

function detectHeaderRow(rows) {
  let bestIdx = 0;
  let bestScore = -1;
  const limit = Math.min(rows.length, MAX_SCAN_ROWS);
  for (let i = 0; i < limit; i++) {
    const s = scoreRow(rows[i]);
    if (s > bestScore) { bestScore = s; bestIdx = i; }
  }
  return bestIdx;
}

// ---------------------------------------------------------------------------
// Invoice batch values — the labels typed in the Invoicing Date columns
// (e.g. "new oct", "oct"); older rows may hold real dates instead.
// ---------------------------------------------------------------------------
function batchKey(val) {
  if (val instanceof Date) {
    return 'd:' + val.getFullYear() + '-' + (val.getMonth() + 1) + '-' + val.getDate();
  }
  if (isBlank(val)) return '';
  return 't:' + String(val).trim().toLowerCase().replace(/\s+/g, ' ');
}

function batchLabel(val) {
  return val instanceof Date ? formatDate(val) : String(val).trim().replace(/\s+/g, ' ');
}

function collectBatches(dataRows, colMap) {
  const map = new Map();
  function add(val, field) {
    const key = batchKey(val);
    if (!key) return;
    if (!map.has(key)) {
      map.set(key, { key, label: batchLabel(val), isDate: val instanceof Date,
                     time: val instanceof Date ? val.getTime() : 0, ins: 0, mig: 0 });
    }
    map.get(key)[field]++;
  }
  for (const row of dataRows) {
    if (colMap.installInvoicingDate !== -1) add(cell(row, colMap.installInvoicingDate), 'ins');
    if (colMap.migInvoicingDate     !== -1) add(cell(row, colMap.migInvoicingDate),     'mig');
  }
  // Text labels first (alphabetical), then dates newest first
  return [...map.values()].sort((a, b) => {
    if (a.isDate !== b.isDate) return a.isDate ? 1 : -1;
    return a.isDate ? b.time - a.time : a.label.localeCompare(b.label);
  });
}

// ---------------------------------------------------------------------------
// Core processing
// ---------------------------------------------------------------------------
function parseWorkbook(fileData) {
  const workbook = XLSX.read(fileData, { type: 'array', cellDates: true });
  const TARGET_SHEET = 'POC3 Tracking';
  const sheetName = workbook.SheetNames.find(
    n => n.trim().toLowerCase() === TARGET_SHEET.toLowerCase()
  );
  if (!sheetName) {
    throw new Error(`Sheet "${TARGET_SHEET}" not found. Available sheets: ${workbook.SheetNames.join(', ')}`);
  }
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  if (!rows || rows.length < 2) {
    throw new Error('The file appears to be empty or contains only a header row.');
  }

  const headerRowIdx = detectHeaderRow(rows);
  const headers = rows[headerRowIdx];
  const dataRows = rows.slice(headerRowIdx + 1).filter(r => r.some(c => c !== ''));
  const colMap = buildColumnMap(headers);

  if (colMap.installInvoicingDate === -1 && colMap.migInvoicingDate === -1) {
    throw new Error('Neither "Installation Invoicing Date ins" nor "Migration Invoicing Date mig" column was found in the header row.');
  }

  return { headers, dataRows, colMap, batches: collectBatches(dataRows, colMap) };
}

// Step 1 = rows whose Installation Invoicing Date holds the selected batch,
// Step 2 = rows whose Migration Invoicing Date holds it.
function analyzeBatch(parsed, selectedKey) {
  const { headers, dataRows, colMap } = parsed;

  const step1 = colMap.installInvoicingDate === -1 ? [] :
    dataRows.filter(row => batchKey(cell(row, colMap.installInvoicingDate)) === selectedKey);

  const step2 = colMap.migInvoicingDate === -1 ? [] :
    dataRows.filter(row => batchKey(cell(row, colMap.migInvoicingDate)) === selectedKey);

  function extractRow(row, stepLabel) {
    const out = { _step: stepLabel };
    for (const { label, key } of OUTPUT_COLUMNS) {
      if (key === 'invoiceAmount') {
        const rawTotal = parseFloat(String(cell(row, colMap.totalAmount) ?? '').replace(/,/g, '')) || 0;
        out[label] = rawTotal / 2;
      } else if (key === 'contractor') {
        // Installation rows carry the INST contractor, migration rows the MIGR contractor
        out[label] = cell(row, stepLabel === 1 ? colMap.instContractor : colMap.migContractor);
      } else {
        out[label] = cell(row, colMap[key]);
      }
    }
    return out;
  }

  const step1Extracted = step1.map(r => extractRow(r, 1));
  const step2Extracted = step2.map(r => extractRow(r, 2));
  const combined = [...step1Extracted, ...step2Extracted];

  function sumExtracted(rows) {
    return rows.reduce((acc, row) => acc + (typeof row['Invoice Amount'] === 'number' ? row['Invoice Amount'] : 0), 0);
  }
  const step1Amount = sumExtracted(step1Extracted);
  const step2Amount = sumExtracted(step2Extracted);
  const totalAmount = step1Amount + step2Amount;

  const warnings = [];
  if (colMap.installInvoicingDate === -1) {
    warnings.push('Column "Installation Invoicing Date ins" not found — no installation rows can be selected.');
  }
  if (colMap.migInvoicingDate === -1) {
    warnings.push('Column "Migration Invoicing Date mig" not found — no migration rows can be selected.');
  }
  if (colMap.instContractor === -1) {
    warnings.push('Column "INST Contractor" not found — Contractor will be empty for installation rows.');
  }
  if (colMap.migContractor === -1) {
    warnings.push('Column "MIGR Contractor" not found — Contractor will be empty for migration rows.');
  }
  const computedKeys = new Set(['invoiceAmount', 'contractor']);
  for (const { label, key } of OUTPUT_COLUMNS) {
    if (!computedKeys.has(key) && colMap[key] === -1) {
      warnings.push(`Output column "${label}" not found in the source file — it will be empty.`);
    }
  }

  return {
    step1Count: step1.length,
    step2Count: step2.length,
    step1Amount,
    step2Amount,
    totalAmount,
    combined,
    warnings,
    originalHeaders: headers,
    colMap,
  };
}

// ---------------------------------------------------------------------------
// Export to Excel — uses ExcelJS for styling
// ---------------------------------------------------------------------------
const HEADER_STYLES = {
  'Job Code':            { fill: '0070C0', font: 'FFFFFF' },
  'Site ID':             { fill: '0070C0', font: 'FFFFFF' },
  'Area':                { fill: '0070C0', font: 'FFFFFF' },
  'VF Owner':            { fill: '0070C0', font: 'FFFFFF' },
  'Contractor':          { fill: '0070C0', font: 'FFFFFF' },
  'Installation Status': { fill: '0070C0', font: 'FFFFFF' },
  'Installation Date':   { fill: '0070C0', font: 'FFFFFF' },
  'Migration Status':    { fill: '0070C0', font: 'FFFFFF' },
  'Migration Date':      { fill: '0070C0', font: 'FFFFFF' },
  'Acceptance Status':   { fill: '92D050', font: '000000' },
  'Certificate':         { fill: '92D050', font: '000000' },
  'FAC Date':            { fill: '92D050', font: '000000' },
  'Line Item':           { fill: 'FFC000', font: '000000' },
  'Price':               { fill: 'FFC000', font: '000000' },
  'Invoice Amount':      { fill: 'FFC000', font: '000000' },
};

const FINANCIAL_LABELS = new Set(['Price', 'Invoice Amount']);
const EGP_FMT  = '#,##0 "EGP"';
const THIN_BORDER   = { style: 'thin',   color: { argb: 'FF000000' } };
const DOUBLE_BORDER = { style: 'double', color: { argb: 'FF000000' } };
const ALL_BORDERS        = { top: THIN_BORDER,   left: THIN_BORDER,   bottom: THIN_BORDER,   right: THIN_BORDER   };
const ALL_DOUBLE_BORDERS = { top: DOUBLE_BORDER, left: DOUBLE_BORDER, bottom: DOUBLE_BORDER, right: DOUBLE_BORDER };

async function exportToExcel(result, originalFileName) {
  const outputHeaders = OUTPUT_COLUMNS.map(c => c.label);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Invoice Prep Output');

  ws.columns = outputHeaders.map(h => {
    const maxLen = Math.max(h.length, ...result.combined.map(r => String(r[h] ?? '').length));
    return { width: Math.min(maxLen + 4, 42) };
  });

  const TOTAL_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00B050' } };

  ws.mergeCells(1, 5, 1, 7);
  const labelCell = ws.getCell(1, 5);
  labelCell.value     = 'Total Invoice Amount';
  labelCell.fill      = TOTAL_FILL;
  labelCell.font      = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
  labelCell.alignment = { horizontal: 'center', vertical: 'middle' };
  labelCell.border    = ALL_DOUBLE_BORDERS;

  ws.mergeCells(1, 8, 1, 10);
  const amountCell = ws.getCell(1, 8);
  amountCell.value     = result.totalAmount;
  amountCell.numFmt    = EGP_FMT;
  amountCell.fill      = TOTAL_FILL;
  amountCell.font      = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
  amountCell.alignment = { horizontal: 'center', vertical: 'middle' };
  amountCell.border    = ALL_DOUBLE_BORDERS;

  ws.getRow(1).height = 22;

  outputHeaders.forEach((h, i) => {
    const c = ws.getCell(3, i + 1);
    c.value = h;
    const style = HEADER_STYLES[h] || { fill: '4472C4', font: 'FFFFFF' };
    c.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + style.fill } };
    c.font      = { bold: true, color: { argb: 'FF' + style.font }, size: 11 };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    c.border    = ALL_DOUBLE_BORDERS;
  });
  ws.getRow(3).height = 22;

  result.combined.forEach((row, rowIdx) => {
    outputHeaders.forEach((h, colIdx) => {
      const c   = ws.getCell(4 + rowIdx, colIdx + 1);
      const val = row[h] ?? '';
      c.value  = val;
      if (FINANCIAL_LABELS.has(h) && typeof val === 'number') {
        c.numFmt = EGP_FMT;
      } else if (val instanceof Date) {
        c.numFmt = 'dd-mmm-yy';
      }
      c.border = ALL_BORDERS;
    });
  });

  const buffer = await wb.xlsx.writeBuffer();
  const blob   = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a   = document.createElement('a');
  a.href    = url;
  a.download = 'Invoice Request Track.xlsx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ---------------------------------------------------------------------------
// UI Logic
// ---------------------------------------------------------------------------
let currentResult = null;
let currentParsed = null;
let currentFileName = '';

const batchWrap     = document.getElementById('pocBatchWrap');
const batchSelect   = document.getElementById('pocBatchSelect');

const dropZone      = document.getElementById('dropZone');
const fileInput     = document.getElementById('fileInput');
const fileInfo      = document.getElementById('fileInfo');
const fileNameEl    = document.getElementById('fileName');
const clearFileBtn  = document.getElementById('clearFile');
const warningsEl    = document.getElementById('warnings');
const progressWrap  = document.getElementById('progressWrap');
const progressFill  = document.getElementById('progressFill');
const progressLabel = document.getElementById('progressLabel');
const resultsSection = document.getElementById('resultsSection');
const step1CountEl  = document.getElementById('step1Count');
const step2CountEl  = document.getElementById('step2Count');
const totalCountEl  = document.getElementById('totalCount');
const step1AmountEl = document.getElementById('step1Amount');
const step2AmountEl = document.getElementById('step2Amount');
const totalAmountEl = document.getElementById('totalAmount');
const downloadBtn   = document.getElementById('downloadBtn');

function setFile(file) {
  if (!file) return;
  const allowed = ['.xlsx', '.xls', '.xlsm', '.csv'];
  const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
  if (!allowed.includes(ext)) {
    showWarning([`Unsupported file type "${ext}". Please upload an Excel (.xlsx, .xls) or CSV file.`]);
    return;
  }
  currentFileName = file.name;
  fileNameEl.textContent = file.name;
  fileInfo.hidden = false;
  dropZone.hidden = true;
  currentParsed = null;
  batchWrap.hidden = true;
  clearResults();
  processFile(file);
}

function clearFile() {
  currentFileName = '';
  currentParsed = null;
  fileInput.value = '';
  fileInfo.hidden = true;
  dropZone.hidden = false;
  batchWrap.hidden = true;
  batchSelect.innerHTML = '';
  clearResults();
  warningsEl.hidden = true;
}

function populateBatches(batches) {
  const opt = (b) => {
    const parts = [];
    if (b.ins) parts.push(`${b.ins} ins`);
    if (b.mig) parts.push(`${b.mig} mig`);
    const o = document.createElement('option');
    o.value = b.key;
    o.textContent = `${b.label}  (${parts.join(' / ')})`;
    return o;
  };
  batchSelect.innerHTML = '<option value="">— Select invoice batch —</option>';
  const text  = batches.filter(b => !b.isDate);
  const dates = batches.filter(b => b.isDate);
  if (text.length) {
    const g = document.createElement('optgroup');
    g.label = 'Batch labels';
    text.forEach(b => g.appendChild(opt(b)));
    batchSelect.appendChild(g);
  }
  if (dates.length) {
    const g = document.createElement('optgroup');
    g.label = 'Invoiced dates';
    dates.forEach(b => g.appendChild(opt(b)));
    batchSelect.appendChild(g);
  }
  batchWrap.hidden = false;
}

function runBatch() {
  clearResults();
  if (!currentParsed || !batchSelect.value) return;
  try {
    currentResult = analyzeBatch(currentParsed, batchSelect.value);
    renderResults(currentResult);
    if (currentResult.warnings.length > 0) showWarning(currentResult.warnings);
    else warningsEl.hidden = true;
  } catch (err) {
    showWarning([`Error: ${err.message}`]);
  }
}

batchSelect.addEventListener('change', runBatch);

function clearResults() {
  currentResult = null;
  resultsSection.hidden = true;
  progressWrap.hidden = true;
  progressFill.style.width = '0%';
}

function setProgress(pct, label) {
  progressWrap.hidden = false;
  progressFill.style.width = pct + '%';
  progressLabel.textContent = label;
}

dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file) setFile(file);
});
dropZone.addEventListener('click', e => { if (e.target.tagName === 'LABEL') return; fileInput.click(); });

fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) setFile(fileInput.files[0]);
});

clearFileBtn.addEventListener('click', e => { e.stopPropagation(); clearFile(); });

function processFile(file) {
  warningsEl.hidden = true;
  setProgress(10, 'Reading file…');

  const reader = new FileReader();
  reader.onload = e => {
    try {
      setProgress(40, 'Parsing workbook…');
      const data = new Uint8Array(e.target.result);

      setProgress(70, 'Reading invoice batches…');
      currentParsed = parseWorkbook(data);

      setProgress(100, 'Done! Select an invoice batch.');
      setTimeout(() => { progressWrap.hidden = true; }, 1200);

      if (currentParsed.batches.length === 0) {
        showWarning(['No values found in the Installation / Migration Invoicing Date columns. Fill them with the batch name (e.g. "new oct") first.']);
        return;
      }
      populateBatches(currentParsed.batches);
      batchSelect.focus();
    } catch (err) {
      progressWrap.hidden = true;
      showWarning([`Error: ${err.message}`]);
    }
  };
  reader.onerror = () => {
    progressWrap.hidden = true;
    showWarning(['Failed to read the file. Please try again.']);
  };
  reader.readAsArrayBuffer(file);
}

function formatEGP(val) {
  return 'EGP ' + val.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

function renderResults(result) {
  step1CountEl.textContent  = result.step1Count;
  step2CountEl.textContent  = result.step2Count;
  totalCountEl.textContent  = result.combined.length;
  step1AmountEl.textContent = formatEGP(result.step1Amount);
  step2AmountEl.textContent = formatEGP(result.step2Amount);
  totalAmountEl.textContent = formatEGP(result.totalAmount);

  resultsSection.hidden = false;
  resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

downloadBtn.addEventListener('click', () => {
  if (!currentResult) return;
  exportToExcel(currentResult, currentFileName).catch(err => {
    showWarning([`Export failed: ${err.message}`]);
  });
});

function showWarning(messages) {
  warningsEl.hidden = false;
  warningsEl.innerHTML = `<strong>Notice</strong><ul>${messages.map(m => `<li>${m}</li>`).join('')}</ul>`;
}

})(); // end IIFE
