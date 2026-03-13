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
  installInvoicingDate:  ['installation invoicing date', 'install invoicing date',
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
};

const OUTPUT_COLUMNS = [
  { label: 'Job Code',            key: 'jobCode'           },
  { label: 'Site ID',             key: 'siteId'            },
  { label: 'Area',                key: 'area'              },
  { label: 'VF Owner',            key: 'vfOwner'           },
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
function findColumn(headers, patterns) {
  const lower = headers.map(h => String(h ?? '').toLowerCase().trim());
  for (const pattern of patterns) {
    const idx = lower.findIndex(h => h.includes(pattern));
    if (idx !== -1) return idx;
  }
  return -1;
}

function buildColumnMap(headers) {
  const map = {};
  for (const [key, patterns] of Object.entries(COL_PATTERNS)) {
    map[key] = findColumn(headers, patterns);
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

function eqCI(val, target) {
  return String(val ?? '').trim().toLowerCase() === target.toLowerCase();
}

function notEqCI(val, target) {
  return String(val ?? '').trim().toLowerCase() !== target.toLowerCase();
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
// Core processing
// ---------------------------------------------------------------------------
function processExcel(fileData) {
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

  const step1 = dataRows.filter(row => {
    return eqCI(cell(row, colMap.installationStatus), 'done') &&
           isBlank(cell(row, colMap.installInvoicingDate)) &&
           notEqCI(cell(row, colMap.lineItem), 'POC2 Migration');
  });

  const step2 = dataRows.filter(row => {
    return eqCI(cell(row, colMap.migrationStatus), 'done') &&
           eqCI(cell(row, colMap.acceptanceStatus), 'fac') &&
           isBlank(cell(row, colMap.migInvoicingDate)) &&
           notEqCI(cell(row, colMap.lineItem), 'POC2 Migration');
  });

  function extractRow(row, stepLabel) {
    const out = { _step: stepLabel };
    for (const { label, key } of OUTPUT_COLUMNS) {
      if (key === 'invoiceAmount') {
        const rawTotal = parseFloat(String(cell(row, colMap.totalAmount) ?? '').replace(/,/g, '')) || 0;
        out[label] = rawTotal / 2;
      } else {
        let v = cell(row, colMap[key]);
        if (v instanceof Date) {
          v = formatDate(v);
        }
        out[label] = v;
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
  const criticalKeys = ['installationStatus', 'lineItem'];
  for (const key of criticalKeys) {
    if (colMap[key] === -1) {
      const patterns = COL_PATTERNS[key];
      warnings.push(`Column not found: expected something like "${patterns[0]}". Check your header row.`);
    }
  }
  const computedKeys = new Set(['invoiceAmount']);
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
  a.download = originalFileName.replace(/\.[^.]+$/, '') + '_Invoice_Output.xlsx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ---------------------------------------------------------------------------
// UI Logic
// ---------------------------------------------------------------------------
let currentResult = null;
let currentFileName = '';

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
  clearResults();
  processFile(file);
}

function clearFile() {
  currentFileName = '';
  fileInput.value = '';
  fileInfo.hidden = true;
  dropZone.hidden = false;
  clearResults();
  warningsEl.hidden = true;
}

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

      setProgress(65, 'Applying filters…');
      currentResult = processExcel(data);

      setProgress(90, 'Building results…');
      renderResults(currentResult);

      setProgress(100, 'Done!');
      setTimeout(() => { progressWrap.hidden = true; }, 1200);

      if (currentResult.warnings.length > 0) {
        showWarning(currentResult.warnings);
      }
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
