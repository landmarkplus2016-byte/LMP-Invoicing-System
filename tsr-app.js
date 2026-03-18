// =============================================================================
// LMP Invoicing System — TSR Submission Checker logic
// Wrapped in IIFE to avoid naming conflicts with POC app.
// =============================================================================
(function () {

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------
let trackingWorkbook = null;
let tsrWorkbook = null;
let analysisResults = [];
let allExportItems = [];
let moneyCanSubmit = 0;
let moneyPending   = 0;
let moneyNeedPo    = 0;

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------
const distanceMultipliers = {
  '0km - 100km':   1.0,
  '0Km - 100Km':   1.0,
  '100km - 400km': 1.1,
  '100Km - 400Km': 1.1,
  '400km - 800km': 1.2,
  '400Km - 800Km': 1.2,
  '> 800km':       1.25,
  '> 800Km':       1.25
};

// ---------------------------------------------------------------------------
// File Picking — Tracking File
// ---------------------------------------------------------------------------
document.getElementById('btn-pick-tracking').addEventListener('click', () => {
  document.getElementById('tracking-input').click();
});

document.getElementById('tracking-input').addEventListener('change', async (e) => {
  try {
    clearError();
    const file = e.target.files[0];
    if (!file) return;
    showFileProgress('tracking', true);
    const data = await file.arrayBuffer();
    trackingWorkbook = XLSX.read(data, { type: 'array', cellDates: true });
    document.getElementById('tracking-filename').textContent = file.name;
    document.getElementById('card-tracking').classList.add('loaded');
    showFileProgress('tracking', false);
    checkReadyToAnalyze();
  } catch (err) {
    showFileProgress('tracking', false);
    showError('Failed to open tracking file: ' + err.message);
  }
});

// ---------------------------------------------------------------------------
// File Picking — TSR File
// ---------------------------------------------------------------------------
document.getElementById('btn-pick-tsr').addEventListener('click', () => {
  document.getElementById('tsr-input').click();
});

document.getElementById('tsr-input').addEventListener('change', async (e) => {
  try {
    clearError();
    const file = e.target.files[0];
    if (!file) return;
    showFileProgress('tsr', true);
    const data = await file.arrayBuffer();
    tsrWorkbook = XLSX.read(data, { type: 'array', cellDates: true });
    document.getElementById('tsr-filename').textContent = file.name;
    document.getElementById('card-tsr').classList.add('loaded');
    showFileProgress('tsr', false);
    checkReadyToAnalyze();
  } catch (err) {
    showFileProgress('tsr', false);
    showError('Failed to open TSR file: ' + err.message);
  }
});

function checkReadyToAnalyze() {
  if (trackingWorkbook && tsrWorkbook) triggerAnalysis();
}

// ---------------------------------------------------------------------------
// Analysis
// ---------------------------------------------------------------------------
async function triggerAnalysis() {
  try {
    clearError();
    showLoading(true);
    document.getElementById('step-analysis').style.display = 'none';
    document.getElementById('step-results').style.display = 'none';
    document.getElementById('step-export').style.display = 'none';

    await new Promise(resolve => setTimeout(resolve, 50));

    analysisResults = runAnalysis();
    renderResults(analysisResults);

    showLoading(false);
    document.getElementById('step-analysis').style.display = 'block';
    document.getElementById('step-results').style.display = 'block';
    document.getElementById('step-export').style.display = 'block';
    document.getElementById('step-analysis').scrollIntoView({ behavior: 'smooth' });
  } catch (err) {
    showLoading(false);
    showError(err.message || 'An unexpected error occurred during analysis.');
  }
}

// ---------------------------------------------------------------------------
// Core Analysis Logic
// ---------------------------------------------------------------------------
function runAnalysis() {

  const trackingSheet = trackingWorkbook.Sheets['Invoicing Track'];
  if (!trackingSheet) {
    throw new Error('Wrong file loaded in Task Tracking. Expected sheet: Invoicing Track');
  }

  const trackingRows = XLSX.utils.sheet_to_json(trackingSheet, { header: 1, defval: null });

  const dataRows  = trackingRows.slice(4);
  const headerRow = trackingRows[3] ?? [];

  let idColIndex               = -1;
  let jobCodeColIndex          = -1;
  let logicalSiteColIndex      = -1;
  let acceptanceStatusColIndex = -1;
  let newTotalColIndex         = -1;
  let vfTaskOwnerColIndex      = -1;
  let vendorColIndex           = -1;
  let siteOptionColIndex       = -1;
  let facingColIndex           = -1;
  let taskDateColIndex         = -1;
  let prqColIndex              = -1;
  let certificateColIndex      = -1;
  let distanceColIndex         = -1;
  let absQtyColIndex           = -1;
  let lineItemColIndex         = -1;
  let facDateColIndex          = -1;
  let acceptanceWeekColIndex   = -1;
  let poStatusColIndex         = -1;

  for (let c = 0; c < headerRow.length; c++) {
    const h = (headerRow[c] ?? '').toString().trim().toLowerCase();
    // Use first-match for every column so a later unrelated column can't override the correct one
    if (idColIndex               < 0 && (h === 'id#' || h === 'id #'))                              idColIndex               = c;
    if (jobCodeColIndex          < 0 && h.includes('job code'))                                      jobCodeColIndex          = c;
    if (logicalSiteColIndex      < 0 && h.includes('logical site'))                                  logicalSiteColIndex      = c;
    if (acceptanceStatusColIndex < 0 && h.includes('acceptance status'))                             acceptanceStatusColIndex = c;
    if (newTotalColIndex         < 0 && h.includes('new total'))                                     newTotalColIndex         = c;
    if (vfTaskOwnerColIndex      < 0 && h.includes('task owner'))                                    vfTaskOwnerColIndex      = c;
    if (vendorColIndex           < 0 && (h === 'vendor' || h.includes('vendor')))                   vendorColIndex           = c;
    if (siteOptionColIndex       < 0 && h.includes('site option'))                                   siteOptionColIndex       = c;
    if (facingColIndex           < 0 && (h === 'facing' || h.includes('facing')))                   facingColIndex           = c;
    if (taskDateColIndex         < 0 && h.includes('task date'))                                     taskDateColIndex         = c;
    if (prqColIndex              < 0 && (h === 'prq' || h.includes('prq')))                         prqColIndex              = c;
    if (certificateColIndex      < 0 && h.includes('certificate'))                                   certificateColIndex      = c;
    if (distanceColIndex         < 0 && h.includes('distance'))                                      distanceColIndex         = c;
    if (absQtyColIndex           < 0 && h.includes('absolute') && (h.includes('qty') || h.includes('quant'))) absQtyColIndex = c;
    if (lineItemColIndex         < 0 && (h === 'line item' || h === 'line items'))                  lineItemColIndex         = c;
    if (facDateColIndex          < 0 && (h.includes('fac date') || h === 'fac'))                    facDateColIndex          = c;
    if (acceptanceWeekColIndex   < 0 && h.includes('acceptance week'))                               acceptanceWeekColIndex   = c;
    if (poStatusColIndex         < 0 && (h.includes('po status') || h.includes('po #')
                                      || h.includes('purchase order')))                              poStatusColIndex         = c;
  }

  // Validate all critical columns were found
  const missingCols = [];
  if (jobCodeColIndex          < 0) missingCols.push('"Job Code"');
  if (logicalSiteColIndex      < 0) missingCols.push('"Logical Site ID"');
  if (lineItemColIndex         < 0) missingCols.push('"Line Item"');
  if (facDateColIndex          < 0) missingCols.push('"FAC Date"');
  if (acceptanceWeekColIndex   < 0) missingCols.push('"Acceptance Week"');
  if (newTotalColIndex         < 0) missingCols.push('"New Total"');
  if (absQtyColIndex           < 0) missingCols.push('"Absolute Quantity"');

  if (missingCols.length > 0) {
    const foundHeaders = headerRow
      .map((h, i) => (h ? 'col' + i + ':"' + h + '"' : null))
      .filter(Boolean).join(' | ');
    throw new Error(
      'Missing required column(s): ' + missingCols.join(', ') + '. ' +
      'All headers found in row 4: [ ' + foundHeaders + ' ]'
    );
  }

  const comboRows = new Map();
  dataRows.forEach((row, idx) => {
    if (row[poStatusColIndex] != null && row[poStatusColIndex] !== '') return;
    const combo = comboKey(row, jobCodeColIndex, logicalSiteColIndex);
    if (combo === '|') return;
    if (!comboRows.has(combo)) comboRows.set(combo, []);
    comboRows.get(combo).push({ row, rowIndex: idx });
  });

  const activeCombos = new Map();
  comboRows.forEach((entries, combo) => {
    if (entries.some(({ row }) => row[facDateColIndex] != null && row[facDateColIndex] !== '')) {
      activeCombos.set(combo, entries);
    }
  });

  if (activeCombos.size === 0) {
    throw new Error('No sites found with FAC Date. Check that FAC Date is filled and PO Status is empty.');
  }

  const groups = new Map();
  activeCombos.forEach((entries) => {
    entries.forEach(({ row, rowIndex }) => {
      if (row[facDateColIndex] == null || row[facDateColIndex] === '') return;
      const lineItem  = (row[lineItemColIndex] ?? '').toString().trim();
      if (!lineItem) return;
      const distance  = (row[distanceColIndex] ?? '').toString().trim();
      const absQty    = Number(row[absQtyColIndex] ?? 1);
      const actualQty = absQty * (distanceMultipliers[distance] ?? 1.0);
      const excelRow  = rowIndex + 5;
      if (!groups.has(lineItem)) {
        groups.set(lineItem, {
          totalQty: 0, excelRowNumbers: [], individualQtys: [],
          facDates: [], acceptanceWeeks: []
        });
      }
      const g = groups.get(lineItem);
      g.totalQty += actualQty;
      g.excelRowNumbers.push(excelRow);
      g.individualQtys.push(actualQty);
      g.facDates.push(row[facDateColIndex]);
      g.acceptanceWeeks.push(row[acceptanceWeekColIndex]);
    });
  });

  if (groups.size === 0) {
    const sample = [];
    activeCombos.forEach((entries) => {
      if (sample.length >= 3) return;
      entries.forEach(({ row }) => {
        if (sample.length >= 3) return;
        if (row[facDateColIndex] == null || row[facDateColIndex] === '') return;
        sample.push(
          'lineItem col' + lineItemColIndex + '="' + (row[lineItemColIndex] ?? 'NULL') + '"' +
          ' facDate col' + facDateColIndex + '="' + (row[facDateColIndex] ?? 'NULL') + '"'
        );
      });
    });
    throw new Error(
      'No line items could be read for FAC-dated rows. ' +
      'lineItemCol=' + lineItemColIndex + ', facDateCol=' + facDateColIndex + '. ' +
      'Sample FAC-dated rows: [ ' + (sample.join(' | ') || 'none') + ' ]'
    );
  }

  const tsrSheet = tsrWorkbook.Sheets['Request Form - VF'];
  if (!tsrSheet) {
    throw new Error('Wrong file loaded in TSR. Expected sheet: Request Form - VF');
  }

  const tsrRows = XLSX.utils.sheet_to_json(tsrSheet, { header: 1, defval: null });

  let headerRowIdx = -1;
  for (let i = 0; i < tsrRows.length; i++) {
    const row = tsrRows[i];
    for (let c = 0; c < row.length; c++) {
      if (row[c] && row[c].toString().toLowerCase().includes('item description')) {
        headerRowIdx = i;
        break;
      }
    }
    if (headerRowIdx >= 0) break;
  }

  // Detect TSR column positions from the header row
  let tsrItemColIndex      = -1;
  let tsrUnitPriceColIndex = -1;
  let tsrRemainingColIndex = -1;

  if (headerRowIdx >= 0) {
    const tsrHeader = tsrRows[headerRowIdx];
    for (let c = 0; c < tsrHeader.length; c++) {
      const h = (tsrHeader[c] ?? '').toString().trim().toLowerCase();
      if (tsrItemColIndex      < 0 && h.includes('item description'))                       tsrItemColIndex      = c;
      if (tsrUnitPriceColIndex < 0 && h.includes('unit price'))                             tsrUnitPriceColIndex = c;
      if (tsrRemainingColIndex < 0 && h.includes('remaining') && !h.includes('after'))     tsrRemainingColIndex = c;
    }
  }

  // Fall back to original hardcoded positions if headers not found
  if (tsrItemColIndex      < 0) tsrItemColIndex      = 6;
  if (tsrUnitPriceColIndex < 0) tsrUnitPriceColIndex = 12;
  if (tsrRemainingColIndex < 0) tsrRemainingColIndex = 50;

  const tsrMap = new Map();
  if (headerRowIdx >= 0) {
    for (let i = headerRowIdx + 1; i < tsrRows.length; i++) {
      const itemDesc = tsrRows[i][tsrItemColIndex];
      if (itemDesc == null || itemDesc === '') continue;
      const remaining = tsrRows[i][tsrRemainingColIndex];
      const unitPrice = tsrRows[i][tsrUnitPriceColIndex];
      if (remaining == null || remaining === '' || Number(remaining) <= 0) continue;
      if (unitPrice == null || unitPrice === '') continue;
      tsrMap.set(itemDesc.toString().trim(), {
        remaining: Number(remaining),
        unitPrice: Number(unitPrice)
      });
    }
  }

  const results = [];

  for (const [lineItem, g] of groups) {
    let tsrEntry = tsrMap.get(lineItem);
    if (tsrEntry === undefined) {
      for (const [key, val] of tsrMap) {
        if (key.includes(lineItem) || lineItem.includes(key)) {
          tsrEntry = val;
          break;
        }
      }
    }

    const trackingQty  = Math.round(g.totalQty * 100) / 100;
    const tsrRemaining = tsrEntry !== undefined ? tsrEntry.remaining : null;
    const tsrUnitPrice = tsrEntry !== undefined ? tsrEntry.unitPrice : null;
    const difference   = tsrRemaining !== null ? tsrRemaining - trackingQty : null;

    let status;
    if (tsrRemaining === null)             status = 'NOT_FOUND';
    else if (trackingQty <= tsrRemaining)  status = 'OK';
    else                                   status = 'EXCEEDS';

    results.push({
      lineItem, trackingQty, tsrRemaining, tsrUnitPrice, difference, status,
      excelRowNumbers: g.excelRowNumbers,
      individualQtys:  g.individualQtys,
      facDates:        g.facDates,
      acceptanceWeeks: g.acceptanceWeeks
    });
  }

  const needPoCombos    = new Set();
  const pendingCombos   = new Set();
  const canSubmitCombos = new Set();

  const canSubmitCandidates = [];
  activeCombos.forEach((entries, combo) => {
    const facEntries           = entries.filter(({ row }) => row[facDateColIndex] != null && row[facDateColIndex] !== '');
    const allFac               = entries.every(({ row }) => row[facDateColIndex] != null && row[facDateColIndex] !== '');
    const allFacHaveAcceptWeek = facEntries.every(({ row }) => row[acceptanceWeekColIndex] != null && row[acceptanceWeekColIndex] !== '');
    if (!allFac || !allFacHaveAcceptWeek) {
      pendingCombos.add(combo);
    } else {
      canSubmitCandidates.push(combo);
    }
  });

  const comboFirstRow    = new Map();
  const comboLineItemQty = new Map();

  for (const combo of canSubmitCandidates) {
    const entries  = activeCombos.get(combo);
    let minRow     = Infinity;
    const liQtyMap = new Map();

    for (const { row, rowIndex } of entries) {
      const excelRow = rowIndex + 5;
      if (excelRow < minRow) minRow = excelRow;
      if (row[facDateColIndex] == null || row[facDateColIndex] === '') continue;
      const li = (row[lineItemColIndex] ?? '').toString().trim();
      if (!li) continue;
      const distance  = (row[distanceColIndex] ?? '').toString().trim();
      const absQty    = Number(row[absQtyColIndex] ?? 1);
      const actualQty = absQty * (distanceMultipliers[distance] ?? 1.0);
      liQtyMap.set(li, (liQtyMap.get(li) ?? 0) + actualQty);
    }

    comboFirstRow.set(combo, minRow);
    comboLineItemQty.set(combo, liQtyMap);
  }

  canSubmitCandidates.sort((a, b) => comboFirstRow.get(a) - comboFirstRow.get(b));

  const tsrAvailable = new Map();
  for (const r of results) {
    if (r.tsrRemaining !== null) tsrAvailable.set(r.lineItem, r.tsrRemaining);
  }

  for (const combo of canSubmitCandidates) {
    const liQtyMap = comboLineItemQty.get(combo);
    let canFit = true;

    for (const [li, qty] of liQtyMap) {
      const available = tsrAvailable.get(li);
      if (available === undefined || qty > available + 0.005) { canFit = false; break; }
    }

    if (canFit && liQtyMap.size > 0) {
      for (const [li, qty] of liQtyMap) tsrAvailable.set(li, tsrAvailable.get(li) - qty);
      canSubmitCombos.add(combo);
    } else {
      needPoCombos.add(combo);
    }
  }

  for (const r of results) {
    if (r.tsrRemaining === null) {
      r.tsrAfterSubmit = null;
      r.tsrAfterAmount = null;
      r.tsrEgpUnitRate = null;
      continue;
    }
    r.tsrAfterSubmit = Math.round((tsrAvailable.get(r.lineItem) ?? r.tsrRemaining) * 100) / 100;
    r.tsrEgpUnitRate = r.tsrUnitPrice !== null ? Math.round(r.tsrUnitPrice * 100) / 100 : null;
    r.tsrAfterAmount = (r.tsrUnitPrice !== null)
      ? Math.round(r.tsrAfterSubmit * r.tsrUnitPrice * 100) / 100
      : null;
  }

  function makeExportRow(row, comment) {
    const distance  = (row[distanceColIndex] ?? '').toString().trim();
    const absQty    = Number(row[absQtyColIndex] ?? 1);
    const actualQty = absQty * (distanceMultipliers[distance] ?? 1.0);
    const rawAcc    = acceptanceStatusColIndex >= 0 ? row[acceptanceStatusColIndex] : null;
    return {
      vfTaskOwner:      colVal(row, vfTaskOwnerColIndex),
      vendor:           colVal(row, vendorColIndex),
      logicalSiteId:    colVal(row, logicalSiteColIndex),
      siteOption:       colVal(row, siteOptionColIndex),
      facing:           colVal(row, facingColIndex),
      taskDate:         taskDateColIndex >= 0 ? row[taskDateColIndex] : null,
      lineItem:         (row[lineItemColIndex] ?? '').toString().trim(),
      absQty,
      prq:              colVal(row, prqColIndex),
      certificate:      colVal(row, certificateColIndex),
      acceptanceStatus: rawAcc != null ? String(rawAcc) : '',
      actualQty,
      newTotal:         newTotalColIndex >= 0 ? row[newTotalColIndex] : null,
      idNum:            colVal(row, idColIndex),
      jobCode:          colVal(row, jobCodeColIndex),
      comment
    };
  }

  allExportItems = [];

  activeCombos.forEach((entries, combo) => {
    if (!canSubmitCombos.has(combo)) return;
    entries.forEach(({ row }) => {
      if ((row[lineItemColIndex] ?? '').toString().trim() === '') return;
      allExportItems.push(makeExportRow(row, 'Can Submit'));
    });
  });

  activeCombos.forEach((entries, combo) => {
    if (!pendingCombos.has(combo)) return;
    entries.forEach(({ row }) => {
      if ((row[lineItemColIndex] ?? '').toString().trim() === '') return;
      allExportItems.push(makeExportRow(row, 'Pending'));
    });
  });

  activeCombos.forEach((entries, combo) => {
    if (!needPoCombos.has(combo)) return;
    entries.forEach(({ row }) => {
      if ((row[lineItemColIndex] ?? '').toString().trim() === '') return;
      allExportItems.push(makeExportRow(row, 'Need PO'));
    });
  });

  allExportItems.sort((a, b) => {
    const s = a.logicalSiteId.localeCompare(b.logicalSiteId);
    return s !== 0 ? s : a.jobCode.localeCompare(b.jobCode);
  });

  moneyCanSubmit = 0;
  moneyPending   = 0;
  moneyNeedPo    = 0;
  activeCombos.forEach((entries, combo) => {
    entries.forEach(({ row }) => {
      const val    = newTotalColIndex >= 0 ? row[newTotalColIndex] : null;
      const amount = (val != null && val !== '') ? (Number(val) || 0) : 0;
      if (amount === 0) return;
      if (canSubmitCombos.has(combo))    moneyCanSubmit += amount;
      else if (pendingCombos.has(combo)) moneyPending   += amount;
      else if (needPoCombos.has(combo))  moneyNeedPo    += amount;
    });
  });

  return results;
}

// ---------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------
function colVal(row, idx) {
  return idx >= 0 ? (row[idx] ?? '').toString().trim() : '';
}

function comboKey(row, jobCodeIdx, logicalSiteIdx) {
  return colVal(row, jobCodeIdx) + '|' + colVal(row, logicalSiteIdx);
}

// ---------------------------------------------------------------------------
// Render Results
// ---------------------------------------------------------------------------
function renderResults(results) {
  function fmtEGP(n) {
    return 'EGP ' + n.toLocaleString('en-EG', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  const totalRowCount = results.reduce((s, r) => s + r.excelRowNumbers.length, 0);
  const lineItemCount = results.length;
  document.getElementById('summary').textContent =
    totalRowCount + ' row' + (totalRowCount !== 1 ? 's' : '') +
    ' matched across ' + lineItemCount + ' line item' + (lineItemCount !== 1 ? 's' : '');

  document.getElementById('money-can-submit').innerHTML =
    '<span class="money-label">\u2705 Can Submit</span><span class="money-value">' + fmtEGP(moneyCanSubmit) + '</span>';
  document.getElementById('money-pending').innerHTML =
    '<span class="money-label">\u23f3 Pending</span><span class="money-value">' + fmtEGP(moneyPending) + '</span>';
  document.getElementById('money-need-po').innerHTML =
    '<span class="money-label">\uD83D\uDCCB Need PO</span><span class="money-value">' + fmtEGP(moneyNeedPo) + '</span>';

  const tsrRows2 = results.filter(r => r.tsrAfterSubmit !== null && r.tsrAfterSubmit > 0);

  let grandTotalQty    = 0;
  let grandTotalAmount = 0;
  let hasAmounts       = false;

  let rowsHtml = '';
  for (const r of tsrRows2) {
    const unitPrice = r.tsrEgpUnitRate !== null ? fmtEGP(r.tsrEgpUnitRate) : '\u2014';
    const afterQty  = r.tsrAfterSubmit.toFixed(2);
    const afterAmt  = r.tsrAfterAmount !== null ? fmtEGP(r.tsrAfterAmount) : '\u2014';
    grandTotalQty += r.tsrAfterSubmit;
    if (r.tsrAfterAmount !== null) { grandTotalAmount += r.tsrAfterAmount; hasAmounts = true; }
    rowsHtml +=
      '<tr>' +
        '<td>' + escapeHtml(r.lineItem) + '</td>' +
        '<td class="num">' + unitPrice + '</td>' +
        '<td class="num">' + afterQty + '</td>' +
        '<td class="num">' + afterAmt + '</td>' +
      '</tr>';
  }

  const grandAmtStr = hasAmounts ? fmtEGP(grandTotalAmount) : '\u2014';
  rowsHtml +=
    '<tr class="totals-row">' +
      '<td><strong>TOTAL REMAINING IN TSR</strong></td>' +
      '<td></td>' +
      '<td class="num"><strong>' + grandTotalQty.toFixed(2) + '</strong></td>' +
      '<td class="num"><strong>' + grandAmtStr + '</strong></td>' +
    '</tr>';

  document.getElementById('tsr-remaining-body').innerHTML = rowsHtml;

  document.getElementById('btn-export').disabled = false;
  document.getElementById('export-status').textContent = '';
  document.getElementById('export-status').className = '';
}

// ---------------------------------------------------------------------------
// Export to Excel
// ---------------------------------------------------------------------------
document.getElementById('btn-export').addEventListener('click', () => {
  try {
    exportToExcel();
  } catch (err) {
    showExportStatus('\u274c Export failed: ' + err.message, 'error');
  }
});

function exportToExcel() {
  const HEADERS = [
    'VF Task Owner', 'Vendor', 'Logical Site ID', 'Site Option', 'Facing',
    'Task Date', 'Line Item', 'Absolute Quantity', 'PRQ', 'Certificate #',
    'Acceptance Status', 'Actual Quantity', 'New Total Price', 'ID#', 'Job Code', 'Comment'
  ];
  const COL_WIDTHS = [
    { wch: 18 }, { wch: 14 }, { wch: 18 }, { wch: 14 }, { wch: 12 },
    { wch: 14 }, { wch: 60 }, { wch: 18 }, { wch: 10 }, { wch: 16 },
    { wch: 20 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 14 }, { wch: 14 }
  ];

  const data = [HEADERS];
  for (const r of allExportItems) {
    data.push([
      r.vfTaskOwner,
      r.vendor,
      r.logicalSiteId,
      r.siteOption,
      r.facing,
      formatDateValue(r.taskDate),
      r.lineItem,
      r.absQty,
      r.prq,
      r.certificate,
      r.acceptanceStatus,
      Number(r.actualQty.toFixed(2)),
      r.newTotal != null ? Number(r.newTotal) : '',
      r.idNum,
      r.jobCode,
      r.comment
    ]);
  }

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = COL_WIDTHS;

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'TSR Analysis');

  const today = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, 'TSR_Submission_Results_' + today + '.xlsx');

  const canCount    = allExportItems.filter(r => r.comment === 'Can Submit').length;
  const pendCount   = allExportItems.filter(r => r.comment === 'Pending').length;
  const needPoCount = allExportItems.filter(r => r.comment === 'Need PO').length;
  showExportStatus(
    '\u2705 Export downloaded! ' +
    canCount + ' can submit, ' + pendCount + ' pending, ' + needPoCount + ' need PO.',
    'success'
  );
}

// ---------------------------------------------------------------------------
// Utility Helpers
// ---------------------------------------------------------------------------
function formatDateValue(val) {
  if (val == null || val === '') return '';
  if (val instanceof Date) return val.toLocaleDateString('en-AU');
  return String(val);
}

function showError(msg) {
  const el = document.getElementById('error-message');
  el.textContent = msg;
  el.style.display = 'block';
  el.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

function clearError() {
  const el = document.getElementById('error-message');
  el.textContent = '';
  el.style.display = 'none';
}

function showExportStatus(msg, type) {
  const el = document.getElementById('export-status');
  el.textContent = msg;
  el.className = 'export-status export-status-' + type;
}

function showLoading(visible) {
  document.getElementById('loading').style.display = visible ? 'flex' : 'none';
}

function showFileProgress(which, visible) {
  const el = document.getElementById(which + '-progress');
  if (el) el.style.display = visible ? 'block' : 'none';
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g,  '&amp;')
    .replace(/</g,  '&lt;')
    .replace(/>/g,  '&gt;')
    .replace(/"/g,  '&quot;')
    .replace(/'/g,  '&#39;');
}

})(); // end IIFE
