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

  // ── Step 1: Read TSR file first ──────────────────────────────────────────
  const tsrSheet = tsrWorkbook.Sheets['Request Form - VF'];
  if (!tsrSheet) {
    throw new Error('Wrong file loaded in TSR. Expected sheet: Request Form - VF');
  }

  const tsrRows = XLSX.utils.sheet_to_json(tsrSheet, { header: 1, defval: null });

  let tsrHeaderRowIdx = -1;
  for (let i = 0; i < tsrRows.length; i++) {
    const row = tsrRows[i];
    for (let c = 0; c < row.length; c++) {
      if (row[c] && row[c].toString().toLowerCase().includes('item description')) {
        tsrHeaderRowIdx = i;
        break;
      }
    }
    if (tsrHeaderRowIdx >= 0) break;
  }

  let tsrItemColIndex      = -1;
  let tsrUnitPriceColIndex = -1;
  let tsrRemainingColIndex = -1;

  if (tsrHeaderRowIdx >= 0) {
    const tsrHeader = tsrRows[tsrHeaderRowIdx];
    for (let c = 0; c < tsrHeader.length; c++) {
      const h = (tsrHeader[c] ?? '').toString().trim().toLowerCase();
      if (tsrItemColIndex      < 0 && h.includes('item description'))                   tsrItemColIndex      = c;
      if (tsrUnitPriceColIndex < 0 && h.includes('unit price'))                         tsrUnitPriceColIndex = c;
      if (tsrRemainingColIndex < 0 && h.includes('remaining') && !h.includes('after')) tsrRemainingColIndex = c;
    }
  }

  if (tsrItemColIndex      < 0) tsrItemColIndex      = 6;
  if (tsrUnitPriceColIndex < 0) tsrUnitPriceColIndex = 12;
  if (tsrRemainingColIndex < 0) tsrRemainingColIndex = 50;

  // tsrMap: canonical item name → { remaining, unitPrice }
  const tsrMap = new Map();
  if (tsrHeaderRowIdx >= 0) {
    for (let i = tsrHeaderRowIdx + 1; i < tsrRows.length; i++) {
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

  // Resolve a tracking line item name to its canonical TSR key (fuzzy match)
  function tsrKey(lineItem) {
    if (tsrMap.has(lineItem)) return lineItem;
    for (const key of tsrMap.keys()) {
      if (key.includes(lineItem) || lineItem.includes(key)) return key;
    }
    return null;
  }

  // ── Step 2: Read tracking file ───────────────────────────────────────────
  const trackingSheet = trackingWorkbook.Sheets['Invoicing Track'];
  if (!trackingSheet) {
    throw new Error('Wrong file loaded in Task Tracking. Expected sheet: Invoicing Track');
  }

  const trackingRows = XLSX.utils.sheet_to_json(trackingSheet, { header: 1, defval: null });
  const dataRows     = trackingRows.slice(4);
  const headerRow    = trackingRows[3] ?? [];

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
  let statusColIndex           = -1;

  for (let c = 0; c < headerRow.length; c++) {
    const h = (headerRow[c] ?? '').toString().trim().toLowerCase();
    if (idColIndex               < 0 && (h === 'id#' || h === 'id #'))                                         idColIndex               = c;
    if (jobCodeColIndex          < 0 && h.includes('job code'))                                                 jobCodeColIndex          = c;
    if (logicalSiteColIndex      < 0 && h.includes('logical site'))                                             logicalSiteColIndex      = c;
    if (acceptanceStatusColIndex < 0 && h.includes('acceptance status'))                                        acceptanceStatusColIndex = c;
    if (newTotalColIndex         < 0 && h.includes('new total'))                                                newTotalColIndex         = c;
    if (vfTaskOwnerColIndex      < 0 && h.includes('task owner'))                                               vfTaskOwnerColIndex      = c;
    if (vendorColIndex           < 0 && (h === 'vendor' || h.includes('vendor')))                              vendorColIndex           = c;
    if (siteOptionColIndex       < 0 && h.includes('site option'))                                              siteOptionColIndex       = c;
    if (facingColIndex           < 0 && (h === 'facing' || h.includes('facing')))                              facingColIndex           = c;
    if (taskDateColIndex         < 0 && h.includes('task date'))                                                taskDateColIndex         = c;
    if (prqColIndex              < 0 && (h === 'prq' || h.includes('prq')))                                    prqColIndex              = c;
    if (certificateColIndex      < 0 && h.includes('certificate'))                                              certificateColIndex      = c;
    if (distanceColIndex         < 0 && h.includes('distance'))                                                 distanceColIndex         = c;
    if (absQtyColIndex           < 0 && h.includes('absolute') && (h.includes('qty') || h.includes('quant'))) absQtyColIndex           = c;
    if (lineItemColIndex         < 0 && (h === 'line item' || h === 'line items'))                             lineItemColIndex         = c;
    if (facDateColIndex          < 0 && (h.includes('fac date') || h === 'fac'))                               facDateColIndex          = c;
    if (acceptanceWeekColIndex   < 0 && h.includes('acceptance week'))                                         acceptanceWeekColIndex   = c;
    if (poStatusColIndex         < 0 && (h.includes('po status') || h.includes('po #')
                                      || h.includes('purchase order')))                                         poStatusColIndex         = c;
    if (statusColIndex           < 0 && h === 'status')                                                        statusColIndex           = c;
  }

  const missingCols = [];
  if (jobCodeColIndex        < 0) missingCols.push('"Job Code"');
  if (logicalSiteColIndex    < 0) missingCols.push('"Logical Site ID"');
  if (lineItemColIndex       < 0) missingCols.push('"Line Item"');
  if (facDateColIndex        < 0) missingCols.push('"FAC Date"');
  if (acceptanceWeekColIndex < 0) missingCols.push('"Acceptance Week"');
  if (newTotalColIndex       < 0) missingCols.push('"New Total"');
  if (absQtyColIndex         < 0) missingCols.push('"Absolute Quantity"');
  if (statusColIndex         < 0) missingCols.push('"Status"');

  if (missingCols.length > 0) {
    const foundHeaders = headerRow
      .map((h, i) => (h ? 'col' + i + ':"' + h + '"' : null))
      .filter(Boolean).join(' | ');
    throw new Error(
      'Missing required column(s): ' + missingCols.join(', ') + '. ' +
      'All headers found in row 4: [ ' + foundHeaders + ' ]'
    );
  }

  // Returns 'done' | 'cancelled' | 'assigned' | 'unknown'
  function taskStatus(row) {
    const s = (row[statusColIndex] ?? '').toString().trim().toLowerCase();
    if (s === 'cancelled' || s === 'canceled') return 'cancelled';
    if (s === 'assigned')                       return 'assigned';
    if (s === 'done')                           return 'done';
    return 'unknown';
  }

  // ── Step 3: Group rows by (Job Code + Logical Site ID) combo ─────────────
  // Exclude rows where PO Status is already filled (already submitted)
  const comboRows = new Map();
  dataRows.forEach((row, idx) => {
    if (row[poStatusColIndex] != null && row[poStatusColIndex] !== '') return;
    const combo = comboKey(row, jobCodeColIndex, logicalSiteColIndex);
    if (combo === '|') return;
    if (!comboRows.has(combo)) comboRows.set(combo, []);
    comboRows.get(combo).push({ row, rowIndex: idx });
  });

  // ── Step 4: Classify each combo ──────────────────────────────────────────
  //
  // Combo statuses:
  //   'eligible'      — all non-cancelled tasks are Done + FAC + Acceptance Week
  //                     → enters TSR allocation (becomes 'Can Submit' or 'Need New PO')
  //   'some_assigned' — at least one non-cancelled task is Assigned
  //                     → output: "FAC with some items assigned"
  //   'some_nfac'     — all non-cancelled tasks are Done but some lack FAC date
  //                     or Acceptance Week
  //                     → output: "Some items FAC, some NFAC yet"
  //   (omitted)       — all tasks Cancelled → excluded silently
  //
  const comboClassification = new Map(); // combo → 'eligible' | 'some_assigned' | 'some_nfac'
  const eligibleCombos      = [];        // combos that enter TSR allocation, ordered later by first row

  comboRows.forEach((entries, combo) => {
    // Filter out Cancelled rows — they are treated as non-existent
    const active = entries.filter(({ row }) => taskStatus(row) !== 'cancelled');

    // All tasks cancelled → exclude combo entirely
    if (active.length === 0) return;

    // Any Assigned task → combo is pending, skip TSR allocation
    if (active.some(({ row }) => taskStatus(row) === 'assigned')) {
      comboClassification.set(combo, 'some_assigned');
      return;
    }

    // All remaining are Done — check FAC Date and Acceptance Week for every row
    const allFac     = active.every(({ row }) => row[facDateColIndex] != null && row[facDateColIndex] !== '');
    const allAccWeek = active.every(({ row }) => row[acceptanceWeekColIndex] != null && row[acceptanceWeekColIndex] !== '');

    if (!allFac || !allAccWeek) {
      comboClassification.set(combo, 'some_nfac');
      return;
    }

    // All Done, all FAC, all Acceptance Week → eligible for TSR allocation
    comboClassification.set(combo, 'eligible');
    eligibleCombos.push(combo);
  });

  if (comboClassification.size === 0) {
    throw new Error(
      'No valid sites found after classification. ' +
      'Ensure the Status column contains "Done", "Assigned", or "Cancelled" and that PO Status is empty.'
    );
  }

  // ── Step 5: Build per-combo quantities for TSR allocation ─────────────────
  // Sort eligible combos by their earliest Excel row number (first-occurrence rule)
  const comboFirstRow    = new Map();
  const comboLineItemQty = new Map(); // combo → Map(lineItem → actualQty)

  for (const combo of eligibleCombos) {
    const entries  = comboRows.get(combo);
    let   minRow   = Infinity;
    const liQtyMap = new Map();

    for (const { row, rowIndex } of entries) {
      if (taskStatus(row) === 'cancelled') continue;
      const excelRow = rowIndex + 5;
      if (excelRow < minRow) minRow = excelRow;
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

  // Sort by first row ascending → earlier sites get priority in TSR allocation
  eligibleCombos.sort((a, b) => comboFirstRow.get(a) - comboFirstRow.get(b));

  // ── Step 6: Greedy TSR allocation ─────────────────────────────────────────
  // For each line item, start with the TSR remaining quantity.
  // Process eligible combos in first-row order: if all of a combo's line items
  // fit in the remaining TSR quantities, mark it "Can Submit" and deduct.
  // Otherwise mark it "Need New PO". A single line item not fitting fails the whole combo.
  const tsrAvailable = new Map(); // canonical TSR key → available qty
  for (const [key, data] of tsrMap) {
    tsrAvailable.set(key, data.remaining);
  }

  const canSubmitCombos = new Set();
  const needPoCombos    = new Set();

  for (const combo of eligibleCombos) {
    const liQtyMap = comboLineItemQty.get(combo);

    // No readable line items → cannot submit (data issue)
    if (liQtyMap.size === 0) { needPoCombos.add(combo); continue; }

    let canFit = true;
    for (const [li, qty] of liQtyMap) {
      const canonKey = tsrKey(li);
      // Line item not found in TSR at all → Need New PO
      if (canonKey === null) { canFit = false; break; }
      const available = tsrAvailable.get(canonKey);
      if (available === undefined || qty > available + 0.005) { canFit = false; break; }
    }

    if (canFit) {
      // Deduct this combo's quantities from TSR available
      for (const [li, qty] of liQtyMap) {
        const canonKey = tsrKey(li);
        tsrAvailable.set(canonKey, tsrAvailable.get(canonKey) - qty);
      }
      canSubmitCombos.add(combo);
    } else {
      needPoCombos.add(combo);
    }
  }

  // ── Step 7: Build summary results (line-item view for the TSR table) ──────
  // Include only combos that went through TSR allocation (can submit + need new PO)
  const groups = new Map();
  comboRows.forEach((entries, combo) => {
    if (!canSubmitCombos.has(combo) && !needPoCombos.has(combo)) return;
    entries.forEach(({ row, rowIndex }) => {
      if (taskStatus(row) === 'cancelled') return;
      const lineItem = (row[lineItemColIndex] ?? '').toString().trim();
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

  const results = [];
  for (const [lineItem, g] of groups) {
    const canonKey     = tsrKey(lineItem);
    const tsrEntry     = canonKey !== null ? tsrMap.get(canonKey) : undefined;
    const trackingQty  = Math.round(g.totalQty * 100) / 100;
    const tsrRemaining = tsrEntry !== undefined ? tsrEntry.remaining : null;
    const tsrUnitPrice = tsrEntry !== undefined ? tsrEntry.unitPrice : null;
    const difference   = tsrRemaining !== null ? tsrRemaining - trackingQty : null;

    let status;
    if (tsrRemaining === null)            status = 'NOT_FOUND';
    else if (trackingQty <= tsrRemaining) status = 'OK';
    else                                  status = 'EXCEEDS';

    // Remaining in TSR after all can-submit deductions
    const afterQty = canonKey !== null ? (tsrAvailable.get(canonKey) ?? tsrRemaining) : null;

    results.push({
      lineItem, trackingQty, tsrRemaining, tsrUnitPrice, difference, status,
      excelRowNumbers: g.excelRowNumbers,
      individualQtys:  g.individualQtys,
      facDates:        g.facDates,
      acceptanceWeeks: g.acceptanceWeeks,
      tsrAfterSubmit:  afterQty !== null ? Math.round(afterQty * 100) / 100 : null,
      tsrEgpUnitRate:  tsrUnitPrice !== null ? Math.round(tsrUnitPrice * 100) / 100 : null,
      tsrAfterAmount:  (tsrUnitPrice !== null && afterQty !== null)
                         ? Math.round(afterQty * tsrUnitPrice * 100) / 100 : null
    });
  }

  // ── Step 8: Build export rows ─────────────────────────────────────────────
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

  // Order: Can Submit → FAC with some items assigned → Some items FAC some NFAC yet → Need New PO
  comboRows.forEach((entries, combo) => {
    if (!canSubmitCombos.has(combo)) return;
    entries.forEach(({ row }) => {
      if (taskStatus(row) === 'cancelled') return;
      if ((row[lineItemColIndex] ?? '').toString().trim() === '') return;
      allExportItems.push(makeExportRow(row, 'Can Submit'));
    });
  });

  comboRows.forEach((entries, combo) => {
    if (comboClassification.get(combo) !== 'some_assigned') return;
    entries.forEach(({ row }) => {
      if (taskStatus(row) === 'cancelled') return;
      if ((row[lineItemColIndex] ?? '').toString().trim() === '') return;
      allExportItems.push(makeExportRow(row, 'FAC with some items assigned'));
    });
  });

  comboRows.forEach((entries, combo) => {
    if (comboClassification.get(combo) !== 'some_nfac') return;
    entries.forEach(({ row }) => {
      if (taskStatus(row) === 'cancelled') return;
      if ((row[lineItemColIndex] ?? '').toString().trim() === '') return;
      allExportItems.push(makeExportRow(row, 'Some items FAC, some NFAC yet'));
    });
  });

  comboRows.forEach((entries, combo) => {
    if (!needPoCombos.has(combo)) return;
    entries.forEach(({ row }) => {
      if (taskStatus(row) === 'cancelled') return;
      if ((row[lineItemColIndex] ?? '').toString().trim() === '') return;
      allExportItems.push(makeExportRow(row, 'Need New PO'));
    });
  });

  allExportItems.sort((a, b) => {
    const s = a.logicalSiteId.localeCompare(b.logicalSiteId);
    return s !== 0 ? s : a.jobCode.localeCompare(b.jobCode);
  });

  // ── Step 9: Money totals ──────────────────────────────────────────────────
  // Can Submit → moneyCanSubmit
  // FAC with some items assigned + Some items FAC some NFAC yet → moneyPending
  // Need New PO → moneyNeedPo
  // Cancelled rows are excluded from all totals
  moneyCanSubmit = 0;
  moneyPending   = 0;
  moneyNeedPo    = 0;

  comboRows.forEach((entries, combo) => {
    const cls = comboClassification.get(combo);
    if (!cls) return; // all-cancelled combo, not in map
    entries.forEach(({ row }) => {
      if (taskStatus(row) === 'cancelled') return;
      const val    = newTotalColIndex >= 0 ? row[newTotalColIndex] : null;
      const amount = (val != null && val !== '') ? (Number(val) || 0) : 0;
      if (amount === 0) return;
      if (canSubmitCombos.has(combo))      moneyCanSubmit += amount;
      else if (needPoCombos.has(combo))    moneyNeedPo    += amount;
      else                                 moneyPending   += amount;
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
  exportToExcel().catch(err => {
    showExportStatus('\u274c Export failed: ' + err.message, 'error');
  });
});

async function exportToExcel() {
  const HEADERS = [
    'VF Task Owner', 'Vendor', 'Logical Site ID', 'Site Option', 'Facing',
    'Task Date', 'Line Item', 'Absolute Quantity', 'PRQ', 'Certificate #',
    'Acceptance Status', 'Actual Quantity', 'New Total Price', 'ID#', 'Job Code', 'Comment'
  ];
  const COL_WIDTHS = [18, 14, 18, 14, 12, 14, 60, 18, 10, 16, 20, 16, 16, 16, 14, 14];

  const wb = new ExcelJS.Workbook();
  wb.creator = 'LMP Invoicing System';
  const ws = wb.addWorksheet('TSR Analysis');

  HEADERS.forEach((_, i) => { ws.getColumn(i + 1).width = COL_WIDTHS[i]; });

  // Header row — blue fill, white bold text
  const THIN    = { style: 'thin', color: { argb: 'FF000000' } };
  const ALL_THIN = { top: THIN, left: THIN, bottom: THIN, right: THIN };

  const headerRow = ws.addRow(HEADERS);
  headerRow.height = 22;
  headerRow.eachCell(c => {
    c.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0070C0' } };
    c.font      = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    c.border    = ALL_THIN;
  });

  // Freeze the header row
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  // Data rows
  for (const r of allExportItems) {
    const row = ws.addRow([
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
    row.height = 14.4;
    row.eachCell({ includeEmpty: true }, c => {
      c.border = ALL_THIN;
      c.font   = { size: 11 };
      c.alignment = { vertical: 'middle' };
    });
  }

  const today = new Date().toISOString().slice(0, 10);
  const filename = 'TSR_Submission_Results_' + today + '.xlsx';
  const buf  = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  const url = URL.createObjectURL(blob);
  const a   = Object.assign(document.createElement('a'), { href: url, download: filename });
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  const canCount      = allExportItems.filter(r => r.comment === 'Can Submit').length;
  const assignedCount = allExportItems.filter(r => r.comment === 'FAC with some items assigned').length;
  const nfacCount     = allExportItems.filter(r => r.comment === 'Some items FAC, some NFAC yet').length;
  const needPoCount   = allExportItems.filter(r => r.comment === 'Need New PO').length;
  showExportStatus(
    '\u2705 Export downloaded! ' +
    canCount + ' can submit, ' +
    assignedCount + ' FAC/assigned, ' +
    nfacCount + ' some NFAC, ' +
    needPoCount + ' need new PO.',
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
