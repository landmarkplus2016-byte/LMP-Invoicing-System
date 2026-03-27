/* ============================================================
   LMP Invoicing System — Finance Sheet logic
   Wrapped in IIFE to avoid naming conflicts with other apps.
   ============================================================ */
(function () {

  let _wb   = null;   // loaded workbook
  let _rows = null;   // all extracted rows

  // ---------------------------------------------------------------------------
  // Output column definitions (key, output label, header fill/font colours)
  // ---------------------------------------------------------------------------
  const OUTPUT_COLS = [
    { key: 'contractor',  label: 'Contractor',          fill: '0070C0', font: 'FFFFFF', width: 18 },
    { key: 'jobCode',     label: 'Job Code',             fill: 'FFC000', font: '000000', width: 14 },
    { key: 'siteId',      label: 'Site ID',              fill: 'FFC000', font: '000000', width: 14 },
    { key: 'lineItem',    label: 'Line Item',            fill: '00B0F0', font: 'FFFFFF', width: 52 },
    { key: 'lmp',         label: 'LMP Portion',          fill: '4472C4', font: 'FFFFFF', width: 16 },
    { key: 'contractor2', label: 'Contractor Portion',   fill: '4472C4', font: 'FFFFFF', width: 18 },
    { key: 'newTotal',    label: 'New Total Price',      fill: 'C00000', font: 'FFFFFF', width: 18 },
    { key: 'taskDate',    label: 'Task Date',            fill: 'ED7D31', font: 'FFFFFF', width: 14 },
    { key: 'vfInvoice',   label: 'VF Invoice #',         fill: '2E75B6', font: 'FFFFFF', width: 20 },
    { key: 'poNumber',    label: 'PO Number',            fill: '2E75B6', font: 'FFFFFF', width: 16 },
    { key: 'conInvoice',  label: 'Contractor Invoice #', fill: 'FFD700', font: '000000', width: 24 },
    { key: 'taskType',    label: 'Task Type',            fill: 'FF0000', font: 'FFFFFF', width: 12 },
  ];

  // POC Tracking uses two separate date columns instead of one Task Date
  const POC_OUTPUT_COLS = [
    { key: 'contractor',  label: 'Contractor',          fill: '0070C0', font: 'FFFFFF', width: 18 },
    { key: 'jobCode',     label: 'Job Code',             fill: 'FFC000', font: '000000', width: 14 },
    { key: 'siteId',      label: 'Site ID',              fill: 'FFC000', font: '000000', width: 14 },
    { key: 'lineItem',    label: 'Line Item',            fill: '00B0F0', font: 'FFFFFF', width: 52 },
    { key: 'lmp',         label: 'LMP Portion',          fill: '4472C4', font: 'FFFFFF', width: 16 },
    { key: 'contractor2', label: 'Contractor Portion',   fill: '4472C4', font: 'FFFFFF', width: 18 },
    { key: 'newTotal',    label: 'New Total Price',      fill: 'C00000', font: 'FFFFFF', width: 18 },
    { key: 'installDate', label: 'Installation Date',    fill: 'ED7D31', font: 'FFFFFF', width: 16 },
    { key: 'migrDate',    label: 'Migration Date',       fill: 'ED7D31', font: 'FFFFFF', width: 16 },
    { key: 'vfInvoice',   label: 'VF Invoice #',         fill: '2E75B6', font: 'FFFFFF', width: 20 },
    { key: 'poNumber',    label: 'PO Number',            fill: '2E75B6', font: 'FFFFFF', width: 16 },
    { key: 'conInvoice',  label: 'Contractor Invoice #', fill: 'FFD700', font: '000000', width: 24 },
    { key: 'taskType',    label: 'Task Type',            fill: 'FF0000', font: 'FFFFFF', width: 12 },
  ];

  const FINANCIAL_KEYS = new Set(['lmp', 'contractor2', 'newTotal']);

  // ---------------------------------------------------------------------------
  // File picker
  // ---------------------------------------------------------------------------
  document.getElementById('btn-pick-fin-track').addEventListener('click', () => {
    document.getElementById('fin-track-input').click();
  });

  document.getElementById('fin-track-input').addEventListener('change', async (e) => {
    clearFinError();
    const file = e.target.files[0];
    if (!file) return;
    finFileProgress(true);
    try {
      const buf = await file.arrayBuffer();
      _wb = XLSX.read(buf, { type: 'array', cellDates: true });
      document.getElementById('fin-track-filename').textContent = file.name;
      document.getElementById('card-fin-track').classList.add('loaded');
      finFileProgress(false);
      runAnalysis();
    } catch (err) {
      finFileProgress(false);
      showFinError('Failed to open file: ' + err.message);
    }
  });

  // ---------------------------------------------------------------------------
  // Analysis
  // ---------------------------------------------------------------------------
  function runAnalysis() {
    clearFinError();
    finLoading(true);
    document.getElementById('fin-results').style.display = 'none';
    setTimeout(() => {
      try {
        _rows = extractRows();
        populateFilters(_rows);
        renderSummary(getFilteredRows());
        finLoading(false);
        document.getElementById('fin-results').style.display = 'block';
        document.getElementById('fin-results').scrollIntoView({ behavior: 'smooth' });
      } catch (err) {
        finLoading(false);
        showFinError(err.message || 'Unexpected error during analysis.');
      }
    }, 50);
  }

  function extractRows() {
    const sheet = _wb.Sheets['Invoicing Track'];
    if (!sheet) throw new Error('Wrong file — expected a sheet named "Invoicing Track".');

    const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
    const header  = allRows[3] ?? [];
    const data    = allRows.slice(4).filter(r => r.some(c => c != null && c !== ''));

    // --- Column detection ---
    const cm = {};
    for (let i = 0; i < header.length; i++) {
      const t = String(header[i] ?? '').trim().toLowerCase();

      if (cm.jobCode    === undefined && t.includes('job code'))              cm.jobCode    = i;
      if (cm.siteId     === undefined && t.includes('logical site'))          cm.siteId     = i;
      if (cm.newTotal   === undefined && t.includes('new total'))             cm.newTotal   = i;
      if (cm.taskDate   === undefined && t.includes('task date'))             cm.taskDate   = i;
      if (cm.vfInvoice  === undefined && t.includes('vf invoice'))            cm.vfInvoice  = i;
      if (cm.poNumber   === undefined &&
          (t.includes('po number') || t.includes('po no') || t === 'po'))    cm.poNumber   = i;
      if (cm.conInvoice === undefined && t.includes('contractor invoice'))    cm.conInvoice = i;

      // Contractor2 must be checked before Contractor to avoid false match
      if (cm.contractor2 === undefined &&
          t.includes('contractor') && t.includes('2') && !t.includes('invoice'))
                                                                              cm.contractor2 = i;
      else if (cm.contractor === undefined &&
               t.includes('contractor') && !t.includes('invoice') && !t.includes('2'))
                                                                              cm.contractor  = i;

      // LMP — exact match first, then broader
      if (cm.lmp === undefined && t === 'lmp')                               cm.lmp = i;
      else if (cm.lmp === undefined && t.includes('lmp') &&
               !t.includes('invoic') && !t.includes('date') && !t.includes('status'))
                                                                              cm.lmp = i;
    }

    // Line Item always col 18 (reliable fallback for this file format)
    cm.lineItem = 18;

    // Validate required columns
    const required = { jobCode: 'Job Code', siteId: 'Logical Site ID', newTotal: 'New Total',
                        taskDate: 'Task Date', vfInvoice: 'VF Invoice #',
                        conInvoice: 'Contractor Invoice #', contractor: 'Contractor' };
    const missing = Object.entries(required)
      .filter(([k]) => cm[k] === undefined)
      .map(([, label]) => label);
    if (missing.length > 0) {
      throw new Error('Columns not found in the header row: ' + missing.join(', '));
    }

    function getVal(row, key) {
      const i = cm[key];
      if (i === undefined || i >= row.length) return '';
      const v = row[i];
      if (v == null) return '';
      return v;
    }

    return data.map(row => {
      const obj = {};
      for (const col of OUTPUT_COLS) {
        let v = getVal(row, col.key);
        if (FINANCIAL_KEYS.has(col.key) && typeof v !== 'number' && v !== '') {
          const n = parseFloat(String(v).replace(/[^0-9.-]/g, ''));
          if (!isNaN(n)) v = n;
        }
        obj[col.key] = v;
      }
      // Derive Task Type from Task Date
      const td = obj.taskDate;
      let yr = (td instanceof Date) ? td.getFullYear() : null;
      if (yr === null && typeof td === 'string' && td) {
        const m = td.match(/\b(20\d{2})\b/);
        if (m) yr = parseInt(m[1]);
      }
      obj.taskType = yr !== null ? (yr >= 2026 ? 'New' : 'Old') : '';
      return obj;
    }).filter(r => OUTPUT_COLS.some(c => r[c.key] !== '' && r[c.key] != null));
  }

  function formatDate(d) {
    const mo = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return `${String(d.getDate()).padStart(2,'0')}-${mo[d.getMonth()]}-${String(d.getFullYear()).slice(-2)}`;
  }

  // ---------------------------------------------------------------------------
  // Filter dropdowns
  // ---------------------------------------------------------------------------
  function populateFilters(rows) {
    const vfSet  = new Set();
    const conSet = new Set();
    rows.forEach(r => {
      const vf  = String(r.vfInvoice  ?? '').trim();
      const con = String(r.conInvoice ?? '').trim();
      if (vf)  vfSet.add(vf);
      if (con) conSet.add(con);
    });
    fillSelect('fin-filter-vf',  [...vfSet].sort());
    fillSelect('fin-filter-con', [...conSet].sort());
  }

  function fillSelect(id, values) {
    const dl = document.getElementById(id + '-list');
    if (dl) dl.innerHTML = values.map(v => `<option value="${esc(v)}"></option>`).join('');
  }

  document.getElementById('fin-filter-vf').addEventListener('input',  onFilterChange);
  document.getElementById('fin-filter-con').addEventListener('input', onFilterChange);
  document.getElementById('btn-fin-clear-filters').addEventListener('click', () => {
    document.getElementById('fin-filter-vf').value  = '';
    document.getElementById('fin-filter-con').value = '';
    onFilterChange();
  });

  function onFilterChange() {
    if (!_rows) return;
    renderSummary(getFilteredRows());
  }

  function getFilteredRows() {
    if (!_rows) return [];
    const vf  = document.getElementById('fin-filter-vf').value.trim().toLowerCase();
    const con = document.getElementById('fin-filter-con').value.trim().toLowerCase();
    return _rows.filter(r => {
      if (vf  && !String(r.vfInvoice  ?? '').toLowerCase().includes(vf))  return false;
      if (con && !String(r.conInvoice ?? '').toLowerCase().includes(con)) return false;
      return true;
    });
  }

  // ---------------------------------------------------------------------------
  // Summary stats
  // ---------------------------------------------------------------------------
  function renderSummary(rows) {
    document.getElementById('fin-stat-rows').textContent  = rows.length.toLocaleString();
    document.getElementById('fin-stat-total').textContent = fmtEGP(sumCol(rows, 'newTotal'));
    document.getElementById('fin-stat-lmp').textContent   = fmtEGP(sumCol(rows, 'lmp'));
    document.getElementById('fin-stat-con2').textContent  = fmtEGP(sumCol(rows, 'contractor2'));
    document.getElementById('btn-export-fin').disabled    = rows.length === 0;
    document.getElementById('fin-export-status').textContent = '';
  }

  function sumCol(rows, key) {
    return rows.reduce((acc, r) => {
      const n = typeof r[key] === 'number' ? r[key]
        : parseFloat(String(r[key] ?? '').replace(/[^0-9.-]/g, ''));
      return acc + (isNaN(n) ? 0 : n);
    }, 0);
  }

  function fmtEGP(val) {
    return 'EGP ' + val.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
  }

  // ---------------------------------------------------------------------------
  // Export
  // ---------------------------------------------------------------------------
  document.getElementById('btn-export-fin').addEventListener('click', async () => {
    if (!_rows) return;
    const btn      = document.getElementById('btn-export-fin');
    const statusEl = document.getElementById('fin-export-status');
    btn.disabled         = true;
    statusEl.textContent = 'Generating…';
    statusEl.className   = '';
    try {
      await exportFinance(getFilteredRows(), 'Finance_Sheet.xlsx');
      statusEl.textContent = '\u2705 File downloaded.';
      statusEl.className   = 'export-status-success';
    } catch (err) {
      statusEl.textContent = '\u274c Export failed: ' + err.message;
      statusEl.className   = 'export-status-error';
    } finally {
      btn.disabled = false;
    }
  });

  async function exportFinance(rows, filename, cols) {
    cols = cols || OUTPUT_COLS;
    const wb = new ExcelJS.Workbook();
    wb.creator = 'LMP Invoicing System';
    const ws = wb.addWorksheet('Finance Sheet');

    const THIN     = { style: 'thin',   color: { argb: 'FF000000' } };
    const DBL      = { style: 'double', color: { argb: 'FF000000' } };
    const ALL_THIN = { top: THIN, left: THIN, bottom: THIN, right: THIN };
    const ALL_DBL  = { top: DBL,  left: DBL,  bottom: DBL,  right: DBL  };
    const EGP_FMT  = '#,##0.00 "EGP"';

    cols.forEach((col, i) => { ws.getColumn(i + 1).width = col.width || 16; });

    // Row 1: Total — "Total" centred in D1 (Line Item column), financial totals in their columns
    const lineItemIdx = cols.findIndex(c => c.key === 'lineItem');
    cols.forEach((col, ci) => {
      const c = ws.getCell(1, ci + 1);
      c.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00B050' } };
      c.font      = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
      c.border    = ALL_THIN;
      c.alignment = { vertical: 'middle' };
      if (ci === lineItemIdx) {
        c.value = 'Total';
        c.alignment.horizontal = 'center';
      } else if (FINANCIAL_KEYS.has(col.key)) {
        c.value  = sumCol(rows, col.key);
        c.numFmt = EGP_FMT;
        c.alignment.horizontal = 'right';
      }
    });
    ws.getRow(1).height = 16;

    // Row 2: blank gap between total and headers
    ws.getRow(2).height = 6;

    // Row 3: Column headers
    cols.forEach((col, i) => {
      const c     = ws.getCell(3, i + 1);
      c.value     = col.label;
      c.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + col.fill } };
      c.font      = { bold: true, color: { argb: 'FF' + col.font }, size: 11 };
      c.alignment = { horizontal: 'center', vertical: 'middle' };
      c.border    = ALL_DBL;
    });
    ws.getRow(3).height = 22;

    // Freeze rows 1-3 (total + gap + header) so header stays visible when scrolling
    ws.views = [{ state: 'frozen', ySplit: 3 }];

    // Data rows (row 4+)
    rows.forEach((row, ri) => {
      cols.forEach((col, ci) => {
        const c   = ws.getCell(4 + ri, ci + 1);
        let   val = row[col.key] ?? '';
        if (val instanceof Date) {
          val = new Date(val.getTime() - val.getTimezoneOffset() * 60000);
          c.value  = val;
          c.numFmt = 'dd-mmm-yy';
        } else {
          c.value = val;
          if (FINANCIAL_KEYS.has(col.key) && typeof val === 'number') c.numFmt = EGP_FMT;
        }
        c.border    = ALL_THIN;
        c.font      = { size: 11 };
        c.alignment = { vertical: 'middle', wrapText: col.key === 'lineItem' };
      });
      ws.getRow(4 + ri).height = 14.4;
    });

    const buf  = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = URL.createObjectURL(blob);
    const a   = Object.assign(document.createElement('a'), {
      href: url, download: filename || 'Finance_Sheet.xlsx'
    });
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // ---------------------------------------------------------------------------
  // UI helpers
  // ---------------------------------------------------------------------------
  function showFinError(msg) {
    const el = document.getElementById('fin-error');
    el.textContent   = msg;
    el.style.display = 'block';
    el.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  function clearFinError() {
    const el = document.getElementById('fin-error');
    el.textContent   = '';
    el.style.display = 'none';
  }

  function finLoading(on) {
    document.getElementById('fin-loading').style.display = on ? 'flex' : 'none';
  }

  function finFileProgress(on) {
    const el = document.getElementById('fin-track-progress');
    if (el) el.style.display = on ? 'block' : 'none';
  }

  function esc(s) {
    return String(s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

// ===========================================================================
// POC TRACKING SUB-TAB
// Reads a POC tracking sheet and produces the same Finance Sheet output.
// Each source row generates TWO output rows (Installation + Migration).
// ===========================================================================

  let _wbPoc   = null;
  let _rowsPoc = null;

  // ── File picker ────────────────────────────────────────────────────────────
  document.getElementById('btn-pick-poc-track').addEventListener('click', () => {
    document.getElementById('poc-track-input').click();
  });

  document.getElementById('poc-track-input').addEventListener('change', async (e) => {
    clearPocError();
    const file = e.target.files[0];
    if (!file) return;
    pocFileProgress(true);
    try {
      const buf = await file.arrayBuffer();
      _wbPoc = XLSX.read(buf, { type: 'array', cellDates: true });
      document.getElementById('poc-track-filename').textContent = file.name;
      document.getElementById('card-poc-track').classList.add('loaded');
      pocFileProgress(false);
      runPocAnalysis();
    } catch (err) {
      pocFileProgress(false);
      showPocError('Failed to open file: ' + err.message);
    }
  });

  // ── Analysis ───────────────────────────────────────────────────────────────
  function runPocAnalysis() {
    clearPocError();
    pocLoading(true);
    document.getElementById('poc-fin-results').style.display = 'none';
    setTimeout(() => {
      try {
        _rowsPoc = extractPocRows();
        populatePocFilters(_rowsPoc);
        renderPocSummary(getPocFilteredRows());
        pocLoading(false);
        document.getElementById('poc-fin-results').style.display = 'block';
        document.getElementById('poc-fin-results').scrollIntoView({ behavior: 'smooth' });
      } catch (err) {
        pocLoading(false);
        showPocError(err.message || 'Unexpected error during analysis.');
      }
    }, 50);
  }

  function extractPocRows() {
    // Accept any sheet — try 'POC3 Tracking' first, fall back to first sheet
    const sheetName = _wbPoc.SheetNames.find(n => n.trim() === 'POC3 Tracking')
                   || _wbPoc.SheetNames[0];
    const sheet = _wbPoc.Sheets[sheetName];
    if (!sheet) throw new Error('No sheet found in the file.');

    const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

    // Detect header row — look for the row that contains "INST Contractor"
    let headerIdx = -1;
    for (let i = 0; i < Math.min(30, allRows.length); i++) {
      if (allRows[i].some(c => String(c ?? '').trim().toLowerCase() === 'inst contractor')) {
        headerIdx = i;
        break;
      }
    }
    if (headerIdx < 0) throw new Error(
      'Cannot find the header row. Expected a cell with exactly "INST Contractor".'
    );

    const header   = allRows[headerIdx];
    const dataRows = allRows.slice(headerIdx + 1)
                            .filter(r => r.some(c => c != null && c !== ''));

    // ── Column detection ──────────────────────────────────────────────────
    const cm = {};
    header.forEach((h, i) => {
      const t = String(h ?? '').trim().toLowerCase();

      // Shared columns
      if (cm.jobCode  === undefined && t === 'job code')           cm.jobCode  = i;
      if (cm.siteId   === undefined && t === 'site id')            cm.siteId   = i;
      if (cm.lineItem === undefined && t === 'line item')          cm.lineItem = i;
      if (cm.total    === undefined && t.includes('total amount')) cm.total    = i;

      // Installation columns
      if (cm.instContractor === undefined
          && t.includes('inst') && t.includes('contractor') && !t.includes('invoice'))
                                                                   cm.instContractor = i;
      if (cm.lmpIns === undefined
          && t.includes('lmp') && t.includes('ins'))               cm.lmpIns = i;
      if (cm.conIns === undefined
          && t.includes('contractor') && t.includes('portion') && t.includes('ins'))
                                                                   cm.conIns = i;
      if (cm.installDate === undefined
          && t.includes('installation') && t.includes('date'))     cm.installDate = i;
      if (cm.invoiceIns === undefined
          && (t.includes('invoice') || t.includes('invoice#'))
          && t.includes('ins') && !t.includes('contractor'))       cm.invoiceIns = i;
      if (cm.poIns === undefined
          && t.startsWith('po') && !t.includes('portion') && t.includes('ins') && !t.includes('mig'))
                                                                   cm.poIns = i;
      if (cm.instConInvoice === undefined
          && t.includes('inst') && t.includes('contractor') && t.includes('invoice'))
                                                                   cm.instConInvoice = i;

      // Migration columns
      if (cm.migrContractor === undefined
          && t.includes('migr') && t.includes('contractor') && !t.includes('invoice'))
                                                                   cm.migrContractor = i;
      if (cm.lmpMig === undefined
          && t.includes('lmp') && t.includes('mig'))               cm.lmpMig = i;
      if (cm.conMig === undefined
          && t.includes('contractor') && t.includes('portion') && t.includes('mig'))
                                                                   cm.conMig = i;
      if (cm.migrDate === undefined
          && t.includes('migration') && t.includes('date'))        cm.migrDate = i;
      if (cm.invoiceMig === undefined
          && (t.includes('invoice') || t.includes('invoice#'))
          && t.includes('mig') && !t.includes('contractor'))       cm.invoiceMig = i;
      if (cm.poMig === undefined
          && t.startsWith('po') && !t.includes('portion') && t.includes('mig'))
                                                                   cm.poMig = i;
      if (cm.migrConInvoice === undefined
          && t.includes('migr') && t.includes('contractor') && t.includes('invoice'))
                                                                   cm.migrConInvoice = i;
    });

    // Validate required columns
    const required = {
      jobCode: 'Job Code', siteId: 'Site ID', lineItem: 'Line Item', total: 'Total Amount',
      instContractor: 'INST Contractor', lmpIns: 'LMP Portion ins',
      installDate: 'Installation Date', invoiceIns: 'Invoice# ins',
      migrContractor: 'MIGR Contractor', lmpMig: 'LMP Portion mig',
      migrDate: 'Migration Date', invoiceMig: 'Invoice# mig',
    };
    const missing = Object.entries(required)
      .filter(([k]) => cm[k] === undefined).map(([, l]) => l);
    if (missing.length > 0) throw new Error('Columns not found in header: ' + missing.join(', '));

    function gv(row, key) {
      const i = cm[key];
      return (i !== undefined && i < row.length) ? (row[i] ?? '') : '';
    }
    function toNum(v) {
      if (typeof v === 'number') return v;
      const n = parseFloat(String(v).replace(/[^0-9.-]/g, ''));
      return isNaN(n) ? '' : n;
    }

    function taskTypeFromDate(d) {
      let yr = (d instanceof Date) ? d.getFullYear() : null;
      if (yr === null && typeof d === 'string' && d) {
        const m = d.match(/\b(20\d{2})\b/);
        if (m) yr = parseInt(m[1]);
      }
      return yr !== null ? (yr >= 2026 ? 'New' : 'Old') : '';
    }

    const rows = [];
    for (const row of dataRows) {
      const totalRaw  = toNum(gv(row, 'total'));
      const halfTotal = typeof totalRaw === 'number' ? totalRaw / 2 : '';

      // Task Type is always based on the installation date for both rows
      const instDate    = gv(row, 'installDate');
      const migrDate    = gv(row, 'migrDate');
      const rowTaskType = taskTypeFromDate(instDate);

      // Installation row
      rows.push({
        contractor:  gv(row, 'instContractor'),
        jobCode:     gv(row, 'jobCode'),
        siteId:      gv(row, 'siteId'),
        lineItem:    gv(row, 'lineItem'),
        lmp:         toNum(gv(row, 'lmpIns')),
        contractor2: toNum(gv(row, 'conIns')),
        newTotal:    halfTotal,
        installDate: instDate,
        migrDate:    migrDate,
        vfInvoice:   String(gv(row, 'invoiceIns')).trim(),
        poNumber:    String(gv(row, 'poIns')).trim(),
        conInvoice:  String(gv(row, 'instConInvoice')).trim(),
        taskType:    rowTaskType,
      });

      // Migration row
      rows.push({
        contractor:  gv(row, 'migrContractor'),
        jobCode:     gv(row, 'jobCode'),
        siteId:      gv(row, 'siteId'),
        lineItem:    gv(row, 'lineItem'),
        lmp:         toNum(gv(row, 'lmpMig')),
        contractor2: toNum(gv(row, 'conMig')),
        newTotal:    halfTotal,
        installDate: instDate,
        migrDate:    migrDate,
        vfInvoice:   String(gv(row, 'invoiceMig')).trim(),
        poNumber:    String(gv(row, 'poMig')).trim(),
        conInvoice:  String(gv(row, 'migrConInvoice')).trim(),
        taskType:    rowTaskType,
      });
    }

    const out = rows.filter(r => r.jobCode !== '' || r.siteId !== '');
    if (out.length === 0) throw new Error('No data rows found in the POC tracking file.');
    return out;
  }

  // ── Filters ────────────────────────────────────────────────────────────────
  function populatePocFilters(rows) {
    const vfSet  = new Set();
    const conSet = new Set();
    rows.forEach(r => {
      if (r.vfInvoice)  vfSet.add(r.vfInvoice);
      if (r.conInvoice) conSet.add(r.conInvoice);
    });
    fillPocSelect('poc-filter-vf',  [...vfSet].sort());
    fillPocSelect('poc-filter-con', [...conSet].sort());
  }

  function fillPocSelect(id, values) {
    const dl = document.getElementById(id + '-list');
    if (dl) dl.innerHTML = values.map(v => `<option value="${esc(v)}"></option>`).join('');
  }

  document.getElementById('poc-filter-vf').addEventListener('input',  onPocFilterChange);
  document.getElementById('poc-filter-con').addEventListener('input', onPocFilterChange);
  document.getElementById('btn-poc-clear-filters').addEventListener('click', () => {
    document.getElementById('poc-filter-vf').value  = '';
    document.getElementById('poc-filter-con').value = '';
    onPocFilterChange();
  });

  function onPocFilterChange() {
    if (!_rowsPoc) return;
    renderPocSummary(getPocFilteredRows());
  }

  function getPocFilteredRows() {
    if (!_rowsPoc) return [];
    const vf  = document.getElementById('poc-filter-vf').value.trim().toLowerCase();
    const con = document.getElementById('poc-filter-con').value.trim().toLowerCase();
    return _rowsPoc.filter(r => {
      if (vf  && !String(r.vfInvoice  ?? '').toLowerCase().includes(vf))  return false;
      if (con && !String(r.conInvoice ?? '').toLowerCase().includes(con)) return false;
      return true;
    });
  }

  // ── Summary stats ──────────────────────────────────────────────────────────
  function renderPocSummary(rows) {
    document.getElementById('poc-fin-stat-rows').textContent  = rows.length.toLocaleString();
    document.getElementById('poc-fin-stat-total').textContent = fmtEGP(sumCol(rows, 'newTotal'));
    document.getElementById('poc-fin-stat-lmp').textContent   = fmtEGP(sumCol(rows, 'lmp'));
    document.getElementById('poc-fin-stat-con2').textContent  = fmtEGP(sumCol(rows, 'contractor2'));
    document.getElementById('btn-export-poc-fin').disabled    = rows.length === 0;
    document.getElementById('poc-fin-export-status').textContent = '';
  }

  // ── Export ─────────────────────────────────────────────────────────────────
  document.getElementById('btn-export-poc-fin').addEventListener('click', async () => {
    if (!_rowsPoc) return;
    const btn      = document.getElementById('btn-export-poc-fin');
    const statusEl = document.getElementById('poc-fin-export-status');
    btn.disabled         = true;
    statusEl.textContent = 'Generating…';
    statusEl.className   = '';
    try {
      await exportFinance(getPocFilteredRows(), 'POC_Finance_Sheet.xlsx', POC_OUTPUT_COLS);
      statusEl.textContent = '\u2705 File downloaded.';
      statusEl.className   = 'export-status-success';
    } catch (err) {
      statusEl.textContent = '\u274c Export failed: ' + err.message;
      statusEl.className   = 'export-status-error';
    } finally {
      btn.disabled = false;
    }
  });

  // ── UI helpers ─────────────────────────────────────────────────────────────
  function showPocError(msg) {
    const el = document.getElementById('poc-fin-error');
    el.textContent   = msg;
    el.style.display = 'block';
    el.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  function clearPocError() {
    const el = document.getElementById('poc-fin-error');
    el.textContent   = '';
    el.style.display = 'none';
  }

  function pocLoading(on) {
    document.getElementById('poc-fin-loading').style.display = on ? 'flex' : 'none';
  }

  function pocFileProgress(on) {
    const el = document.getElementById('poc-track-progress');
    if (el) el.style.display = on ? 'block' : 'none';
  }

})(); // end IIFE
