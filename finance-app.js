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
    { key: 'contractor',  label: 'Contractor',          fill: '0070C0', font: 'FFFFFF' },
    { key: 'jobCode',     label: 'Job Code',             fill: '00B050', font: 'FFFFFF' },
    { key: 'siteId',      label: 'Site ID',              fill: '00B050', font: 'FFFFFF' },
    { key: 'lineItem',    label: 'Line Item',            fill: '00B0F0', font: 'FFFFFF' },
    { key: 'lmp',         label: 'LMP Portion',          fill: '4472C4', font: 'FFFFFF' },
    { key: 'contractor2', label: 'Contractor Portion',   fill: '4472C4', font: 'FFFFFF' },
    { key: 'newTotal',    label: 'New Total Price',      fill: 'C00000', font: 'FFFFFF' },
    { key: 'taskDate',    label: 'Task Date',            fill: 'ED7D31', font: 'FFFFFF' },
    { key: 'vfInvoice',   label: 'VF Invoice #',         fill: '2E75B6', font: 'FFFFFF' },
    { key: 'poNumber',    label: 'PO Number',            fill: '2E75B6', font: 'FFFFFF' },
    { key: 'conInvoice',  label: 'Contractor Invoice #', fill: 'FFD700', font: '000000' },
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
    document.getElementById(id).innerHTML =
      '<option value="">-- All --</option>' +
      values.map(v => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
  }

  document.getElementById('fin-filter-vf').addEventListener('change',  onFilterChange);
  document.getElementById('fin-filter-con').addEventListener('change', onFilterChange);
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
    const vf  = document.getElementById('fin-filter-vf').value.trim();
    const con = document.getElementById('fin-filter-con').value.trim();
    return _rows.filter(r => {
      if (vf  && String(r.vfInvoice  ?? '').trim() !== vf)  return false;
      if (con && String(r.conInvoice ?? '').trim() !== con) return false;
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
      await exportFinance(getFilteredRows());
      statusEl.textContent = '\u2705 File downloaded.';
      statusEl.className   = 'export-status-success';
    } catch (err) {
      statusEl.textContent = '\u274c Export failed: ' + err.message;
      statusEl.className   = 'export-status-error';
    } finally {
      btn.disabled = false;
    }
  });

  async function exportFinance(rows) {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'LMP Invoicing System';
    const ws = wb.addWorksheet('Finance Sheet');

    const THIN     = { style: 'thin',   color: { argb: 'FF000000' } };
    const DBL      = { style: 'double', color: { argb: 'FF000000' } };
    const ALL_THIN = { top: THIN, left: THIN, bottom: THIN, right: THIN };
    const ALL_DBL  = { top: DBL,  left: DBL,  bottom: DBL,  right: DBL  };
    const EGP_FMT  = '#,##0.00 "EGP"';

    const COL_WIDTHS = [18, 14, 14, 52, 16, 18, 18, 14, 20, 16, 24];
    OUTPUT_COLS.forEach((_, i) => { ws.getColumn(i + 1).width = COL_WIDTHS[i] || 16; });

    // Header row
    OUTPUT_COLS.forEach((col, i) => {
      const c     = ws.getCell(1, i + 1);
      c.value     = col.label;
      c.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + col.fill } };
      c.font      = { bold: true, color: { argb: 'FF' + col.font }, size: 11 };
      c.alignment = { horizontal: 'center', vertical: 'middle' };
      c.border    = ALL_DBL;
    });
    ws.getRow(1).height = 22;

    // Data rows
    rows.forEach((row, ri) => {
      OUTPUT_COLS.forEach((col, ci) => {
        const c   = ws.getCell(2 + ri, ci + 1);
        const val = row[col.key] ?? '';
        c.value     = val;
        c.border    = ALL_THIN;
        c.font      = { size: 11 };
        c.alignment = { vertical: 'middle', wrapText: col.key === 'lineItem' };
        if (FINANCIAL_KEYS.has(col.key) && typeof val === 'number') c.numFmt = EGP_FMT;
        else if (val instanceof Date) c.numFmt = 'dd-mmm-yy';
      });
      ws.getRow(2 + ri).height = 14.4;
    });

    // Total row
    const totalRn = 2 + rows.length;
    OUTPUT_COLS.forEach((col, ci) => {
      const c  = ws.getCell(totalRn, ci + 1);
      c.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00B050' } };
      c.font   = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
      c.border = ALL_THIN;
      c.alignment = { vertical: 'middle' };
      if (ci === 0) {
        c.value = 'Total';
        c.alignment.horizontal = 'left';
      } else if (FINANCIAL_KEYS.has(col.key)) {
        c.value  = sumCol(rows, col.key);
        c.numFmt = EGP_FMT;
        c.alignment.horizontal = 'right';
      }
    });
    ws.getRow(totalRn).height = 16;

    const buf  = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = URL.createObjectURL(blob);
    const a   = Object.assign(document.createElement('a'), {
      href: url, download: 'Finance_Sheet.xlsx'
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

})(); // end IIFE
