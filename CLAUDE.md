# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

A static, no-build PWA (Progressive Web App) for the Landmark Plus Telecom Department. There is no package manager, no bundler, no test framework, and no server. Open `index.html` directly in a browser — that is the entire dev workflow.

**To run:** Open `index.html` in a browser, or serve the directory with any static file server (e.g. `python -m http.server 8080`).

## Architecture

All logic lives in four files loaded by `index.html` as plain `<script>` tags:

| File | Purpose |
|---|---|
| `poc-app.js` | POC Invoice Prep — reads one Excel file, filters rows, exports styled XLSX |
| `tsr-app.js` | TSR Submission Checker — reads two Excel files, cross-references them, exports XLSX |
| `contractor-app.js` | Contractor Invoices — three sub-tabs (2026 Tasks, Pre-2026 Tasks, POC Invoices), generates one styled XLSX per contractor |
| `finance-app.js` | Finance Sheet — two sub-tabs (TX-RF Track, POC Tracking), filters by invoice numbers, exports styled XLSX |
| `styles.css` | All styling for all four tabs, sub-tab bars, and shared components |
| `sw.js` | Service worker — caches app shell for offline use |

All app files are wrapped in IIFEs to avoid global namespace collisions. They share two CDN libraries loaded in `index.html`:
- **SheetJS** (`XLSX`) — reading all Excel formats (.xlsx, .xls, .xlsm, .xlsb, .csv)
- **ExcelJS** — writing styled Excel output (POC, Contractor, and Finance apps; TSR app uses SheetJS `writeFile`)

## Layout (`index.html` + `styles.css`)

The app uses a **sidebar + main content** layout (not a top header + tab bar):

- **`.sidebar`** (left, fixed width 230px) — `#0070C0` blue gradient background. Contains:
  - `.sidebar-header` — logo (`LMP Big Logo.jpg`) + brand title/subtitle
  - `.sidebar-nav` — four `.nav-item` tab buttons (one per app)
  - `.sidebar-footer` — **New Analysis** button (`#btn-refresh`) at the bottom; clicking it calls `window.location.reload()` to clear all state
- **`.main-wrapper`** (right, flex:1) — ledger-paper background (`#eaf0f7`). Contains:
  - `.content-header` — white bar showing the current tab's page title (`#page-title`), updated by the tab-switching JS
  - Four `.tab-panel` divs (one per app)
  - `<footer>` — navy blue (`#1a3a5c`) with white text

**Color scheme:**
- Sidebar: `linear-gradient(180deg, #005a9e, #0070C0, #0082d8)`
- Active nav item: `rgba(255,255,255,0.25)` background + **white** (`#ffffff`) 4px left border
- Gold accent stripe on right edge of sidebar: `#c8972a` / `#e8b84b`
- Body background: `#eaf0f7` with repeating ledger-paper grid lines
- Cards: white with `#1a3a5c` top border
- Footer background: `#1a3a5c`

The tab-switching JS updates both the `.nav-item` active state and the `#page-title` text. At ≤768px the sidebar collapses to a 60px icon-only strip.

## Sub-Tab Pattern

Both the Contractor and Finance panels use a pill-style sub-tab UI. **Critical implementation rule: the two panels use different CSS classes to prevent JS cross-contamination.**

| Panel | Button class | Panel class | JS scope |
|---|---|---|---|
| Contractor | `.con-subtab` | `.con-subpanel` | `document.getElementById('panel-contractor').querySelectorAll('.con-subtab')` |
| Finance | `.fin-subtab` | `.fin-subpanel` | `document.getElementById('panel-finance').querySelectorAll('.fin-subtab')` |

If both panels used the same class, the `querySelectorAll` in one panel's switching JS would accidentally activate/deactivate buttons in the other panel.

## POC App Data Flow (`poc-app.js`)

1. User drops/selects an Excel file
2. SheetJS reads it → targets sheet named **`POC3 Tracking`** (exact name required)
3. `detectHeaderRow()` scans the first 30 rows and scores each against `COL_PATTERNS` to find the real header row (handles files with metadata rows above headers)
4. Two filter passes produce two arrays:
   - **Step 1 (Installation):** `installationStatus == "done"` AND `installInvoicingDate` blank AND `lineItem != "POC2 Migration"`
   - **Step 2 (Migration):** `migrationStatus == "done"` AND `acceptanceStatus == "fac"` AND `migInvoicingDate` blank AND `lineItem != "POC2 Migration"`
5. `Invoice Amount` = `Total Amount / 2` for every row
6. ExcelJS writes the output with colour-coded column headers (blue = tracking fields, green = acceptance fields, gold = financial fields) plus a merged total amount cell at row 1
7. Date columns (Installation Date, Migration Date, FAC Date) are written as native Excel date values with format `dd-mmm-yy` — not as strings

## TSR App Data Flow (`tsr-app.js`)

Requires **two** files loaded simultaneously before analysis runs:

- **Tracking file (.xlsm)** — sheet `Invoicing Track`, header at row index 3 (row 4), data from row index 4.
- **TSR file (.xlsb)** — sheet `Request Form - VF`, header row found by scanning any column for "item description".

### Column Detection — Tracking File

**All columns are detected purely by header name (first match, left-to-right). No hardcoded column numbers are used as defaults.** Required columns — a clear error listing missing columns is thrown if any are absent:

| Column | Header pattern matched |
|---|---|
| Job Code | `includes('job code')` |
| Logical Site ID | `includes('logical site')` |
| Line Item | `=== 'line item'` or `=== 'line items'` (exact) |
| FAC Date | `includes('fac date')` or `=== 'fac'` |
| Acceptance Week | `includes('acceptance week')` |
| New Total | `includes('new total')` |
| Absolute Quantity | `includes('absolute')` + `includes('qty')` or `includes('quant')` |

Optional columns (export only, no error if absent): Distance, PO Status, Acceptance Status, VF Task Owner, Vendor, Site Option, Facing, Task Date, PRQ, Certificate, ID#.

### Column Detection — TSR File

Columns are also detected from the TSR header row after it is located. Fallbacks to original hardcoded positions if the header is not found:

| Column | Header pattern | Fallback |
|---|---|---|
| Item Description | `includes('item description')` | col 6 |
| Unit Price | `includes('unit price')` | col 12 |
| Remaining Qty | `includes('remaining')` (excluding 'after') | col 50 |

### Analysis Logic

1. Groups tracking rows by `(JobCode, LogicalSiteId)` combo key, skipping rows where PO Status is filled
2. Only combos that have at least one FAC Date are "active"
3. Actual quantity = `Absolute Quantity × distanceMultiplier` (based on distance band column)
4. Combos are classified into three buckets:
   - **Can Submit** — all rows in combo have FAC Date + Acceptance Week, AND their actual quantities fit within TSR remaining (greedy first-fit ordered by first Excel row number). Requires `liQtyMap.size > 0` — combos with no readable line items are never auto-passed.
   - **Pending** — not all rows have FAC Date or Acceptance Week
   - **Need PO** — all rows ready but TSR has insufficient remaining quantity
5. Financial totals use `newTotal` column, not computed from qty × price
6. Comparison: `actualQty` (with distance multiplier applied) is compared against TSR remaining qty

## Contractor App Data Flow (`contractor-app.js`)

The Contractor tab has **three sub-tabs**, each with its own state, file picker, and export logic.

### Sub-tab 1: 2026 Tasks

Reads the **Tracking file (.xlsm)** (sheet `Invoicing Track`, header row index 3, data from row index 4).

**Filter conditions (all three must be true):**
- `Task Date` year ≥ 2026
- `Acceptance Week` is not blank
- `Contractor Invoice #` is blank (rows not yet invoiced)

**Amount source:** reads directly from the `Contractor2` column via `parseAmount()`. No percentage calculation is applied.

State: `_wb`, `_groups`. Entry point: `analyzeContractors()` → `renderSummary()`.

### Sub-tab 2: Pre-2026 Tasks

Reads the same **Tracking file (.xlsm)**.

**Filter conditions:**
- `Contractor Invoice #` is blank
- Non-In-House contractor

**No date filter.** The VF Invoice # dropdown is the sole filter — user selects an invoice number and all matching rows are grouped by contractor for export.

**Amount source:** `Contractor2` column via `parseAmount()`.

State: `_wb2`, `_allRows2`. Entry point: `extractCon2Rows()` → `populateCon2Filter()` → `renderCon2Summary()`.

### Sub-tab 3: POC Invoices

Reads the **POC3 Tracking file** (same format as Finance Sheet POC Tracking sub-tab). Scans for a header row containing `"inst contractor"` (case-insensitive).

Each source row generates **two output rows** — one for Installation and one for Migration — if the respective contractor field is non-blank and non-In-House.

**Amount source:** reads `conIns` / `conMig` columns directly (contractor portion pre-calculated in the tracking file).

**Filter:** VF Invoice # only (no Contractor Invoice # filter).

**Stats boxes:** Rows / New Total Price / LMP Portion / Contractor Portion — updated live on filter change.

State: `_wbPocCon`, `_allRowsPocCon`. Entry point: `extractPocConRows()` → `populatePocConFilter()` → `renderPocConSummary()`.

### Shared contractor logic

- **Canonical contractor list:** `['Connect', 'DAM Tel', 'El-Khayal', 'New Plan', 'Upper Telecom']`
- **`normalizeContractor(raw)`** — Levenshtein fuzzy match, threshold = 40% of canonical name length (min 3). `In-House` is detected by regex before fuzzy matching.
- **`parseAmount(raw)`** — converts a raw cell value to a number; replaces the old `toContractorAmount()` which applied a 70% multiplier (that multiplier has been removed — amounts now come directly from the `Contractor2` column).
- **`buildAndDownload(name, rows)`** — shared by all three sub-tabs. Creates a workbook with a Draft sheet and a Deduction sheet, triggers browser download as `[Name] Draft.xlsx`.

### Draft / Deduction sheet layout

- **Draft sheet** — cols B–F: Job Code, Site ID, Facing, Line Item, Amount.
  - Title cell E3:F4 (merged, blue fill, blank — for manual invoice number entry).
  - Headers at row 5. Freeze `ySplit: 5`.
  - Data from row 6. Alternating fill: first row of each Job Code group = peach (`#F8CBAD`), rest = white.
  - **Total row at the bottom** (green `#00FF00`).
- **Deduction sheet** — cols B–E: Job Code, Site ID, Facing, Deduction Amount (empty, for manual entry). No amount column.

### Column detection — Contractor2 vs Contractor

`Contractor2` must be matched **before** `Contractor` in the header scan to avoid false matches:
```
Contractor2: t.includes('contractor') && t.includes('2') && !t.includes('invoice')
Contractor:  t === 'contractor' || (t.includes('contractor') && !t.includes('invoice'))
```

## Finance Sheet Data Flow (`finance-app.js`)

The Finance tab has **two sub-tabs**.

### Sub-tab 1: TX-RF Track

Reads the **Tracking file (.xlsm)** (sheet `Invoicing Track`, header row index 3, data from row index 4). **No row filtering is applied** — all non-empty rows are extracted. Filtering is done interactively in the UI via two dropdowns (VF Invoice # and Contractor Invoice #, acting as AND).

### Sub-tab 2: POC Tracking

Reads the **POC3 Tracking file**. Scans for a header row containing `"INST Contractor"`. Each source row generates **two output rows** (Installation + Migration), with `Total Amount ÷ 2` as the `newTotal` for each row.

### Output Columns (both sub-tabs, in order)

| Source Column | Output Label | Header Colour |
|---|---|---|
| Contractor | Contractor | `#0070C0` blue |
| Job Code | Job Code | `#00B050` green |
| Logical Site ID | Site ID | `#00B050` green |
| Line Item (col 18) | Line Item | `#00B0F0` light blue |
| LMP | LMP Portion | `#4472C4` medium blue |
| Contractor2 | Contractor Portion | `#4472C4` medium blue |
| New Total | New Total Price | `#C00000` dark red |
| Task Date | Task Date | `#ED7D31` orange |
| VF Invoice # | VF Invoice # | `#2E75B6` steel blue |
| PO Number | PO Number | `#2E75B6` steel blue |
| Contractor Invoice # | Contractor Invoice # | `#FFD700` yellow |

**Column detection notes:**
- `Line Item` always hardcoded to col 18 (same as other apps for this file format)
- `Contractor2` matched before `Contractor` to avoid false matches — `includes('contractor') && includes('2')`
- `LMP` exact match (`=== 'lmp'`) preferred; falls back to `includes('lmp')` excluding invoicing/date/status columns
- **PO Number:** matched with `t.startsWith('po') && !t.includes('portion') && t.includes('ins')` — the `startsWith` + `!includes('portion')` guard is required because "LMP Portion ins" and "Contractor Portion ins" both contain the substring `'po'`, which caused false matches before this fix.

### Export layout (`exportFinance(rows, filename)`)

- **Row 1:** Total row — "Total" centred in column D, financial totals in LMP/Contractor/NewTotal columns, green fill across all cells.
- **Row 2:** Blank gap row (height 6pt).
- **Row 3:** Column headers.
- **Row 4+:** Data rows.
- Freeze: `ws.views = [{ state: 'frozen', ySplit: 3 }]`
- The `filename` parameter lets POC Tracking export as `POC_Finance_Sheet.xlsx` while TX-RF Track uses its own name.

### UI Filters

- **TX-RF Track:** VF Invoice # and Contractor Invoice # dropdowns (AND logic).
- **POC Tracking:** VF Invoice # dropdown only.

Summary stats (row count, New Total Price sum, LMP Portion sum, Contractor Portion sum) update live on filter change for both sub-tabs.

## Date Output Convention

**All apps** write date values as native Excel dates (JS `Date` objects passed directly to ExcelJS) with number format `dd-mmm-yy`. Do **not** use the `formatDate()` string helper when writing to ExcelJS cells — that produces text strings that Excel cannot sort or filter as dates.

### Timezone fix (critical)

SheetJS with `cellDates: true` creates `Date` objects at **local midnight**. ExcelJS serialises dates using **UTC**. In UTC+2 (Egypt), local midnight = 22:00 previous UTC day, so dates appear one day behind in the output.

**Fix — apply before writing any Date to an ExcelJS cell:**
```javascript
if (val instanceof Date) {
  val = new Date(val.getTime() - val.getTimezoneOffset() * 60000);
  c.value  = val;
  c.numFmt = 'dd-mmm-yy';
}
```
This offset shifts the value so UTC midnight equals the correct local date.

The `formatDate()` function remains in the codebase but is no longer used for Excel output.

## Icons

`icon.svg` is the source design. `icon-192.png` and `icon-512.png` are generated from it using a PowerShell + `System.Drawing` script (no build tooling). To regenerate PNGs after editing the SVG, recreate the drawing commands in PowerShell — the SVG cannot be auto-converted without an external tool like Inkscape.

## Service Worker Cache

When updating any cached file, bump the `CACHE` version string in `sw.js` (e.g. `lmp-invoicing-v6` → `lmp-invoicing-v7`). Without this, installed PWA users will continue running stale files. Current version: `lmp-invoicing-v7`.
