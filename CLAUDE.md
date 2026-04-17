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
| `contractor-app.js` | Contractor Invoices — four sub-tabs (2026 Tasks, Pre-2026 Tasks, TX-RX Tasks, POC Invoices), generates one styled XLSX per contractor |
| `finance-app.js` | Finance Sheet — two sub-tabs (TX-RF Track, POC Tracking), filters by invoice numbers, exports styled XLSX |
| `styles.css` | All styling for all four tabs, sub-tab bars, and shared components |
| `sw.js` | Service worker — caches app shell for offline use |

All app files are wrapped in IIFEs to avoid global namespace collisions. They share two CDN libraries loaded in `index.html`:
- **SheetJS** (`XLSX`) — reading all Excel formats (.xlsx, .xls, .xlsm, .xlsb, .csv)
- **ExcelJS** — writing styled Excel output (all four apps; TSR was migrated from SheetJS `writeFile` to ExcelJS to support header styling)

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
| Status | `=== 'status'` (exact) |

Optional columns (export only, no error if absent): Distance, PO Status, Acceptance Status, VF Task Owner, Vendor, Site Option, Facing, Task Date, PRQ, Certificate, ID#, Comments.

### Column Detection — TSR File

Columns are also detected from the TSR header row after it is located. Fallbacks to original hardcoded positions if the header is not found:

| Column | Header pattern | Fallback |
|---|---|---|
| Item Description | `includes('item description')` | col 6 |
| Unit Price | `includes('unit price')` | col 12 |
| Remaining Qty | `includes('remaining')` (excluding 'after') | col 50 |

### Analysis Logic

**Processing order: TSR file is read first**, then the tracking file. The TSR remaining quantities are fully loaded into `tsrMap` before any combo classification begins.

**Row-level exclusions (applied before grouping):**
- Rows where PO Status is filled → excluded (already submitted to a PO)
- Rows where the Comments cell matches `/\buso\d*\b/i` → excluded (USO scope, not invoiced by LMP). Matches `USO`, `uso`, `USO1`, `uso2`, `USO123`, etc. but not words that merely contain "uso" (e.g. "cursor"). Comments column is optional — if not found the check is silently skipped.

**Combo classification** — each (Job Code + Logical Site ID) group is classified using the `Status` column values (`Done`, `Assigned`, `Cancelled`):

| Situation | Classification | Output status |
|---|---|---|
| All tasks Cancelled | Excluded silently | — |
| Any task is `Assigned` | `some_assigned` | **FAC with some items assigned** |
| All tasks `Done` but some missing FAC Date or Acceptance Week | `some_nfac` | **Some items FAC, some NFAC yet** |
| All tasks `Done` + all have FAC Date + all have Acceptance Week | `eligible` | enters TSR allocation → **Can Submit** or **Need New PO** |

Notes:
- Cancelled rows are stripped before any check — a combo with only Done + Cancelled tasks is treated as if the Cancelled rows don't exist
- A Done task with no FAC Date is treated the same as not-yet-done — it flags the whole combo as `some_nfac`
- Combos with no readable line items after filtering are assigned **Need New PO**

**Greedy TSR allocation (eligible combos only):**
1. Eligible combos are sorted by their earliest Excel row number across all their tasks (first-occurrence rule)
2. For each combo in order: check whether every line item's actual quantity fits within the current TSR available quantity
3. If all items fit → **Can Submit**; deduct those quantities from TSR available for subsequent combos
4. If any item does not fit, or the line item is not found in the TSR at all → **Need New PO** for the entire combo (no partial submission)

Actual quantity = `Absolute Quantity × distanceMultiplier` (based on distance band column). Financial totals use the `newTotal` column directly, not computed from qty × price.

**Money totals (summary boxes):**
- Can Submit → green "Can Submit" box
- `some_assigned` + `some_nfac` combos → amber "Pending" box
- Need New PO → red "Need PO" box
- Cancelled rows contribute nothing to any total

### Export (`exportToExcel` — uses ExcelJS)

- **Row 1:** Column headers — `#0070C0` blue fill, white bold text, centered, thin borders.
- **Row 1 frozen** — stays visible when scrolling (`ws.views = [{ state: 'frozen', ySplit: 1 }]`).
- **Rows 2+:** Data rows with thin borders and 11pt font.
- Export row order: Can Submit → FAC with some items assigned → Some items FAC some NFAC yet → Need New PO.
- File downloaded via `URL.createObjectURL` (same pattern as other apps). `exportToExcel` is `async`; the click handler uses `.catch()` for error handling.

## Contractor App Data Flow (`contractor-app.js`)

The Contractor tab has **four sub-tabs**, each with its own state, file picker, and export logic.

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

**No date filter.** Two interactive combobox filters — VF Invoice # and Contractor — act as AND. Both allow typing (partial, case-insensitive match) or selecting from the dropdown list.

**Amount source:** `Contractor2` column via `parseAmount()`.

State: `_wb2`, `_allRows2`. Entry point: `extractCon2Rows()` → `populateCon2Filter()` → `renderCon2Summary()`.

### Sub-tab 3: TX-RX Tasks

Reads the same **Tracking file (.xlsm)**.

**Filter conditions:**
- Non-In-House contractor only

**No date filter. No Contractor Invoice # filter** — includes all tasks regardless of whether they have already been invoiced. This makes it a complete view of all contractor work, compared to Pre-2026 which only shows uninvoiced rows.

Two interactive combobox filters — VF Invoice # and Contractor — act as AND.

**Amount source:** `Contractor2` column via `parseAmount()`.

State: `_wbTxRx`, `_allRowsTxRx`. Entry point: `extractTxRxRows()` → `populateTxRxFilter()` → `renderTxRxSummary()`.

All element IDs use the `contxrx-*` prefix to avoid collision with Pre-2026 (`con2-*`) IDs.

### Sub-tab 4: POC Invoices

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
- **`buildAndDownload(name, rows)`** — shared by all four sub-tabs. Creates a workbook with a Draft sheet and a Deduction sheet, triggers browser download as `[Name] Draft.xlsx`.

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

### Output Columns — TX-RF Track (`OUTPUT_COLS`)

| Source Column | Output Label | Header Colour | Header Font |
|---|---|---|---|
| Contractor | Contractor | `#0070C0` blue | white |
| Job Code | Job Code | `#FFC000` gold | black |
| Logical Site ID | Site ID | `#FFC000` gold | black |
| Line Item (col 18) | Line Item | `#00B0F0` light blue | white |
| LMP | LMP Portion | `#4472C4` medium blue | white |
| Contractor2 | Contractor Portion | `#4472C4` medium blue | white |
| New Total | New Total Price | `#C00000` dark red | white |
| Task Date | Task Date | `#ED7D31` orange | white |
| VF Invoice # | VF Invoice # | `#2E75B6` steel blue | white |
| PO Number | PO Number | `#2E75B6` steel blue | white |
| Contractor Invoice # | Contractor Invoice # | `#FFD700` yellow | black |
| *(derived)* | Task Type | `#FF0000` red | white |

### Output Columns — POC Tracking (`POC_OUTPUT_COLS`)

Same as TX-RF Track except **Task Date is replaced by two separate date columns**:

| Key | Output Label | Header Colour |
|---|---|---|
| `installDate` | Installation Date | `#ED7D31` orange |
| `migrDate` | Migration Date | `#ED7D31` orange |

Both Installation and Migration output rows carry **both** dates (read from the same source row). This avoids ambiguity — the user can always see both dates regardless of which output row they're reading.

**Task Type column (both sub-tabs):**
- Value is `"New"` if the year ≥ 2026, `"Old"` otherwise.
- For TX-RF Track: derived from the `Task Date` column of each row.
- For POC Tracking: **always derived from the Installation Date**, even for the Migration output row. The migration date is irrelevant for this classification.

**Column detection notes:**
- `Line Item` always hardcoded to col 18 (same as other apps for this file format)
- `Contractor2` matched before `Contractor` to avoid false matches — `includes('contractor') && includes('2')`
- `LMP` exact match (`=== 'lmp'`) preferred; falls back to `includes('lmp')` excluding invoicing/date/status columns
- **PO Number:** matched with `t.startsWith('po') && !t.includes('portion') && t.includes('ins')` — the `startsWith` + `!includes('portion')` guard is required because "LMP Portion ins" and "Contractor Portion ins" both contain the substring `'po'`, which caused false matches before this fix.
- **VF Invoice #:** matched with `t.includes('vf invoice') && !t.includes('date')` — the `!t.includes('date')` guard is required because some tracking files have a "VF Invoice Date" column that appears before "VF Invoice #" in the header row. Without the guard, the date column is matched first and its `Date` object values appear in the filter dropdown instead of invoice number strings. This guard is applied in both `finance-app.js` and `contractor-app.js`.

### Export layout (`exportFinance(rows, filename, cols)`)

- **`cols` parameter** — optional; defaults to `OUTPUT_COLS`. Pass `POC_OUTPUT_COLS` for the POC Tracking export. Each column definition object carries `{ key, label, fill, font, width }` — widths are stored on the definition, not in a separate array.
- **Row 1:** Total row — "Total" centred in the Line Item column (located via `findIndex`), financial totals in LMP/Contractor/NewTotal columns, green fill across all cells.
- **Row 2:** Blank gap row (height 6pt).
- **Row 3:** Column headers.
- **Row 4+:** Data rows.
- Freeze: `ws.views = [{ state: 'frozen', ySplit: 3 }]`
- The `filename` parameter lets POC Tracking export as `POC_Finance_Sheet.xlsx` while TX-RF Track uses its own name.

### UI Filters

All filter controls are **comboboxes** (`<input type="text" list="...">` + `<datalist>`), not plain `<select>` elements. This lets users both pick from the list and type a partial value for real-time filtering.

- **TX-RF Track:** VF Invoice # and Contractor Invoice # (AND logic).
- **POC Tracking:** VF Invoice # and Contractor Invoice # (AND logic).

Filter logic uses case-insensitive `includes` (not exact match), so partial strings work. Event listeners use `input` (not `change`) for keystroke-level reactivity.

**Critical — do not add `autocomplete="off"` to filter inputs.** Firefox suppresses datalist suggestions entirely when `autocomplete="off"` is present on the same input, making the dropdown invisible.

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

When updating any cached file, bump the `CACHE` version string in `sw.js` (e.g. `lmp-invoicing-v9` → `lmp-invoicing-v10`). Without this, installed PWA users will continue running stale files. Current version: `lmp-invoicing-v11`.
