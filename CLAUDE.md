# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

A static, no-build PWA (Progressive Web App) for the Landmark Plus Telecom Department. There is no package manager, no bundler, no test framework, and no server. Open `index.html` directly in a browser — that is the entire dev workflow.

**To run:** Open `index.html` in a browser, or serve the directory with any static file server (e.g. `python -m http.server 8080`).

## Architecture

All logic lives in three files loaded by `index.html` as plain `<script>` tags:

| File | Purpose |
|---|---|
| `poc-app.js` | POC Invoice Prep — reads one Excel file, filters rows, exports styled XLSX |
| `tsr-app.js` | TSR Submission Checker — reads two Excel files, cross-references them, exports XLSX |
| `contractor-app.js` | Contractor Invoices — reads the tracking file, generates one styled XLSX per contractor |
| `styles.css` | All styling for all three tabs and shared components |
| `sw.js` | Service worker — caches app shell for offline use |

All three app files are wrapped in IIFEs to avoid global namespace collisions. They share two CDN libraries loaded in `index.html`:
- **SheetJS** (`XLSX`) — reading all Excel formats (.xlsx, .xls, .xlsm, .xlsb, .csv)
- **ExcelJS** — writing styled Excel output (POC and Contractor apps; TSR app uses SheetJS `writeFile`)

## Layout (`index.html` + `styles.css`)

The app uses a **sidebar + main content** layout (not a top header + tab bar):

- **`.sidebar`** (left, fixed width 230px) — `#0070C0` blue gradient background. Contains:
  - `.sidebar-header` — logo (`LMP Big Logo.jpg`) + brand title/subtitle
  - `.sidebar-nav` — three `.nav-item` tab buttons (one per app)
  - `.sidebar-footer` — **New Analysis** button (`#btn-refresh`) at the bottom; clicking it calls `window.location.reload()` to clear all state
- **`.main-wrapper`** (right, flex:1) — ledger-paper background (`#eaf0f7`). Contains:
  - `.content-header` — white bar showing the current tab's page title (`#page-title`), updated by the tab-switching JS
  - Three `.tab-panel` divs (one per app)
  - `<footer>` — navy blue (`#1a3a5c`) with white text

**Color scheme:**
- Sidebar: `linear-gradient(180deg, #005a9e, #0070C0, #0082d8)`
- Active nav item: `rgba(255,255,255,0.25)` background + `#1a3a5c` left border
- Gold accent stripe on right edge of sidebar: `#c8972a` / `#e8b84b`
- Body background: `#eaf0f7` with repeating ledger-paper grid lines
- Cards: white with `#1a3a5c` top border
- Footer background: `#1a3a5c`

The tab-switching JS updates both the `.nav-item` active state and the `#page-title` text. At ≤768px the sidebar collapses to a 60px icon-only strip.

## POC App Data Flow (`poc-app.js`)

1. User drops/selects an Excel file
2. SheetJS reads it → targets sheet named **`POC3 Tracking`** (exact name required)
3. `detectHeaderRow()` scans the first 30 rows and scores each against `COL_PATTERNS` to find the real header row (handles files with metadata rows above headers)
4. Two filter passes produce two arrays:
   - **Step 1 (Installation):** `installationStatus == "done"` AND `installInvoicingDate` blank AND `lineItem != "POC2 Migration"`
   - **Step 2 (Migration):** `migrationStatus == "done"` AND `acceptanceStatus == "fac"` AND `migInvoicingDate` blank AND `lineItem != "POC2 Migration"`
5. `Invoice Amount` = `Total Amount / 2` for every row
6. ExcelJS writes the output with colour-coded column headers (blue = tracking fields, green = acceptance fields, gold = financial fields) plus a merged total amount cell at row 1

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

Reads the same **Tracking file (.xlsm)** used by the TSR tab (sheet `Invoicing Track`, header row index 3, data from row index 4).

**Filter conditions (all three must be true):**
- `Task Date` year ≥ 2026
- `Acceptance Week` is not blank
- `Contractor Invoice #` is blank (rows not yet invoiced)

**Key implementation notes:**
- `Line Item` is always read from hardcoded **col 18** (same as TSR app) — header-name detection is attempted first but col 18 is the reliable fallback for this file format.
- Contractor names are fuzzy-matched against the canonical list in `list.xlsx` (`Connect`, `DAM Tel`, `El-Khayal`, `New Plan`, `Upper Telecom`) using Levenshtein distance with a 40%-of-name-length threshold.
- `In-House` rows are explicitly excluded before fuzzy matching.
- One `.xlsx` file is generated per contractor, named `[Contractor Name] Draft.xlsx`, each containing:
  - **Draft sheet** — cols B–F: Job Code, Site ID, Facing, Line Item, Amount. Title cell E3:F4 (merged, blue, blank for manual draft number entry). Alternating row colours: first row of each Job Code group = peach (`#F8CBAD`), rest = white. Total row = green (`#00FF00`).
  - **Deduction sheet** — cols B–E only: Job Code, Site ID, Facing, Deduction Amount (empty, for manual entry). No Price/Amount column.

## Icons

`icon.svg` is the source design. `icon-192.png` and `icon-512.png` are generated from it using a PowerShell + `System.Drawing` script (no build tooling). To regenerate PNGs after editing the SVG, recreate the drawing commands in PowerShell — the SVG cannot be auto-converted without an external tool like Inkscape.

## Service Worker Cache

When updating any cached file, bump the `CACHE` version string in `sw.js` (e.g. `lmp-invoicing-v2` → `lmp-invoicing-v3`). Without this, installed PWA users will continue running stale files. Current version: `lmp-invoicing-v3`.
