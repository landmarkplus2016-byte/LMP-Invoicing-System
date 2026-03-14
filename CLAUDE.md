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

The tab system in `index.html` is vanilla JS — clicking a tab toggles `.active` and `display` on `.tab-panel` divs.

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

- **Tracking file (.xlsm)** — sheet `Invoicing Track`, header at row index 3 (row 4), data from row index 4. Key hardcoded column positions used alongside header-name detection: col 28 = FAC Date, col 30 = Acceptance Week, col 32 = PO Status, col 18 = Line Item, col 11 = Distance band, col 12 = Absolute Quantity.
- **TSR file (.xlsb)** — sheet `Request Form - VF`, header row found by scanning col G for "item description". Key column positions: col 6 = Item Description, col 12 = Unit Price, col 50 = Remaining Qty.

**Analysis logic:**
1. Groups tracking rows by `(JobCode, LogicalSiteId)` combo key, skipping rows where col 32 (PO Status) is filled
2. Only combos that have a FAC Date (col 28) are "active"
3. Quantities are multiplied by `distanceMultipliers` based on the distance band in col 11
4. Combos are classified into three buckets:
   - **Can Submit** — all rows in combo have FAC Date + Acceptance Week, AND their quantities fit within TSR remaining (greedy first-fit ordered by first Excel row number)
   - **Pending** — not all rows have FAC Date or Acceptance Week
   - **Need PO** — all rows ready but TSR has insufficient remaining quantity
5. Financial totals use `newTotal` column (col index from `newTotalColIndex`), not computed from qty × price

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

## Header Layout

Logo (`LMP Big Logo.jpg`) sits on the **left** of the header, text on the right. Uses `box-sizing: content-box` on `.header-logo` so `height` refers to the image itself and padding is additive (prevents the rounded-corner container from shrinking the visible logo).

## Icons

`icon.svg` is the source design. `icon-192.png` and `icon-512.png` are generated from it using a PowerShell + `System.Drawing` script (no build tooling). To regenerate PNGs after editing the SVG, recreate the drawing commands in PowerShell — the SVG cannot be auto-converted without an external tool like Inkscape.

## Service Worker Cache

When updating any cached file, bump the `CACHE` version string in `sw.js` (e.g. `lmp-invoicing-v1` → `lmp-invoicing-v2`). Without this, installed PWA users will continue running stale files.
