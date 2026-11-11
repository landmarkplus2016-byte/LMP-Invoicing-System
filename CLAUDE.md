# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

A static, no-build PWA (Progressive Web App) for the Landmark Plus Telecom Department. There is no package manager, no bundler, no test framework, and no server. Open `index.html` directly in a browser — that is the entire dev workflow.

**To run:** Open `index.html` in a browser, or serve the directory with any static file server (e.g. `python -m http.server 8080`).

## Architecture

All logic lives in the files below, loaded by `index.html` as plain `<script>` tags:

| File | Purpose |
|---|---|
| `poc-app.js` | POC Invoice Prep — reads one Excel file, filters rows, exports styled XLSX |
| `tsr-app.js` | TSR Sub Prep — reads two Excel files, cross-references them, exports XLSX |
| `tsr-validation-app.js` | TSR Sub Validation — validates a TSR submission against Excel/PDF mail attachments |
| `tsr-order-app.js` | TSR Sub Order — lists one TSR Sub#'s tasks in acceptance-mail order, capped by TSR qty, up to a target amount |
| `tsr-core.js` | Shared TSR rules (`window.TSRCore`) — matching keys, distance bands, week parsing, mail reading, tracking + TSR file readers. Must load before `tsr-planner-app.js` |
| `tsr-planner-app.js` | TSR Planner — readiness check of every Site + Job Code combo, oldest-first plan within the TSR Remaining, paste-ready blocks for the protected PO Break Down sheet, receiving template, mail folders, allocation list |
| `tsr-newpo-app.js` | TSR Planner › **Plan a New PO** — Request Form quantities for a new TSR / PO up to a budget, what stays in the current TSR, sent-task ID list. Uses the files loaded in the Planner (`window.TSRPlanner`); must load after `tsr-planner-app.js` |
| `file-sources.js` | My Files — remembers file / folder locations (File System Access API, Chrome / Edge) and loads them into every tab with one **Load My Files** click; shares one parsed tracking workbook across tabs |
| `contractor-app.js` | Contractor Invoices — four sub-tabs (2026 Tasks, Pre-2026 Tasks, TX-RX Tasks, POC Invoices), generates one styled XLSX per contractor |
| `finance-app.js` | Finance Sheet — two sub-tabs (TX-RF Track, POC Tracking), filters by invoice numbers, exports styled XLSX |
| `acceptance-check-app.js` | Acceptance Check — compares tracking line items for one week + area against the acceptance sheet (`Total` tab) and lists mismatches |
| `styles.css` | All styling for all six tabs, sub-tab bars, and shared components |
| `sw.js` | Service worker — caches app shell for offline use |

All app files are wrapped in IIFEs to avoid global namespace collisions. They share two CDN libraries loaded in `index.html`:
- **SheetJS** (`XLSX`) — reading all Excel formats (.xlsx, .xls, .xlsm, .xlsb, .csv)
- **ExcelJS** — writing styled Excel output (all four export apps; TSR was migrated from SheetJS `writeFile` to ExcelJS to support header styling)

Two additional libraries are loaded **on demand** (dynamic `import()` / script injection) only when needed:
- **`@kenjiuno/msgreader`** (from `esm.sh`) — parses Outlook `.msg` email files to extract attachments
- **PDF.js** (`pdfjsLib`, from cdnjs) — extracts text content from PDF files for TOC validation

## Layout (`index.html` + `styles.css`)

The app uses a **sidebar + main content** layout (not a top header + tab bar):

- **`.sidebar`** (left, fixed width 230px) — `#0070C0` blue gradient background. Contains:
  - `.sidebar-header` — logo (`LMP Big Logo.jpg`) + brand title/subtitle
  - `.sidebar-nav` — five `.nav-item` tab buttons (one per app)
  - `.sidebar-footer` — **New Analysis** button (`#btn-refresh`) at the bottom; clicking it calls `window.location.reload()` to clear all state
- **`.main-wrapper`** (right, flex:1) — ledger-paper background (`#eaf0f7`). Contains:
  - `.content-header` — white bar showing the current tab's page title (`#page-title`), updated by the tab-switching JS
  - Five `.tab-panel` divs (one per app)
  - `<footer>` — navy blue (`#1a3a5c`) with white text

**Color scheme:**
- Sidebar: `linear-gradient(180deg, #005a9e, #0070C0, #0082d8)`
- Active nav item: `rgba(255,255,255,0.25)` background + **white** (`#ffffff`) 4px left border
- Gold accent stripe on right edge of sidebar: `#c8972a` / `#e8b84b`
- Body background: `#eaf0f7` with repeating ledger-paper grid lines
- Cards: white with `#1a3a5c` top border
- Footer background: `#1a3a5c`

The tab-switching JS updates both the `.nav-item` active state and the `#page-title` text. At ≤768px the sidebar collapses to a 60px icon-only strip.

Every panel's `<main>` (`.poc-main`, `.tsr-main`, `.contractor-main`) is centred with `max-width: 1600px`, so wide screens use the space beside the sidebar.

## Drag & Drop (all tabs)

A single inline script in `index.html` makes every `.file-card` a drop target — no per-app code. A dropped file is put into the card's own `<input type="file">` via `DataTransfer` and a synthetic `change` event is dispatched, so each app's existing change handler (which reads `e.target.files`) loads it unchanged. It also adds a `.drop-hint` line to each card and toggles `.drag-over` while dragging. `dragover`/`drop` are cancelled on `window` so a missed drop doesn't navigate the browser to the file.

Folder inputs (`webkitdirectory`, i.e. the TSR Mails Folder) can't be filled that way: the script walks the dropped directory with `webkitGetAsEntry()`, rebuilds `webkitRelativePath` (`Root/1/mail.msg`) on each file, and dispatches a `folderdrop` CustomEvent (`detail` = files) on the input. `tsr-validation-app.js` listens for it and calls the same `loadFolderFiles()` as its `change` handler. The POC tab keeps its own `#dropZone` handler.

**New upload cards get drag & drop automatically** as long as they use `.file-card` with an `<input type="file">` inside.

## My Files — remembered sources (`file-sources.js`)

A **My Files** tab (`panel-sources`, IDs `src-*`) plus a gold **⚡ Load My Files** button in the sidebar footer (`#btn-load-sources`, last-load time in `#btn-load-sources-note`). Chrome / Edge only (`showOpenFilePicker` / `showDirectoryPicker`); other browsers see a notice and keep the normal Browse buttons.

**Sources (`SOURCES`)** — each a file or folder handle saved in IndexedDB (`lmp-invoicing` / `sources`, key `h:<id>`; this browser on this computer only):

| id | Kind | Fills these inputs |
|---|---|---|
| `tracking` | file | `tracking-input`, `tsrval-input-trk`, `tso-input-track`, `tsp-input-track`, `con-track-input`, `con2-track-input`, `contxrx-track-input`, `fin-track-input`, `acc-input-track` |
| `mails` | folder | `tso-input-folder`, `tsp-input-folder` (via `folderdrop`) |
| `accsheet` | file | `acc-input-acc` |
| `poc` | file | `fileInput` (POC tab), `con-poc-track-input`, `poc-track-input` |
| `tsrroot` | folder | the selected TSR file → `tsr-input`, `tsrval-input-tsr`, `tso-input-tsr`, `tsp-input-tsr`; its Sub folder's **numbered** sub-folders → `tsrval-input-folder` |

**TSR file choice** — `scanTsr()` lists every `.xlsb` up to `Root/<TSR#>/<Sub-n>/` (newest first) in `#src-tsr-select`; the choice is saved as `tsr:selected` (relative path), default = newest.

**Loading (`loadAll`)** — on the button click (user activation): `queryPermission` → `requestPermission({mode:'read'})` for each handle (Chrome asks once per session unless the user picks *Allow on every visit*; a `SecurityError` asks the user to click again). Then each file is read with `handle.getFile()` and handed to every input of that source exactly like drag & drop — `DataTransfer` + `change` event for file inputs, `folderdrop` CustomEvent with `webkitRelativePath` rebuilt (`Root/sub/…/name`) for folder inputs — so **no app needed changing**. A missing file (`NotFoundError`) is shown as "moved or renamed? Set it again".

**Wrong-file guard** — file sources declare the sheets they must contain (`tracking` → `Invoicing Track`, `accsheet` → `Total`, `poc` → `POC3 Tracking`, the TSR file → `Request Form - VF` + `PO Break Down- Contractor & Acc`). `missingSheets()` reads only the sheet names (`bookSheets: true`) when a file is picked (a wrong file is not saved) and again on every load (a wrong file is never handed to the tabs; the row shows "Wrong file … it has: …"). Added after a neighbouring workbook (`TX_Upgrade_Eligibility_Apr2026.xlsx`) was picked as the tracking file and every tab showed "Invoicing Track not found".

**Shared read** — `File.prototype.arrayBuffer` is wrapped so every File with the same `name|size|lastModified` as a loaded source returns one cached buffer, and `XLSX.read` is wrapped to cache the workbook per buffer + options (only for those registered buffers). One load of all tabs = one parse of the tracking file (tested: 3 parses for 14 inputs). Safe because no app mutates a workbook. The POC tab reads with `FileReader`, so it parses its own copy.

**Testing note** — reading a `FileSystemHandle` back from IndexedDB crashes the sandboxed headless Chromium used for tests; it is the standard pattern in real Chrome. Tests use `window.LMPSources.setHandle()` with OPFS handles and load in the same page.

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
4. `parseWorkbook()` collects every distinct value in the **Installation Invoicing Date ins** and **Migration Invoicing Date mig** columns (`collectBatches`) and fills the **Invoice Batch** `<select>` (`#pocBatchSelect`). Text labels the user types for the month's batch (`new oct`, `oct`) are listed first; old rows holding real dates are listed after, newest first. Values compare via `batchKey()` (case/whitespace-insensitive text, or calendar day for dates). Nothing is analysed until a batch is selected.
5. `analyzeBatch()` produces two arrays — no status/FAC/line-item filtering, the batch label is the decision:
   - **Step 1 (Installation):** `installInvoicingDate` == selected batch
   - **Step 2 (Migration):** `migInvoicingDate` == selected batch
6. `Invoice Amount` = `Total Amount / 2` for every row
6a. `Contractor` output column (after VF Owner) is per step: Installation rows take **INST Contractor**, Migration rows take **MIGR Contractor**. Detection skips headers containing `invoice` or `portion` (`COL_EXCLUDE`) so `INST Contractor Invoice#` / `Contractor Portion ins` are never picked. A missing column only produces a warning.
7. ExcelJS writes the output with colour-coded column headers (blue = tracking fields, green = acceptance fields, gold = financial fields) plus a merged total amount cell at row 1
8. Date columns (Installation Date, Migration Date, FAC Date) are written as native Excel date values with format `dd-mmm-yy` — not as strings

## TSR Sub Prep Data Flow (`tsr-app.js`)

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
1. Eligible combos are sorted by **acceptance week ascending, then by earliest Excel row number** (first-occurrence rule within a week)
2. For each combo in order: check whether every line item's actual quantity fits within the current TSR available quantity
3. If all items fit → **Can Submit**; deduct those quantities from TSR available for subsequent combos
4. If any item does not fit, or the line item is not found in the TSR at all → **Need New PO** for the entire combo (no partial submission)

Actual quantity = `Absolute Quantity × distanceMultiplier` (based on distance band column). Financial totals use the `newTotal` column directly, not computed from qty × price.

**Acceptance week grouping (`parseAcceptanceWeek`):**

The allocation prefers to keep one submission inside a single acceptance week.

- **Grouping is by the full label string, not by the week number.** Real values carry a meaningful prefix — `D-`, `U-`, `A-` — so `D-W05-W06-2026` and `U-W05-2026` share a week number but are different groups and must not be mixed in one submission.
- `parseAcceptanceWeek(raw)` parses values like `D-W33-W34-2026` into `{ label, sortKey }` where `sortKey = year * 100 + lowest week number` (so `D-W33-W34-2026` → `202633`, and the combined form `U-W09-2026/U-W15-2026` → `202609`). The `sortKey` is used **only for ordering**; the `label` is what groups combos. Values with no `20xx` year or no `W<n>` token get `sortKey = Infinity` and sort last; blank values give an empty label.
- A combo whose tasks span more than one acceptance week is assigned its **earliest** week. It remains a single all-or-nothing unit.
- Sort order is week `sortKey` ascending → full `label` (`localeCompare`) → `comboFirstRow` ascending. The label tiebreaker keeps same-numbered labels with different prefixes contiguous instead of interleaved. Combos with `Infinity` week keys always sort after parseable ones.
- **"Prefer, then fill"** — the oldest label is allocated first, but once it is exhausted the greedy loop keeps going into the following labels so leftover TSR quantity is not wasted. The same-week rule is a priority, not a hard filter.
- `weekSummary` (module state) records the target label and how much of it fit: `{ label, totalTasks, selectedTasks, totalCombos, selectedCombos, otherTasks }`. `renderResults` renders it into `#week-summary` as a blue `.week-card` above the money cards (e.g. "24 of 38 tasks selected (63%)"), and hides the element when there are no eligible combos.

**Money totals (summary boxes):**
- Can Submit → green "Can Submit" box
- `some_assigned` + `some_nfac` combos → amber "Pending" box
- Need New PO → red "Need PO" box
- Cancelled rows contribute nothing to any total

### Export (`exportToExcel` — uses ExcelJS)

- **Columns:** VF Task Owner, Vendor, Logical Site ID, Site Option, Facing, Task Date, Line Item, Absolute Quantity, PRQ, Certificate #, Acceptance Status, **Acceptance Week**, Actual Quantity, New Total Price, ID#, Job Code, Comment. `HEADERS` and `COL_WIDTHS` are positional and must stay the same length as the data-row array in `exportToExcel`.
- **Row 1:** Column headers — `#0070C0` blue fill, white bold text, centered, thin borders.
- **Row 1 frozen** — stays visible when scrolling (`ws.views = [{ state: 'frozen', ySplit: 1 }]`).
- **Rows 2+:** Data rows with thin borders and 11pt font.
- Export row order: Can Submit → FAC with some items assigned → Some items FAC some NFAC yet → Need New PO.
- File downloaded via `URL.createObjectURL` (same pattern as other apps). `exportToExcel` is `async`; the click handler uses `.catch()` for error handling.

## TSR Sub Validation (`tsr-validation-app.js`)

Validates a TSR submission by comparing it against Excel or PDF attachments extracted from numbered mail folders.

### UI

- **TSR File upload** — the `.xlsb` TSR file (sheet `PO Break Down- Contractor & Acc`)
- **TSR Mails Folder upload** — a folder whose sub-folders are numbered (1, 2, 3…) — FAC and TOC folders alike. Each sub-folder contains an email (`.msg` or `.eml`) whose attachment is the reference file, or a direct `.xlsx`/`.pdf` file.
- **Submission Number** — integer typed by the user; filters TSR rows by the Submission # column
- **Master Tracking (optional)** — only used to fill blank Distance values for the quantity check

Both file inputs show an animated progress bar while loading. Results show a grey diagnostic strip listing the detected column letters for each field.

### TSR File — Sheet and Columns

Sheet: **`PO Break Down- Contractor & Acc`** (exact name required; error lists available sheets if not found).

Header row is located by scanning ALL rows for one that contains both `"item description"` AND `"site id"` in the same row. This dual-keyword check avoids false positives from title rows that mention only one keyword.

| Column | Header pattern | Fallback |
|---|---|---|
| Site ID | `includes('site id')` or `=== 'site_id'` | H (7) |
| Facing Site # | `includes('facing')` | J (9) |
| Item Description | `includes('item description')` | N (13) |
| Certificate # | `includes('certificate')` | U (20) |
| Contractor Status | `includes('contractor') && includes('status')` | V (21) |
| Submission # | `includes('submission') && !includes('date')` | W (22) |
| Contractor Comments | `includes('contractor') && includes('comment')` | Y (24) |

**Submission # value normalisation** — `extractNumber()` handles any format: bare integer `5`, float `5.0`, text `"Submission 5"`, `"Sub 5"` etc. Date objects are ignored (return `null`).

**Contractor Comments → folder key (FAC and TOC rows alike):**
- `folderKey = normalizeFolderNumber(comment)` — finds the first digit sequence **anywhere** in the comment. Handles `"Folder 1"`, `"Folder 1 (not the same link)"`, `"1"`, `"1.0"` → all → `1`.
- There is **no "TOC N" rule any more** — every mail folder is a plain number, whether its rows are FAC or TOC. Contractor Status = `"TOC"` only sets `isToc` on the row, which gives the folder section a teal **TOC** badge; it does not change routing or validation.

### Folder Map Construction (`buildFolderDataMap`)

Iterates all uploaded files grouped by sub-folder name. Every sub-folder is keyed by its number (`normalizeFolderNumber(subName)`); sub-folders with no number are ignored.

For each folder: Excel preferred, PDF (Completion Certificate) accepted — tries direct Excel → direct PDF → `.msg` attachment (Excel, else PDF) → `.eml` attachment (Excel, else PDF).

The validation path is chosen by the **attachment type actually found** (`folderData.type`) — an Excel is validated with the Excel rules, a PDF with the PDF rules below. PDF folders that hold no TOC rows show a grey **PDF** badge in the results header.

- `.msg` parsing: `@kenjiuno/msgreader` loaded via `import('https://esm.sh/@kenjiuno/msgreader')` (cached after first load)
- `.eml` parsing: plain text MIME parser (base64 attachment extraction, no library)
- PDF.js: loaded by injecting a `<script>` tag pointing to the cdnjs CDN (cached after first load); worker also loaded from cdnjs

### Validation Logic

**Phase 1 — Matching (per folder group):**

**Excel attachments (regular numbered folders):**
- Key: `Site ID + Facing # + itemMatchKey(Item Description)`
- `itemMatchKey()` extracts the catalogue code prefix: `"EX06 - S&I of …"` → `"EX06"`. This makes matching prefix-invariant so minor description wording/punctuation differences don't cause failures. Renamed codes go through the same `ITEM_ALIASES` table as TSR Sub Order (`DS07` → `DS09`) and merged RF items through `ITEM_GROUPS`, so a TSR written with the new price-list name still matches a mail that kept the old one. PDF rows compare via `pdfItemKey()` (group from the description, else the aliased activity code).
- **One-to-one pairing (`excelPair`)** — the same Site + Facing + item can appear several times in one mail (a hub site's `TX16` for several links, each with its own Request #), so every TSR row is paired with its own mail row before any check, in four passes over the group: (1) same Site + Facing + item **and** Request # = Certificate #, (2) same Site + Facing + item, (3) same Site + item with Request # = Certificate # (Facing differs), (4) same Site + item. Each mail row is used once. Passes 1–2 are a full match; passes 3–4 report a Facing mismatch.
  - Before this, only the first mail row per key was kept, so all repeats of `H3153 / TX16` mapped to one position — giving false Certificate # mismatches and a cascade of false order errors.
- After matching: Certificate # (TSR) is validated against Request # (Excel). This is the only additional field check — Site/Facing/Item were already used as the key.
- A TSR row whose Site + item rows in the mail are all taken by other TSR rows fails as **Duplicate** (naming those TSR rows) instead of "Row not found".

**PDF attachments (TOC folders and numbered folders holding a Completion Certificate):**

The only fields the certificate PDF carries are Site ID, Activity code, Activity Description and Quantity — so the check is **Site ID + Activity code**. No Certificate # check (there is no Request # in the PDF).

- The TSR item's catalogue prefix is compared against the PDF Activity code: `itemMatchKey("EX01 - Site Visit")` → `"EX01"` vs `normalizeActivityCode("EX.01")` → `"EX01"`.
- `normalizeActivityCode()` strips all spaces, dots, dashes and underscores: `"EX.01"` / `"EX-01"` / `"EX 01"` → `"EX01"`.
- **Site matching is token-based** (`siteTokens` + `siteMatchLevel`), because the PDF Site ID cell is written two different ways:
  - a single site — `"5227"` (site-visit style certificates)
  - an MW-link pair — `"D7677 - 4637"` (TX/DS style certificates)

  `siteTokens()` splits the cell on every non-alphanumeric run. **Order is significant — the first token is the Site ID, the second is the facing site.** `siteMatchLevel(tsrRow, pdfRow)` returns `{ level, issue }`:

  | Level | Situation | Issue reported |
  |---|---|---|
  | 4 | Site ID is the first token and the facing agrees (or the PDF has no second token) | — |
  | 3 | Site ID is first but the second token ≠ TSR Facing # | Facing # mismatch |
  | 2 | The pair is written the other way round (facing first) | Site ID / Facing # reversed |
  | 1 | Site ID appears but not as the first token | Site ID is not the first entry |
  | 0 | TSR Site ID not in the cell at all | — (no match) |

- Selection: over all PDF rows, `score = (activity code matches ? 100 : 0) + siteMatchLevel`; the highest score wins. The code is the primary check, so a code match outranks a better site level and the row is paired with the line it is really about.
  - Code matches and level 4 → **pass**
  - Code matches but level 1–3 → **fail** with the site/facing issue above
  - Right site but a different code → **Activity Code mismatch** issue (the row still gets a position, so the order check still applies)
  - No site match at all → **Row not found in PDF**
- No de-duplication: the same Site ID never appears twice with the same activity code in one certificate, so each TSR row is matched independently.

**Phase 2 — Bidirectional order check (per folder group):**

After all rows are matched, every row with a valid Excel/PDF position is checked in both directions:
- **Forward check (i→j, j>i):** if a later TSR row maps to an earlier position in the attachment file → row `i` is flagged ("this item comes after X in the file but appears before it in the TSR")
- **Backward check (i→j, j<i):** if an earlier TSR row maps to a later position → row `i` is also flagged ("this item comes before X in the file but appears after it in the TSR")

This ensures **both rows in every swap are flagged**, not just the second one. Uses O(n²) nested loops — acceptable for typical submission sizes.

### PDF Table Parser (`parsePdfToDataRows`)

1. Extracts all text items from all pages with x/y coordinates (PDF y-axis flipped)
2. Groups items into row buckets by y-coordinate (tolerance ±3 pt, per page)
3. Finds header row: must contain `"site"` AND at least one of `"activity"` / `"item"` / `"request"`
4. Classifies each header cell into a logical key:
   - `"site"` → `site`
   - `"activity" + "code"` → `req` (reused for activity code in TOC PDFs)
   - `"activity"` / `"descript"` / `"item"` → `item`
   - `"request"` → `req`
   - `"facing"` → `facing`
   - anything else (Quantity, Area…) → `null` (ignored)
5. For each data row cell, finds nearest header by x-distance and assigns to that column's key
6. Returns array of `{ pos, request, siteId, facing, item }`

### Quantity Check (`buildQtyCheck`)

For the submission's rows on `PO Break Down- Contractor & Acc`, **sent qty = `Qty` × distance factor** (`Distance` column: 0–100 km 1.0, 100–400 km 1.1, 400–800 km 1.2, > 800 km 1.25), summed per item (`itemMatchKey`, so DS07/DS09 sum together). Each item total must be ≤ the item's **Remaining** on `Request Form - VF` (`readRequestForm`: header row found by "item description", Remaining = first header containing `remaining` but not `after`; also reads `Qty` and `Unit Price`). An item over its remaining, or missing from the Request Form, fails. Shown as a collapsible **Quantity Check** section under the summary banner listing **every TSR item** in Request Form order — sent or not (lines with no TSR Qty and no Remaining are skipped unless sent) — with TSR Qty, Remaining, Rows, Qty, Sent (Actual), Balance (= Remaining − Sent), Unit Price, Sent Value, Balance Value, and a totals footer. Unsent items are grey with a **Not sent** badge; sent items missing from the Request Form are appended and fail and counted in the banner; it does not change the per-row Pass/Fail counts. If the `Qty` column or the sheet/Remaining column is missing, the check is skipped with a notice.

Remaining on the Request Form is **before** the current submission, so the rule is sent ≤ Remaining.

**Blank Distance** on PO Break Down is filled in this order: (1) the optional **Master Tracking** upload (`#tsrval-input-trk`, sheet `Invoicing Track`, read by `readTrackingDistances`) via `lookupTrackingDistance` — matched by Site ID, trying Site + Item + Certificate # → Site + Facing + Item → Site + Item → Site + Certificate # → Site + Facing → Site, and a level is used only when all its tracking rows agree on one band; (2) `Total Value ÷ (Unit Price × Qty)` when it equals a band factor; (3) factor 1.0, flagged as unknown. Every filled row is listed under the Quantity Check table with its source.

### Results Display


- Grey diagnostic strip showing detected column letters
- Summary stat boxes: Submission #, Total Rows, Passed, Failed
- Green all-OK banner or red issues-found banner
- Per-folder sections with colour-coded headers (green border = all pass, red = has issues). Each section is a `<details>` — click the header to open/close; folders with issues start open, all-pass folders start closed. An **Expand all / Collapse all** bar sits above the folders.
- TOC folders show a teal **TOC** badge and teal left border
- Per-row table: TSR Row #, Site ID, Facing #, Item Description, Certificate #, Status badge, Issues list
- Order errors report the specific other row they're swapped with and both Excel positions

## TSR Sub Order (`tsr-order-app.js`)

Builds a TSR submission from one **TSR Sub#** in the order the tasks appear in the weekly acceptance mails. All element IDs use the `tso-*` prefix.

**Inputs:** Tracking file (sheet `Invoicing Track`, header row found by scanning for "Logical Site" + "Acceptance Week"), TSR file (sheet `Request Form - VF`, same Item Description / Remaining detection as TSR Sub Prep), a **mails folder**, the **TSR Sub#** (type-or-pick combobox — `<input list>` + `<datalist id="tso-sub-list">` of every distinct value in the tracking column whose header contains `tsr` + `sub`, with task counts; the typed text must match a value exactly, case/space-insensitive via `subKey()`, before Build Submission is enabled; `#tso-sub-hint` shows the task count or how many values contain the text) and a **Target Amount** (`500000`, `500,000`, `500K`, `1.5M`; blank = no limit).

**Mails folder:** one parent folder with `Upper-Cairo`, `Alex` and `Delta` sub-folders. Area comes from the deepest folder name (`alex` → A, `delta` → D, `upper`/`cairo`/`giza` → U), else from the mail name. Weeks = every number after a `W` in the mail name (`… Status W34-2026` → 34; `… 26W03-W04` → 3, 4); year from `20xx` or the 2 digits before `W`. `.msg`, `.eml` and direct Excel files are accepted; the first Excel attachment is read. In it, the sheet with a Site ID + Line Item header (`Site_ID` / `Item_Description` and variants, first 60 rows) is used — the one with the most rows if several.

**Matching:** tracking rows with the selected Sub# (Cancelled skipped). Each Acceptance Week segment (`D-W38-W39-2026`) is mapped to mails of the same area + year that share a week number — an exact week set first, then the most recently saved file. The task is matched by Site ID (`siteKey`) + catalogue code (`itemKey`), count-aware (each mail row used once, same facing preferred).

**Renamed codes (`ITEM_ALIASES`)** — the new price list renamed some codes for the same task, so the tracking and the mail can disagree (`DS07 - Dismantle & Return Microwave Link or Dish` = `DS09 - Dismantle & Return Microwave Link`). `itemKey()` maps an alias to its canonical code (`DS07` → `DS09`), so they match each other; `tsrKey()` also falls back to a same-code match so the item is found in the TSR under either name. Add new pairs to `ITEM_ALIASES`.

**Merged items (`ITEM_GROUPS`)** — items the new price list merged are matched by **description text**, checked before the code (the merged item's new code may be reused for another task or missing): `RF01` Three (3) + `RF02` Four (4) Technologies = `RF01` Up to four (4) → `RF-MACRO-UPTO4`; `RF03` Five (5) + `RF04` More than Five (5) = "Five (5) or more than Five (5) Technologies" → `RF-MACRO-5PLUS`. Any "Macro / Mega Site" text containing five/(5) is 5+, else three/four/(3)/(4) is up-to-4.

**Corrections from the mail (`applyMailFixes`)** — the TSR Sub Validation Excel rules applied up front, with the mail row as the reference: the mail's **Request #** column must equal **Certificate #**, and its **Facing** must equal the tracking Facing. A blank or different tracking value is replaced by the mail value (`NA`, `N/A`, `-` count as blank on both sides). A blank mail Request # never overwrites the Certificate #, but a blank (or `NA`) mail **Facing** — when the mail sheet has a Facing column — **clears** the tracking Facing, including an `NA` ("Facing … removed (blank in mail)"). Each change is recorded in `task.fixes` (`{ field, text }`): shown in a "Corrected from mail" column on screen, a yellow banner with counts, yellow-filled cells in the export and a **Corrected from Mail** column at the end of the `Submission` sheet. Site ID and Line Item need no correction — they are the match key.

**Order:** mails oldest week first (then Alex, Delta, Upper-Cairo), then row order inside the mail.

**Selection — whole mails, whole sites:** the acceptance mail is the unit (optimising the acceptance sheet comes first; the **Target Amount is only a reference**). Mails are walked in mail order; once a mail is started, **every** site in it that fits the TSR is taken, even past the target. The target only decides whether to start the next mail (`total >= target` → the remaining mails are not started). Inside a mail, found tasks are grouped by **Site ID**; a site is taken with every item its mail lists for it, or left out entirely (never a partial site). A site is left out (every task gets a "Whole site left out — …" reason) when:
- the mail lists an item for that site that no task of this Sub# matched — the reason names it and where it is in the tracking (another TSR Sub#, matched to another mail, or not in tracking / Cancelled);
- any of its line items is not in the TSR, or the site's actual qty (`Absolute Qty × distance factor`, summed per TSR item) exceeds the TSR remaining left;
- the target was already reached before its mail was started.

The Mails used list shows "n of N" rows taken per mail and flags a mail that was only partly taken.

Amount = `New Total`. The selected list is re-sorted to mail order and the running total recomputed.

**Mail folders:** every mail holding at least one selected task gets a folder number 1, 2, 3… in mail order (`result.folders` / `folderOf`, `task.folder`).

**Results UI:** the Selected / Not included / Mails used lists are collapsible `<details class="tso-details">` sections, closed by default, with the count and amount in the summary line.

**Export** (`TSR_Sub_<n>_Order_<date>.xlsx`): sheet `Submission` — ID#, VF Task Owner, Vendor, Logical Site ID, Site Option, Facing, Task Date, Line Item, Absolute Quantity, PRQ, Certificate #, Acceptance Status, **Folder**, **Corrected from Mail**, in mail order; sheet `Not Included` — the same 12 columns + Acceptance Week + Reason.

**Download Mail Folders** (`TSR_Sub_<n>_Mails_<date>.zip`): the original mail files placed in folders `1/`, `2/`, `3/`… matching the Folder column. Uses **JSZip** (cdnjs 3.10.1), loaded on demand by script injection like PDF.js — a browser cannot save loose folders, so they are zipped.

## TSR Planner (`tsr-planner-app.js` + `tsr-core.js`)

Plans the **next submission of a TSR** straight from the master tracking. All element IDs use the `tsp-*` prefix. Shared rules live in `tsr-core.js` (`window.TSRCore`) so a rule is written once; the Planner is the first user (the older TSR tabs keep their own copies for now).

**Inputs:** Tracking file (`Invoicing Track`, every column found by header name — `readTracking`), TSR file (.xlsb — `readTsrFile`), acceptance mails folder (Upper-Cairo / Alex / Delta, same reading as TSR Sub Order — `buildMailList`, `parseMailFile`). Settings are pre-filled and editable: **TSR #** (Request Form "Req Form#" cell, else digits in the file name), **Submission #** (last submission in PO Break Down + 1), **Paste at TSR row** (first empty row after the last used Site ID / item row, never above row 10), **Submission Date** (optional, column X), **PO#** (for the receiving file; pre-filled from the tracking `PO number` of tasks whose TSR Sub# starts with the TSR #).

### What `readTsrFile` reads

- **Request Form - VF** — item name (kept **verbatim**, trailing spaces included, because the PO Break Down price lookup `VLOOKUP(N, CQ50:CS125)` needs an exact match), Qty, Unit Price, Remaining.
- **PO Break Down- Contractor & Acc** — header row = the row holding both `Site ID` and `Item Description` (row 8 in VF's template, data from row 10). Read cell by cell (`cellV`), never `sheet_to_json`, because the sheet carries a ~100 000-row hidden site list. Existing rows → submissions present, last used row, per-row qty and the Distance shown (`#N/A` error cells → factor null).
- **Hidden site list** — the `Site Id | Physical | Region | Governorate | Site Type | Distance` block (columns CV:DA) that the Region / Governorate / Distance / Physical Site formulas look up. Keyed by `siteListKey()`, which mimics `VLOOKUP(…, FALSE)`: case-insensitive, space-sensitive, and **numbers only match numbers** (`3618` ≠ `"3618"`).
- **Dropdown lists** at fixed addresses, each sanity-checked: vendors `CO58:CO61`, General Stream `CO63:CO155`; price catalogue `CQ:CS`.
- **Receiving** — `Submission #n` columns; a submission whose total (row below the header) is > 0 has been **received** by VF.
- **Receiving Template** — VF's catalogue for the receiving file: Vendor Name, Vendor ID, Description, Unit Cost, Line Row, EVO Mat. No., Stream Name.

### Planning rules (`runPlan`)

1. **Live tasks** — `statusClass()`: Cancelled (+ misspellings Cancalled / Cancel), Withdraw, Repeated, Duplicated, Closed, REJ are dead; `Done` is done; anything else (Assigned, Hold, blank) is open.
2. **Already submitted** — a task matched to a row already in this TSR file (Site + item key, Certificate # preferred), or with a PO status — **except PO status `ok`** when the task is tagged with this TSR # (`1326`, `1326 sub-2`, `check 1326`) or not tagged at all: in the tracking `ok` means "picked for a submission, not received yet", and for this TSR the loaded TSR file is the truth, so such a task stays available. `Received`, `REJ`, other statuses and other TSRs' tags exclude. **Planned in another TSR** — TSR Sub# filled with another TSR # (and not submitted). (Before this rule, running on a TSR file without Sub-2 skipped every Sub-2 task marked `ok` and planned a different batch on top of it.)
3. **TOC track** — EX01 / EX02, Acceptance Status `TOC`, or week starting `TOC` are kept out of the combos and listed separately (planned later, separate TOC mails).
4. **Combo = Job Code + Site** (FAC tasks only). Skipped: fully submitted, no item in this TSR, USO comment. A combo is **ready** only if it has no reason: no task planned in another TSR, no part already submitted, every task Done + FAC date + readable acceptance week + item in the TSR + a known distance + found in its acceptance mail (count-aware, Certificate # then Facing preferred; mail rows are claimed oldest combo first). Corrections from the mail (`mailFixes`) are the same rules as TSR Sub Order.
5. **Distance used for quantity** — VF's distance from the hidden site list when the site is listed (that is what the TSR formula uses), else the tracking distance (VF fills `#N/A` sites after the file is sent). `t.effQty = Absolute Qty × that factor`.
6. **Available quantity** = Request Form Remaining − rows of submissions in the file that VF has **not received yet** (Remaining only drops on receiving). Each such row uses the Distance the TSR shows, else the matched task's tracking distance.
7. **Allocation** — ready combos sorted by acceptance week (`weekSortKey`), label, first row; each whole combo is taken if every item fits, else it goes to **Waiting for qty** with the shortfall. Rows are output in mail order; mails holding planned rows are numbered folders 1, 2, 3…
8. **General Stream** — per combo, from the main item (`STREAM_RULES`: RF macro → New Physical Sites, TX04 → Upgrade-MW Link E-Band 2+0, TX03 → E-Band 1+0, TX05 / TX19 → XPIC 2+0, TX09 → IDU, TX16 → Additional Cards with Cables, RF09 → Add Sector, PW02 → Add Batteries & Rectifiers, TX10–15 → Upgrade-TX Cutover). A TX03/04/05 combo on a site that also has an RF macro task is flagged (New Physical Sites suggested). Valid values = the TSR dropdown + the Receiving Template stream names (VF uses `Upgrade-MW Link E-Band 2+0`, which is missing from the dropdown). Each combo has a `<select>` in the results; a choice overrides the rule in the exports (`_streamOverride`).

### Exports

- **TSR Paste File** (`TSR_<no>_Sub-<n>_Paste_<date>.xlsx`) — sheet **TSR Paste** uses the **same column letters as PO Break Down** (row 3 = the TSR headers, data from row 4). The sheet is protected and Excel refuses a paste touching a locked cell, so the editable columns are 5 blocks — **C:D, H:J, M:N, P, S:Y** — each labelled in row 2 with its exact copy range and target cell (`paste at C421`). Values: C owner, D vendor (dropdown spelling), H Site ID (trimmed, numeric IDs as numbers), I option, J facing (mail-corrected), M task date, N Request Form item name, P absolute qty, S stream, T PRQ, U Certificate # (mail-corrected), V `FAC`, W `Submission #n`, X submission date, Y folder #. The formula columns E F G K L O Q R are grey **previews** of what the TSR will calculate (red `#N/A` = site not in VF's list) — never pasted. Yellow = corrected from mail, pink stream = needs a check. Other sheets: How to paste, **Allocation** (tracking row → `"<TSR> sub-<n>"` to write in the TSR Sub# column), Qty Check (formula balance), Waiting for Qty, Not Ready, In Another TSR, TOC Track (later).
- **Receiving Template** (`Receiving Template LMP TSR <no> Sub-<n>.xlsx`) — LMP's layout (sheet `Receiving`, Aptos Narrow 11, header row 2 filled `#C0E6F5`, Grand Total row, `I = E*H`): one row per Receiving Template item with qty > 0; QTY = Σ effQty of the planned rows; **Line Row = 1 and PO unit price = 1** (the user's convention, confirmed); PO# from the setting.
- **Mail Folders** (`TSR_<no>_Sub-<n>_Mails_<date>.zip`) — same as TSR Sub Order.

### Shared helpers in `tsr-core.js` (Planner + New PO)

`matchExisting(tsr, tasks, tag)` → `{ inTsr: Map(task → submission #), pairs: [{ex, task}], unmatched }` (Site + item key, Certificate # preferred, then a task tagged with this TSR). `tsrAvailable(tsr, pairs)` → `{ avail, pendingUse, pendingSubs }` (Remaining − rows of submissions not yet received). `tagTest(tag)` (TSR Sub# tagged with this TSR: `1326`, `1326 sub-2`, `check 1326`), `tsrSite(t)`. `readTsrFile` items also carry `rfRow`, `unit`, `matGroup`, `evo`, `pa`, `paItem`; tracking tasks carry `price` (`New Price`).

### Plan a New PO (`tsr-newpo-app.js`, IDs `tnp-*`)

A card under the Planner results. Needs only the tracking and the **newest TSR file** loaded in the Planner (no mails); the planner fires a `tsp-files` event after each file load and exposes `window.TSRPlanner = { trk, tsr }`. Input: **New PO budget** (default `11,990,000` = the 12M PO limit − 10K; accepts `11.99M`, `500K`). TSR # comes from the Planner's TSR # box.

1. **Sent** — tasks matched to rows of the TSR file (`matchExisting`); PO status rules as the Planner (`ok` + this TSR's tag or no tag = not sent). Tasks tagged with another TSR (not submitted) are excluded and listed.
2. **Combos** (Job Code + Site) of the remaining live tasks; **TOC lines (EX01/EX02, TOC status/week) ride with their combo** and only need to be Done. USO combos are skipped. Priority `tier`: **1** all Done + FAC + week, **2** all Done + FAC, **3** all Done, **4** some line not Done but a Done line anywhere in the combo (sent or not) — the open lines go with it; **0** nothing Done → *Not started*, left out. Split combos (part already sent / in another TSR) are included by their tier and noted. Order: tier → acceptance week → earliest FAC → earliest task date → first row.
3. **Stays in the current TSR** — combos with a task in the TSR file or tagged with this TSR, any tier, oldest first: stay if every item is on its Request Form and fits `tsrAvailable` (exact qty); otherwise **Moved** into the new-PO pool.
4. **New PO** — the rest, greedy in order, whole combos; a combo that would pass the budget is skipped (**Over budget**) and the loop continues. Qty per task = Absolute Qty × distance factor (VF site list in the TSR, else tracking, else 1.0 flagged). Request Form quantities are whole numbers, so the budget is checked on Σ ceil(qty per item) × TSR unit price. Items not on the TSR price list use the tracking price (`New Price`, else New Total ÷ qty) and are flagged red.

**Exports:** **New PO File** (`New_PO_Request_<date>.xlsx`) — `Request Form` (every TSR item in Request Form order with Material Group / EVO / PA / PA Item, Qty rounded up, Unit, Unit price, Total formula, Exact Qty, Tasks; non-catalogue items appended red; Grand Total, Budget, Left), `New PO Tasks` (ID# first, write `New PO` in the tracking), `Stays in TSR <no>`, `Qty Check TSR <no>`, `Sent in TSR <no>`, `Over Budget`, `Not Started`, `In Another TSR`, `USO (excluded)`. **Sent Tasks (ID list)** (`TSR_<no>_Sent_Tasks_<date>.xlsx`) — every TSR-file row matched to the tracking: ID#, tracking row, Site, item, TSR row, Submission, the TSR Sub# to write (`1326 sub-2`), the value in the tracking now, `Update?` = Yes (yellow) when they differ; unmatched TSR rows listed in red at the end.

Note: the tracking `New Total` is priced lower than the TSR price list for some items (e.g. TX13 4,400 vs 5,000); the budget uses the **TSR** unit prices, which is what the Request Form totals.

### Verified against TSR 1326 (Nov 2026)

Pasting the blocks into a copy of the real TSR and evaluating VF's lookups: every price and distance lookup resolved and the grey previews equalled the TSR results; receiving quantities equalled the quantities the TSR computes; Activity Dates matched the tracking in both UTC and Africa/Cairo.

`excelDate()` in `tsr-core.js` rounds to the nearest minute and writes the calendar day as UTC midnight — SheetJS can return a date a few seconds **before** midnight (16-Aug read as 15-Aug 23:59:51), which the older `getTimezoneOffset()` shift used by other tabs does not correct.

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

## Acceptance Check (`acceptance-check-app.js`)

Inputs: Tracking file (sheet `Invoicing Track`, header row found by scanning the first 30 rows for "Logical Site" + "Acceptance Week"), Acceptance sheet (tab `Total`, header row found anywhere by scanning for "Site_ID" + "Item_Description" — the table does not start at A1), an **Area** (`Delta` / `Cairo-Giza-Upper` / `Alex`), one or more **Weeks** and a **Year** (defaults to the current year). All element IDs use the `acc-*` prefix.

Both files are parsed as soon as they load (`_trk` / `_acc`), so the **week picker** can list every week found in either file for the selected area/year as clickable chips (`#acc-week-chips`). Chips and the Weeks text box (`36, 37`) stay in sync both ways; a dashed chip means that week exists in only one of the two files. A row is in scope if it shares **any** week with the selection.

**Scope filtering:**
- Tracking: `Acceptance Week` is split on `/` into segments like `D-W36-W37-2026`; a row is in scope if any segment has the area prefix (`D` Delta, `U` Cairo-Giza-Upper, `A` Alex), the selected year (or no year), and shares a week number with the selection. Cancelled rows are skipped.
- Acceptance: `Area` starts with `delta` / `cairo`|`giza`|`upper` / `alex`, and `Week` (e.g. `34`, `36--37`, `36—37`; a cell Excel turned into a date like 3-Apr reads as weeks 3 and 4) shares a week number with the selection.

**Matching** (per Site ID — `Logical Site ID` vs `Site_ID`, uppercased, leading zeros stripped from numeric IDs):
- Line items compare by catalogue code (`TX03 - …` → `TX03`, `EX6` → `EX06`); items with no code compare by full text. Each acceptance row matches at most one tracking row (count-aware).
- Leftovers are paired as **Line item mismatch**, preferring the same code family (TX↔TX), then the same facing. Unpaired leftovers are **Missing in acceptance** / **Missing in tracking**; a site present on only one side is **Site not in acceptance** / **Site not in tracking**.
- Notes point to the same site + item elsewhere (other week/area) in the other file, and to near-identical Site IDs (same digits, e.g. `3844` vs `U3844`), to help spot a wrong week or site rather than a wrong item.

Results show the tracking Excel row number for every issue so it can be corrected in the master tracking. Matched rows are hidden unless "Show matched rows" is ticked.

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

When updating any cached file, bump the `CACHE` version string in `sw.js` (e.g. `lmp-invoicing-v24` → `lmp-invoicing-v25`). Without this, installed PWA users will continue running stale files. Current version: `lmp-invoicing-v53`.
