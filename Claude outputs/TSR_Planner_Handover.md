# Handover — TSR Planner & LMP Invoicing System (11 Nov 2026)

## Context
- Khaled (Operations Manager, LMP Telecom Dept) prepares TSR submissions ("subs") for Vodafone POs.
- The work is automated in the **LMP Invoicing System** PWA:
  - Static, no build step; hosted on GitHub Pages.
  - Repo: `landmarkplus2016-byte/lmp-invoicing-system`, branch `feature/tsr-planner`.
  - Local folder: `H:\My Drive\Claude Projects\Visual Studio\Telecom Department\LMP Invoicing System`.
- The repo's `CLAUDE.md` is the full technical reference. Read it first.

## Workflow rules
- **Khaled commits and pushes from VS Code himself.** Claude has no GitHub push access (403). Don't try to push; write the changed files into his local folder:
  - CRLF line endings.
  - Check the file's current mtime first, so nothing he changed is overwritten.
- Bump the `CACHE` version in `sw.js` on every change. It is currently **`lmp-invoicing-v52`**.
- Browser: Chrome. He saves outputs into folders himself.

## Business rules (confirmed by Khaled)
- LMP requests the TSR (Request Form) first, then VF creates the PO. Quantities are fixed until the end of the PO.
- **Whole site = Site ID + Job Code combo.** All its line items go in the same PO-TSR, or none of them do.
- A combo can be submitted only if:
  - every task is Done;
  - every line item has a FAC date and an acceptance week;
  - every item is in the TSR;
  - it fits the TSR quantity.
- Never exceed the TSR quantity:
  - Actual qty = Absolute Qty × distance factor (0–100 km 1.0, 100–400 km 1.1, 400–800 km 1.2, >800 km 1.25).
  - Available = Request Form Remaining − submissions in the file that VF has not received yet. Remaining drops only on receiving.
- **Oldest acceptance first.**
- EX01 / EX02 and some other items are **TOC**, with separate mails. They are planned separately, later.
- Tracking (`Invoicing Track`) is the source for VF Task Owner, Activity Date and Option.
- Distance: the TSR computes it from a hidden site list (CV:DA). For sites missing from that list (#N/A), VF fills it after the file is sent. Until then, use the tracking distance.
- Corrections from the mail: Certificate # = mail Request #, Facing = mail Facing.
- Receiving template: Line Row = 1, PO unit price = 1, QTY = Σ actual qty per item.
- In the tracking, PO status `ok` = picked for a submission, not received yet.

## App pieces built in this project
- **TSR Planner tab** (`tsr-planner-app.js` + shared `tsr-core.js`).
  - Plans the next sub from the tracking, the TSR .xlsb and the acceptance mails folder.
  - Exports:
    - TSR Paste File: editable blocks C:D, H:J, M:N, P, S:Y, with grey previews of the formula columns.
    - Receiving Template.
    - Mail-folder zip.
    - Allocation list.
  - Always load the **newest TSR file** (the one holding every submission already sent). The Planner plans the submission after it.
- **My Files tab + ⚡ Load My Files** (`file-sources.js`):
  - Remembers file and folder locations and loads every tab in one click.
  - A wrong-file guard checks that each file has the sheets it should.
- **TSR Sub Validation**, fix-33: the banner now counts only failed quantity items. Before, it also counted unsent items (that was the false "7 quantity issues").

## Current status — TSR 1326
- **Sub-2 is final and validated:** 179 rows, all pass; quantities all fit (TX11 4.5/4.5, EX09 27.4/27.6, TX12 9.2/9.4).
- **Receiving Template LMP TSR 1326 Sub-2.xlsx** has been made:
  - 15 items, PO 3006420072, Grand Total ≈ 2,473,856.
- 19 Sub-2 rows have #N/A distance in the TSR. Tracking distance was used for them.
  - Factor 1.1 / 1.2 rows: FA9380, FA7755, FA7165 (TX13), Absheway (TX16), U0323, D4905, X0004, D0471.
  - If VF fills a different distance band for these, update the receiving file.
- Tracking data issues seen in the comparison:
  - MA0481 / 3321 duplicate Assigned rows (11345–11349).
  - Missing FAC dates: H3001, HR1294 TX09, U0323 TX05, K3660, D4905.
  - Split combos: U3983, 3578, HR1294 TX11, D1596.
- Commits fix-29 to fix-33 are local only, waiting for Khaled to push from VS Code.

## Likely next steps
- Plan **Sub-3**: run the Planner with the TSR1326 Sub-2 file. The Submission # box should show 3.
- Plan the **TOC track** (EX01 / EX02) once the TOC mails folder is shared.
- Optionally move TSR Sub Prep / Sub Order / Validation onto the shared `tsr-core.js` rules.
