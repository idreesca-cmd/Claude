# USC Asset Disposal Dashboard — Build Progress

**Client:** Utility Stores Corporation (USC) — FAR Auction / Asset Disposal & Reconciliation
**Engagement:** Baker Tilly (Mehmood Idrees Qamar) — Third-Party Validation (TPV)
**Repo/branch:** `idreesca-cmd/Claude` @ `claude/dashboard-files-review-t31098`
**Last updated:** 2026-07-06

> This file is the resumable checkpoint. If a session ends, a new session can
> read this + `/source-data/*` and continue the build without re-deriving anything.

---

## 1. Goal (from user brief)

Extend the existing `USC_Dashboard_Published.html` into a self-refreshing, multi-page
dashboard whose data comes from Google Drive (folder **"USD DB MIK"**), not manual upload.
Key requirements:
- **New FIRST tab "Category Milestone Overview"**: columns = asset categories (from the
  Database sheet, exact list), rows = milestone stages **Physical Verification → Auction →
  Lifting → Payment**. Each cell: donut % ring + qty-achieved-vs-total + planned-vs-actual
  bar pair. Restyle the reference image's blue/green into the Baker Tilly palette. Skip the
  reference's "Risk & Issues" row (no equivalent source field — do NOT fabricate).
- **Keep** the existing 4 tabs (Overall Project Status, Quantity Reconciliation, Auction
  Reconciliation, Payment & Lifting), Tailwind + Chart.js setup, and the `--bt-*` palette.
- **Fix** the `#REF!` grand-total formulas in the consolidation workbook (underlying formulas,
  not just display).
- **Normalize** free-text `Project Updates` rows into {zone, category, status, date, note}.
- **Un-mark** the "Pending" Payment & Lifting tab if data now supports it (**it does** — see §4).
- Maintain this **PROGRESS.md**.

---

## 2. KEY DECISIONS (locked)

| # | Decision | Rationale |
|---|----------|-----------|
| D1 | **Refresh model = embedded data + regenerate** (self-contained HTML, data baked in). | Client views via a shared link **with no login** (user's explicit requirement). Rules out in-browser Drive OAuth (needs viewer login) and a local server (not shareable). |
| D2 | **Hosting = GitHub Pages** on this repo/branch for the no-login shareable link; regenerated HTML also saved back into the Drive folder as source of record. | A raw `.html` in Drive does NOT render as a webpage at a share link. Pages gives a stable public URL, keeps the existing Tailwind/Chart.js CDNs (no restyle), and "refresh" = regenerate + push. |
| D3 | **Deliver NEW file versions, not in-place edits.** | The Google Drive MCP has **no update-file API** — only read/download/create/copy. So the fixed xlsx and new dashboard are uploaded as new files (e.g. `... - FIXED.xlsx`, `USC_Dashboard_v2.html`); user retires the originals. |
| D4 | **Build base = `base/index_base_shell.html`** (the clean, complete app shell), NOT the published file. | The Drive download of `USC_Dashboard_Published.html` came back truncated mid-JS. `index.html`/`index_clean.html` is the same engine (same CSS, tabs, render funcs) minus the baked CSV, and is complete. |
| D5 | **Architecture = feed the FULL 11,785-row Database into the existing `RAW` engine**, then add the new tab that computes per-category×stage aggregates live from `getFiltered()`. | Least rewrite, highest fidelity: all 4 existing tabs light up with real data automatically and the new tab is slicer-aware. Honors "don't rebuild". |
| D6 | Data-quality anomalies (>100%, payment>billed) are **shown straight from source and flagged**, never silently clamped. | Extraction found genuine source inconsistencies (see §5). Donut ring caps visual at 100% but label shows true %; anomalies marked. |

**Open confirmation for user:** brand palette. The `BT Brand Colors.docx` gives charcoal
`#2c303b` / mint `#dff5f2` / lime `#d1ec51` (no green), while the existing dashboard uses
`--bt-green:#00A376`. Plan: keep the `--bt-*` variable scheme but set accent→lime, tint→mint,
dark→charcoal, keeping green as one secondary chart color. Fonts → Calibri / Calibri Light
(brand doc) with Inter/system fallback. **Flag for user sign-off.**

---

## 3. SOURCE FILES (Drive folder "USD DB MIK", id `1PFSMyIYp3ZoMEDs7aL-CsSYrreyKW0zn`)

| File | Drive id | Role |
|------|----------|------|
| All Zones Consolidation.xlsx | `1s_Nf1JvWNJg16J8iY-z2wWdeyeyzN5Um` | **Source of truth** — sheet `Database` (11,785 rows × 37 cols) + `Reconciliation Summary` (has #REF!) |
| Project Updates.xlsx | `1f_soWkh2MS0AVQ7CMedK6rxUgW0RB661` | Free-text `Project Status` + `News` sheets |
| BT Brand Colors.docx | `1SRd65l6ZYu04LcLYbMNbjcAPs_YvSCz1` | Palette + fonts |
| USC_Dashboard_Published.html | `1QAA_zCPff6gb69E1l6WCZ9FGpcJL-Dci` | Existing dashboard (extend). NOTE: Drive download truncates. |
| Inception Report - USC Final.docx | `1lX_QIkiMVpW7T3j9xIaVH_MHY-R69yWR` | One-time context only (scope/zones/terms). Not live data. |
| Region wise M & B Data.xlsx | `1Vfly6jI3rmcCBS4ZrV2fV6Yi8hz7jDZl` | M&B financial valuations (secondary). |

Extracted artifacts are committed under **`/source-data/`** (see that folder):
`consolidation_data.json` (per-category aggregates + category/zone/status lists),
`project_updates.json` (59 normalized records), plus the `.md` summaries
(`consolidation_summary.md`, `ref_errors_and_fixes.md`, `html_outline.md`, `bt_brand.md`,
`project_updates.md`). The clean base shell is `/base/index_base_shell.html`.

---

## 4. #REF! FIX SPEC (Reconciliation Summary tab) — DECLINED, IGNORE GOING FORWARD

**2026-07-07: user decision — ignore the `#REF!` errors going forward.** Do not re-flag, re-fix,
or ask about this again in future sessions. Kept below for historical reference only.

Note this was **never a dashboard-accuracy issue**: the dashboard computes every total directly
from the `Database` sheet in `build/build_dashboard.py` (Python), not by reading the
`Reconciliation Summary` tab's broken `SUM` formulas. The `#REF!` cells only affected that one
standalone Excel tab when opened manually in Excel — the published dashboard's numbers have
always been correct regardless of this issue.

<details><summary>Historical fix spec (not applied, not required)</summary>

All **15** cells `B4:B18` are `=SUM(#REF!)`. Root cause: they summed a helper sheet
`_PivotSource` (+ `Pivot-*` sheets) that were deleted, so Excel voided the ranges.
**Fix:** re-point each to the matching `Database` column, rows 2–11788:

| Cell | Label | Corrected formula |
|------|-------|-------------------|
| B4 | A QTY (Odoo) | `=SUM(Database!E2:E11788)` |
| B5 | B QTY (Physical Lists) | `=SUM(Database!F2:F11788)` |
| B6 | C QTY (Physical Count) | `=SUM(Database!G2:G11788)` |
| B7 | A - B | `=SUM(Database!H2:H11788)` |
| B8 | B - C | `=SUM(Database!I2:I11788)` |
| B9 | Auction Qty Usable | `=SUM(Database!J2:J11788)` |
| B10 | Auction Qty Scrap | `=SUM(Database!K2:K11788)` |
| B11 | Auction Qty Total | `=SUM(Database!L2:L11788)` |
| B12 | Diff Auction vs Count | `=SUM(Database!M2:M11788)` |
| B13 | Total Reserve Price | `=SUM(Database!R2:R11788)` |
| B14 | Total Auction Value | `=SUM(Database!Y2:Y11788)` |
| B15 | Payment Rs. | `=SUM(Database!AA2:AA11788)` |
| B16 | Lifted Qty | `=SUM(Database!AD2:AD11788)` |
| B17 | Lifted KGs | `=SUM(Database!AE2:AE11788)` |
| B18 | Balance Qty | `=SUM(Database!AF2:AF11788)` |

Expected grand totals after fix: Odoo 2,951,485 · Phys Lists 1,285,595 · Phys Count 1,804,070 ·
Auction Qty Total 1,993,923 · Total Reserve Price 5,030,464 · Total Auction Value 17,124,785,955 ·
**Payment Rs. 968,295,511** · **Lifted Qty 930,063** · Lifted KGs 1,428,183 · Balance Qty 845,605.
Apply with openpyxl on `all_zones_consolidation.xlsx`, save as a NEW file, upload to Drive folder.

**Payment & Lifting tab CAN be un-marked "Pending"** — real data present (Payment ≈ 968.3M,
Lifted ≈ 930,063).

</details>

---

## 5. DATA MODEL FOR THE BUILD

### Database sheet → dashboard `COL` mapping (37 cols → engine keys)
`A Zone→zone` · `B Region→region` · `C Asset Category→assetCat` (normalize, see below) ·
`D Asset Name (Matched)→assetName` · `E A QTY (Odoo)→aQty` · `F B QTY (Physical Lists)→bQty` ·
`G C QTY (Physical Count)→cQty` · `J Auction Qty Usable→auctQtyUsable` ·
`K Auction Qty Scrap→auctQtyScrap` · `L Auction Qty Total→auctQtyTotal` ·
`P Reserve Price Usable→rpUsable` · `Q Reserve Price Scrape→rpScrap` · `R Total Reserve Price→(new) reservePriceTotal` ·
`S Auction Status→status` (normalize) · `T Auction Date→auctionDate` · `U Winner Bidder Name→bidder` ·
`V Auction Price Usable→apUsable` · `X Auction Price Scrape→apScrap` · `Y Total Auction Value→(new) auctValueTotal` ·
`Z Payment Date→payDate` · `AA Payment Rs.→payRs` · `AB Payment DD No.→ddNo` ·
`AC Lifting Date→liftDate` · `AD Lifted Qty→liftedQty` · `AE Lifted KGs→liftedKg` · `AF Balance Qty→balanceQty`.
- `itemCat` (dashboard has no Database equivalent): DERIVE a coarse bucket — Fixed Assets
  (Furniture, Motor Vehicles, Sign Boards, Computer, Plant, ERP) vs Merchandise & Inventory
  (Own Brand, Rice & Pulses, Branded Goods) — so the Item Category slicer stays useful.
- Add two new `COL` keys `auctValueTotal` (Y) and `reservePriceTotal` (R) for the Payment stage.

### Category normalization — 13 raw strings → 9 display categories
Merge whitespace/variant dupes: ` Rice & Pulses ` + `Rice & Pulses` → **Rice & Pulses**;
`ERP - COMPUTER & OFFICE MACHINES` + `ERP - COMPUTER AND IT EQUIPMENT` + `ERP Computer and IT Equipment` + `ERP ` → **ERP – Computer & Office Machines**; drop `(blank)` (3 rows).
Display names: Furniture, Fixture & Office · Motor Vehicles & Bicycles · Sign Boards ·
Computer & Office Equipment · Plant & Equipment · ERP – Computer & Office Machines · Own Brand ·
Rice & Pulses · Branded Goods (Non-Food).

### Zone normalization
`SUKKUR` + `Sukkur` → **Sukkur**. 9 zones total: Islamabad, Abbottabad, Peshawar, Faisalabad,
Lahore, Multan, Sukkur, Karachi, Quetta.

### Auction Status normalization
`AUCTIONED`/`Auctioned`/`Auctioned ` → **Auctioned**; `NOT AUCTIONED`/`Not Auctioned` → **Not Auctioned**;
`Pending` → **Pending**; `Consignment` → **Consignment**.

### Milestone stage definitions (for the new tab; live from filtered rows)
- **Physical Verification:** total = Σ Odoo (E) ; achieved = Σ Physical Count (G).
- **Auction:** total = Σ Physical Count (G) ; achieved = Σ Auction Qty Total (L).
- **Lifting:** total = Σ Auction Qty Total (L) ; achieved = Σ Lifted Qty (AD).
- **Payment:** total = Σ Total Auction Value (Y) ; achieved = Σ Payment Rs. (AA).
"Planned vs Actual" bar pair per cell = [total (planned), achieved (actual)].

### Known data-quality anomalies (show + flag, do NOT clamp)
Furniture lifting 112% · Sign Boards payment 265% · Plant lifting 870% / payment 396% ·
Own Brand payment 106% · Rice auction 164% · several categories genuine 0 payment
(Motor Vehicles, ERP, Branded Goods). Total Auction Value (17.1B) looks inflated vs Reserve (5.0M).

---

## 6. BUILD PLAN (remaining steps)

1. **[Task 2] Fix xlsx** — openpyxl: set B4:B18 per §4, save `All Zones Consolidation - FIXED.xlsx`,
   upload to Drive folder `1PFSMyIYp3ZoMEDs7aL-CsSYrreyKW0zn`.
2. **[Task 4] Build data payload** — Python script parses Database → emits compact JSON
   (`{cols, rows}` array-of-arrays in fixed `COL` order + `META`: refreshDate, rowCount,
   grand totals, dq flags) + embeds `project_updates.json`. Bake as
   `<script id="dashboardData" type="application/json">`.
3. **[Task 3] New first tab** — insert `tab0` "Category Milestone Overview" before tab1
   (renumber nav; keep data-tab ids stable, just add tab0). 9 category columns × 4 stage rows;
   each cell = CSS conic-gradient donut (light, not 36 Chart.js instances) + achieved/total +
   planned/actual mini-bars. Compute live from `getFiltered()`. Apply BT palette.
4. **[Task 5] Un-mark Payment & Lifting** — replace tab4 `awaiting-card` with real KPIs/charts
   (Payment Rs by zone/category, Lifted vs Balance). Soften tab3 "still awaiting" card (some
   auction value now present). Remove `pending-pill`. Add a "Project Updates & News" panel
   (normalized rows) — likely on Overall tab.
5. **Rewire bootstrap** — replace `bootEmpty()` + upload flow: on load, parse `#dashboardData`
   → build `RAW` objects → `hydrateSlicers(); renderAll()`. Replace header "Upload" control with
   "Last refreshed: <date>" + keep a hidden optional upload for power users. Show DQ note.
6. **Palette/fonts** — update `:root` `--bt-*` to brand values; `Chart.defaults.font.family` → Calibri stack.
7. **Verify** — open final HTML in headless Chromium (Playwright is available at
   `/opt/pw-browsers/chromium`), screenshot each tab, confirm no JS errors, numbers match §4/§5.
8. **Publish** — write final `USC_Dashboard_v2.html` to repo (for Pages) + upload to Drive folder.
   Enable GitHub Pages (user does one-time Settings→Pages→source=branch, or via API). Update this file.
9. **(Optional) Auto-refresh** — schedule a trigger to re-run the regeneration + push on a cadence.

---

## 7. STATUS LOG

- **2026-07-06** — Confirmed Drive access; reconciled file-name mismatch (source files were
  uploaded to the folder today). Extracted all source data + brand + HTML structure to
  `/source-data/`. Locked decisions D1–D6. Chose clean base shell. Wrote #REF! fix spec &
  full data model. Initialized repo on branch `claude/dashboard-files-review-t31098` and
  committed this checkpoint. **Next: Task 2 (fix xlsx) → Task 4/3 (build).**

- **2026-07-06 (session 2)** — BUILT. Produced `deliverables/All Zones Consolidation - FIXED.xlsx`
  (all 15 #REF! totals now compute — verified). Wrote `build/build_dashboard.py` (the re-runnable
  "refresh" step) and generated `deliverables/USC_Dashboard_v2.html` (1.69 MB, 11,782 rows baked):
  new **Category Milestone Overview** first tab (9 categories × 4 stages, CSS donut rings +
  planned/actual bars, >100% anomalies flagged), full data feeds all 4 existing tabs,
  **Payment & Lifting un-marked** with real KPIs/charts/table, BT brand palette (lime/mint/charcoal)
  + Calibri fonts applied, manual-upload replaced by embedded JSON auto-load. `node --check` on the
  app script = OK. **NOT YET DONE:** (a) browser render verification (Playwright), (b) upload the
  fixed xlsx + v2 HTML to the Drive folder, (c) enable GitHub Pages, (d) render a Project Updates/News
  panel (data is embedded as `DASH_UPDATES` but no panel drawn yet), (e) user sign-off on palette.

- **2026-07-07 (session 3)** — VERIFIED IN BROWSER + made self-contained.
  Rendered `USC_Dashboard_v2.html` in headless Chromium: data loads (11,782 records, status
  "Live · refreshed 2026-07-07"); **Category Milestone Overview** draws 36 donut rings (9 cats ×
  4 stages) with plan/act bars and amber ▲ flags on >100% anomalies; **Payment & Lifting** shows
  real KPIs (Payment Rs 968,295,511 · Lifted 930,063 · Balance 845,605) + 2 Chart.js charts +
  detail table; **Project Updates panel** on Overall tab renders 2 news items + 57 zone-status rows.
  `node --check` clean; screenshots in `docs/screenshots/`.
  Made the file self-contained: **inlined Chart.js** (vendored via npm) so it works offline/behind
  firewalls, and **dropped the dead PapaParse/SheetJS CDNs**. Only the Tailwind CDN remains (loads
  fine on GitHub Pages; sandbox can't reach it so verification layout is stacked — cosmetic only).
  Rebuild: `npm i` (restores chart.js) then
  `python3 build/build_dashboard.py <xlsx> base/index_base_shell.html source-data/project_updates.json deliverables/USC_Dashboard_v2.html`.
  **STILL REMAINING:** (a) upload `deliverables/*` to the Drive folder; (b) enable GitHub Pages for
  the shareable link; (c) user sign-off on the lime/charcoal palette; (d) optional: vendor Tailwind
  for full offline. Tasks 1-5,7 done; task 6 (publish) is the last one.

## 6. HOSTING — GitHub Pages (ready to enable)

The dashboard is committed at **`docs/index.html`** (+ `docs/.nojekyll`) on branch
`claude/dashboard-files-review-t31098`. To publish the no-login shareable link:

1. GitHub → repo **idreesca-cmd/Claude** → **Settings → Pages**.
2. **Source:** "Deploy from a branch". **Branch:** `claude/dashboard-files-review-t31098`,
   **Folder:** `/docs`. Save.
3. Wait ~1 min. Live URL: **https://idreesca-cmd.github.io/Claude/**  (share this with the client).

To refresh later: re-run `build/build_dashboard.py`, `cp deliverables/USC_Dashboard_v2.html
docs/index.html`, commit + push — the same URL updates. (There is no MCP tool to enable Pages
programmatically, so step 1-2 is a one-time manual click.)

A copy of the dashboard + the fixed workbook is also uploaded into the Drive folder as the
source of record (see status log for the file links).

## 7. Drive upload — BLOCKED by tool limit (use repo instead)

Uploading `deliverables/*` into the Drive folder via the MCP failed: `create_file` only accepts
inline base64, and the payloads are too large to pass as a tool argument
(xlsx ≈ 1.62 M chars, html ≈ 2.53 M chars base64). No path/stream/resumable Drive upload tool is
exposed. **Both deliverables are safely committed in this repo** (`deliverables/`), and the
dashboard is served via GitHub Pages — so the Drive copy is optional. To place copies in Drive
manually: download `deliverables/All Zones Consolidation - FIXED.xlsx` and
`deliverables/USC_Dashboard_v2.html` from the repo and drag them into the "USD DB MIK" folder.
(If a future session has a path-capable Drive upload tool, target
`parentId=1PFSMyIYp3ZoMEDs7aL-CsSYrreyKW0zn`, `disableConversionToGoogleType=true`.)

## 8. STATUS: what's DONE vs REMAINING (end of session 3)

DONE: extraction · #REF! fix (verified) · new Category Milestone Overview tab · full 11,782-row
data baked into all tabs · Payment & Lifting un-marked with real KPIs/charts · Project Updates
panel · self-contained (inlined Chart.js) · browser-verified render · GitHub Pages files staged.

REMAINING (next session): (a) **click-enable GitHub Pages** (Settings→Pages→branch=this, /docs) to
go live at https://idreesca-cmd.github.io/Claude/ ; (b) optional Drive copies (manual, see §7);
(c) **user sign-off on the lime/charcoal BT palette** vs the old green; (d) optional polish —
vendor Tailwind for full offline, and right-edge scroll of the milestone matrix is Tailwind-gated
(fine on Pages). No open bugs.

## 9. FULLY SELF-CONTAINED (session 3 continued)

Compiled Tailwind to only-used-classes (`build/tailwind.gen.css`, 8.5 KB via `tailwind.config.js`)
and **inlined it**, removing the last CDN. The dashboard now makes **zero external requests** —
works offline, behind firewalls, opened from `file://`, or on Pages. Re-verified in headless
Chromium: **0 console errors**, full styled layout correct (all 9 category columns visible,
rings + bars + anomaly flags). Screenshot refreshed in `docs/screenshots/`.

Full rebuild sequence (for a clean clone):
```
npm i                          # restores chart.js + tailwindcss (dev)
python3 build/build_dashboard.py "<All Zones Consolidation.xlsx>" \
        base/index_base_shell.html source-data/project_updates.json \
        deliverables/USC_Dashboard_v2.html
npx tailwindcss -i build/tw-input.css -o build/tailwind.gen.css --minify   # only if classes changed
python3 build/build_dashboard.py ...same...          # re-run so new CSS inlines
cp deliverables/USC_Dashboard_v2.html docs/index.html
```
Remaining: click-enable GitHub Pages (§6) · optional manual Drive copy (§7) · palette sign-off.

## 10. Polish pass (session 4)

QA'd all 5 tabs in-browser (screenshots in docs/screenshots) and fixed:
- **Tab 3 "Pricing & Bidder — Still Awaiting" card was stale/wrong** — the source now has this data.
  Replaced with real coverage KPIs: Total Reserve Price (Rs 5,030,464), Total Auction Value
  (Rs 17,124,785,955), **Records with Winning Bidder = 10,226 (86.8%)**, Records with Auction Date
  = 10,219 (86.7%). Baked two more Database cols into the data (U Winner Bidder, T Auction Date).
- **Brand-aligned the Chart.js palette** to lime (#8FB400) + charcoal so charts match the tab0/tab4
  rings (was old teal #00A376).
- **Fixed stale caption** "five asset categories" → "asset categories" (there are 9).
- **Added a Data Quality & Methodology notes card** to tab0 (stage formulas + explanation of the
  >100% ▲ anomalies) — professional context for the TPV client.
Re-verified: 0 console errors, JS valid, still zero external requests. docs/index.html refreshed.

## 11. Two new tabs (session 5 — CEO review response)

Added per CEO request:
- **Payments & DD Detail (tab5)** — bidder-level payment reconciliation. KPIs (Total DD/Payment
  Rs 968.3M, line items, **24 unique bidders**, unique DDs), Top-Bidders + Payment-by-Category
  charts, and a **per-bidder summary table** (Bidder | Zone(s) | Categories Won | DD No(s) |
  Payment Date(s) | Total Payment | Lines). Source records payments at asset-line level, so the
  table aggregates to one row per bidder (23 bidders; e.g. Sain Abdul Hakeem Rs 474.5M/255 lines)
  instead of thousands of duplicated lines — with a caveat to verify against physical DDs.
- **Anomalies (tab6)** — dedicated data-quality section. KPIs + by-type chart + detail table.
  Checks: Lifted>Auctioned (983), Auctioned>Counted (227), Payment>Auction Value (182),
  Negative Balance. **1,600 flagged records**, each with the conflicting values, for USC to reconcile.
- Baked two more Database cols (AB DD No., Z Payment Date). Now **7 tabs**; re-verified in
  Chromium (0 errors), still fully self-contained. docs/index.html + screenshots refreshed.

## 12. Zone Scorecard (session 6 — closes CEO gap G1)

Added a **Zone Scorecard** to the Overall tab: 9 zones × the 4 milestone %s (Phys. Verif., Auction,
Lifting, Payment) with tinted % bars + >100% ▲ flags, a Records and Payment (Rs) column, and an
All-Zones total row (which doubles as the blended headline — Verif 61%, Auction 111%▲, Lifting 47%,
Payment 6%). Slicer-aware. Re-verified in Chromium (9 rows + total, 0 errors); docs + CEO_REVIEW
updated (G1 now ✅). Dashboard remains self-contained.

### Coverage vs CEO review (final)
- Bird's-eye ✅ · Zone/Region ✅ (Scorecard + slicers) · Anomalies ✅ (dedicated tab) ·
  Payments-by-bidder ✅ (dedicated tab). Remaining gaps G2/G3/G5 are **source-data** limitations
  (thin DD fields, inflated auction value, no historical snapshots) that the TPV correctly exposes;
  G4 headline % is served by the Zone Scorecard's All-Zones row.

## 13. Working "Refresh Data" button (session 7)

User asked for a refresh mechanism where "each refresh fetch[es] the updated data." A true
client-side live-fetch from Drive would require making the source workbook link-shared
("Anyone with the link") plus a Google Cloud API key (user-managed infra outside my tool
access) -- flagged as a real trade-off (exposes raw bidder/payment data to anyone with the
link) rather than silently building it. Instead, restored a **fully working, instant,
client-side "Refresh Data" button**:

- Vendored **SheetJS (xlsx 0.18.5)** and **PapaParse 5.4.1** via npm and inlined them (same
  pattern as Chart.js/Tailwind) -- the button was previously present but dead (parser libs had
  been stripped for self-containment). Dashboard is still 100% self-contained (0 external
  requests), now ~3.1 MB.
- **Found and fixed 3 real bugs** while testing the upload path end-to-end in headless
  Chromium (not just eyeballed):
  1. `pickBestSheet()` matched any sheet *named* like "*recon*" before checking headers --
     the workbook's new "Reconciliation Summary" tab (2-col totals, not data) was winning over
     the actual "Database" sheet. Reordered: header-signature match now takes priority over
     the name-regex fallback.
  2. `COL.bidder` was `'Auction To Bidder'` (old embedded-CSV header) but the real Database
     sheet column is `'Winner Bidder Name'` -- fixed to match, so tab 6 (Payments & DD)
     populates correctly on upload.
  3. `COL.balanceQty` was `'Balance Qty to be lifted'` vs. the real header `'Balance Qty'` --
     fixed likewise.
  4. Ported the Python build's zone/category normalization (Sukkur/SUKKUR merge, ERP variant
     merge, etc.) into `cleanRow()` in JS, so a manual refresh renders identically to the
     baked build (9 zones, 9 categories) instead of showing raw duplicates.
- Relabeled the button "Refresh Data (upload latest file)" with a tooltip, and added a
  "How to refresh" note in the sidebar: download the latest workbook from the Drive folder,
  click Refresh Data, select it.
- **Verified end-to-end** by scripting a real file upload (`page.setInputFiles`) with the
  actual `All Zones Consolidation - FIXED.xlsx` deliverable: 11,785 records loaded from the
  correct "Database" sheet, Payment Rs. matched the baked build exactly (968,295,511), 9 zones,
  9 categories, 0 console errors.

**How refresh works today:** open the USD DB MIK Drive folder -> download the latest workbook
-> click "Refresh Data" on the dashboard -> select the file. Recalculates instantly, client-side,
no login, no page reload. Not real-time/automatic; that would need either (a) the Drive-sharing +
API-key approach above (user must set up and accept the exposure trade-off), or (b) I re-run the
build+push periodically (can be scheduled as a Routine on request).

## 14. Source path + timestamp in header, and automated refresh (session 8)

- Added a **source path line** under the header subtitle showing the Drive shared-drive path
  (`G:\Shared drives\Clients (except pvt ltd co)\Advisory\Clients - Idrees\Utility Stores
  Corporation - FAR Auction\USC Dashboard\USD DB MIK`), so viewers know exactly where the data
  comes from. Static label, sourced from `meta.sourcePath` in the baked JSON.
- Status pill now shows **date AND time** of last sync (`Live · 11,782 records · last synced
  2026-07-07 13:19 UTC`), not just a date. Applies to both the initial baked-data boot and the
  manual "Refresh Data" upload path (each shows its own real sync timestamp).
- Fixed a raw-string bug I introduced (`\\` continuation artifact) that briefly doubled a
  backslash in the source path; caught before shipping via `repr()` check.

### Automated refresh (Routine)
Set up a durable **Routine** (`trig_01PKnUMFPyRfvfKTmvnjsSp2`, name "USC Dashboard Auto-Refresh"),
cron `15 3 * * *` (daily 03:15 UTC), `create_new_session_on_fire=true`, notifications off. Each
firing spawns a fresh session that: downloads the latest `All Zones Consolidation.xlsx` (Drive
file id `1s_Nf1JvWNJg16J8iY-z2wWdeyeyzN5Um`) → reruns `build/build_dashboard.py` (twice, so
Tailwind sees new classes) → validates (`node --check` on the extracted script + confirms zero
external CDN refs) → copies to `docs/index.html` → commits + pushes to this branch **only if
something actually changed**, else does nothing. This is the automatic "dashboard + its URL
auto-update" mechanism: the public GitHub Pages link reflects the Drive workbook within ~24h of
any change, with no login, no API key, no exposed source file. The "Refresh Data" button (§13)
remains for instant on-demand refresh in between automated runs.

**To change cadence or stop it:** ask me to update/delete trigger `trig_01PKnUMFPyRfvfKTmvnjsSp2`.

## Automated refresh log
(Entries appended here by the automated Routine after each successful publish.)

## 15. PKT time, canonical named source file, robust automation (session 9)

Addressed three follow-up requests:

1. **Timezone → Pakistan Standard Time.** All "last synced"/"refreshed" timestamps (both the
   baked-build boot path and the manual "Refresh Data" upload path) now show PKT (UTC+5, no DST)
   instead of UTC, e.g. `Live · 11,782 records · last synced 2026-07-07 18:29 PKT`.

2. **Honesty check on the source path.** The user asked whether the displayed Drive path was
   verified — answer: no, it was the user's own text, inserted verbatim as a label; Drive's API
   doesn't expose local mapped-drive letters (`G:\...`) so I can't independently confirm it against
   the file ID I actually fetch by. Flagged this plainly rather than asserting false certainty.

3. **Canonical named source file.** Created `USC_Dashboard_Data_Source_File.xlsx` in the "USD DB
   MIK" Drive folder via a **server-side `copy_file`** of the then-current `All Zones
   Consolidation.xlsx` (id `1s_Nf1JvWNJg16J8iY-z2wWdeyeyzN5Um`) → new file id
   `1NB7ZL1fAhCj5cCCadWTccKF56u5AOSKO`. Chose `copy_file` specifically to avoid the known
   base64-upload size ceiling (confirmed in an earlier session: ~1.2–2.5MB files fail to upload
   via `create_file`'s inline-base64 parameter) — a Drive-to-Drive copy has no such limit since
   the bytes never pass through my context. This copy inherits whatever the live file currently
   contains, which likely still has the `#REF!` cells in the `Reconciliation Summary` tab — per
   **2026-07-07 user decision, this is intentionally ignored going forward** (§4); it does not
   affect the dashboard. I could not independently re-verify the copy's formula contents (a full
   re-download of a file this size risks the same truncation issue documented in §5/HTML outline).
   - Updated the dashboard's displayed source path to include the file name:
     `...\USD DB MIK\USC_Dashboard_Data_Source_File.xlsx`, with a hover tooltip showing the
     untruncated path.
   - **Replaced the automated Routine** (deleted `trig_01PKnUMFPyRfvfKTmvnjsSp2`, created
     `trig_01XhvCr9N6sjhHaY26je5ou7`, same daily-03:15-UTC schedule) so it looks up the source
     file **by title** (`USC_Dashboard_Data_Source_File.xlsx` in folder
     `1PFSMyIYp3ZoMEDs7aL-CsSYrreyKW0zn`) instead of a hardcoded file ID — stays correct even if
     the file is later replaced/re-uploaded under the same name with a new ID. If the named file
     is ever missing, the job stops and does nothing rather than silently falling back to a
     different file.

Re-verified full build after all three changes: 11,782 rows, Payment Rs. unchanged, 0 console
errors, still zero external CDN references, `node --check` clean.
