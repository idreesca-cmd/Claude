# All Zones Consolidation.xlsx — summary

Saved to `scratchpad/all_zones_consolidation.xlsx`. Machine-readable aggregates in `scratchpad/consolidation_data.json`.

## Sheets
`Database` (11,788 rows × 37 cols), `Reconciliation Summary` (32×2), `Multan Summary` (95×14), `Karachi Summary` (97×12), `Islamabad Steel W - Audit` (121×3), `Peshawar Steel W - Audit` (82×3).

## Database sheet
- **Data rows:** 11,785 (row 1 header; rows 2–11788; 3 effectively blank).
- **Header row (37 cols, in order):**
  `A Zone | B Region | C Asset Category | D Asset Name (Matched) | E A QTY (Odoo) | F B QTY (Physical Lists) | G C QTY (Physical Count) | H A - B | I B - C | J Auction Qty Usable | K Auction Qty Scrap | L Auction Qty Total | M Diff Auction vs Count | N Estimated Weight Per Piece | O Total Estimated Weight (Kg) | P Reserve Price Usable (Per KG/Piece) | Q Reserve Price Scrape (Per KG/Piece) | R Total Reserve Price | S Auction Status | T Auction Date | U Winner Bidder Name | V Auction Price Usable | W Per unit KGs | X Auction Price Scrape | Y Total Auction Value | Z Payment Date | AA Payment Rs. | AB Payment Demand Draft (DD) No. | AC Lifting Date | AD Lifted Qty | AE Lifted KGs | AF Balance Qty | AG–AJ (blank) | AK jl`

### Key column mapping (exact header names)
| Concept | Column | Header |
|---------|--------|--------|
| Odoo / system qty | E | `A QTY (Odoo)` |
| Physical list qty | F | `B QTY (Physical Lists)` |
| Physical / verified count | G | `C QTY (Physical Count)` |
| Auction / sold qty | L | `Auction Qty Total` (usable J + scrap K) |
| Auction status | S | `Auction Status` |
| Reserve price (total) | R | `Total Reserve Price` |
| Payment / amount received | AA | `Payment Rs.` |
| Auction value (billed) | Y | `Total Auction Value` |
| Lifting status → lifted | AD / AF | `Lifted Qty` / `Balance Qty` (no text "lifting status" col; AC `Lifting Date` + AD/AF are the lifting fields) |

### Distinct Asset Categories — 13 exact strings (CRITICAL, verbatim; note whitespace/case dupes)
1. `FURNITURE, FIXTURE AND OFFICE` (3991 rows)
2. `MOTOR VEHICLES AND BICYCLES` (506)
3. `SIGN BOARDS` (143)
4. `COMPUTER AND OFFICE MACHINE/EQUIPMENT` (1413)
5. `PLANT AND EQUIPMENT` (383)
6. `ERP - COMPUTER & OFFICE MACHINES` (1141)
7. `Own Brand` (1020)
8. ` Rice & Pulses ` (873)  ← note leading AND trailing spaces
9. `Branded Goods (BG) Non Food` (2297)
10. `ERP - COMPUTER AND IT EQUIPMENT` (2)
11. `Rice & Pulses` (5)  ← no spaces (duplicate of #8)
12. `ERP Computer and IT Equipment` (2)
13. `ERP ` (6)  ← trailing space

Rows 10–13 are tiny near-duplicate/variant tags. The 9 "real" categories are #1–9. For display, main agent should probably merge `Rice & Pulses`/` Rice & Pulses ` and the three `ERP*` variants.

### Distinct Zones — 10 (note SUKKUR vs Sukkur case dupe)
`SUKKUR`, `Sukkur`, `Karachi`, `Multan`, `Quetta`, `Lahore`, `Abbottabad`, `Islamabad`, `Peshawar`, `Faisalabad`

### Auction Status values (messy — needs normalization)
`Auctioned` 38 · `Pending` 173 · `NOT AUCTIONED` 5334 · `AUCTIONED` 1063 · `Auctioned ` (trailing space) 1294 · `Not Auctioned` 29 · `Consignment` 49.

## Per-category milestone aggregates (4 stages)
Definitions used (columns that actually exist):
- **Physical Verification:** total = Σ Odoo (E); achieved = Σ Physical Count (G); % = G/E.
- **Auction:** total = Σ Physical Count (G, what's on hand); achieved = Σ Auction Qty Total (L); % = L/G.
- **Lifting:** total = Σ Auction Qty Total (L); achieved = Σ Lifted Qty (AD); % = AD/L.
- **Payment:** total = Σ Total Auction Value (Y); achieved = Σ Payment Rs. (AA); % = AA/Y. (Reserve value R also summed for reference.)

Full numbers per category live in `consolidation_data.json` → `per_category`. Compact view (main 9 categories; qty rounded):

| Category | PhysVerif Odoo→Count (%) | Auction Count→Auct (%) | Lifting Auct→Lifted (%) | Payment Value→Received (%) |
|----------|--------------------------|------------------------|-------------------------|----------------------------|
| FURNITURE, FIXTURE AND OFFICE | 345,529 → 78,407 (22.7%) | 78,407 → 24,641 (31.4%) | 24,641 → 27,661 (112%*) | 13,060,520,391 → 331,979,849 (2.5%) |
| MOTOR VEHICLES AND BICYCLES | 486 → 306 (63.0%) | 306 → 0 (0%) | 0 → 0 (n/a) | 0 → 0 (n/a) |
| SIGN BOARDS | 4,373 → 1,892 (43.3%) | 1,892 → 701 (37.1%) | 701 → 524 (74.8%) | 5,823,610 → 15,441,340 (265%*) |
| COMPUTER AND OFFICE MACHINE/EQUIPMENT | 9,700 → 3,768 (38.8%) | 3,768 → 417 (11.1%) | 417 → 197 (47.2%) | 388,422 → 135,642 (34.9%) |
| PLANT AND EQUIPMENT | 5,535 → 1,760 (31.8%) | 1,760 → 37 (2.1%) | 37 → 322 (870%*) | 4,761,563 → 18,872,514 (396%*) |
| ERP - COMPUTER & OFFICE MACHINES | 29,343 → 25,912 (88.3%) | 25,912 → 2,875 (11.1%) | 2,875 → 0 (0%) | 0 → 0 (n/a) |
| Own Brand | 1,032,033 → 559,080 (54.2%) | 559,080 → 187,374 (33.5%) | 187,374 → 175,462 (93.6%) | 50,967,805 → 54,150,418 (106%*) |
| ` Rice & Pulses ` | 1,342,954 → 1,053,129 (78.4%) | 1,053,129 → 1,728,702 (164%*) | 1,728,702 → 725,897 (42.0%) | 4,002,324,163 → 547,715,748 (13.7%) |
| Branded Goods (BG) Non Food | 150,745 → 49,176 (32.6%) | 49,176 → 49,176 (100%) | 49,176 → 0 (0%) | 0 → 0 (n/a) |

`*` = >100% or otherwise anomalous — the raw source has inconsistencies (lifted qty exceeding auctioned qty, payments exceeding billed value, Rice auction qty exceeding count). Numbers are reported straight from source; flag for data-quality review, do not silently clamp.

### Stage coverage notes
- Every stage has a supporting column, so no stage is missing. But: several categories have **zero payment/auction-value** (Motor Vehicles, both ERP variants, Branded Goods) → their Payment stage is genuinely 0 / n/a, not missing.
- There is no dedicated free-text "Lifting Status" column; lifting is derived from `Lifted Qty` (AD) / `Balance Qty` (AF) / `Lifting Date` (AC).
- Grand totals across all categories (for the fixed Reconciliation Summary) are in `ref_errors_and_fixes.md`.

## Payment & Lifting — real data now present?
**YES.** Grand totals: `Payment Rs.` ≈ **968,295,511**, `Lifted Qty` ≈ **930,063**, `Lifted KGs` ≈ **1,428,183**, `Balance Qty` ≈ 845,605. The dashboard's tab4 "Pending"/"Awaiting Data" placeholder can be replaced with live figures.
