# Reconciliation Summary — #REF! diagnosis & proposed fixes

Workbook: `All Zones Consolidation.xlsx`. Tab: **Reconciliation Summary** (a 2-column grand-totals pack, A=label, B=value). No edits made to the file — diagnosis only.

## Affected cells
All **15** value cells **B4:B18** are broken. Every one has formula `=SUM(#REF!)` and cached value `#REF!`. No other sheet in the workbook contains #REF!.

| Cell | Label (col A) | Broken formula |
|------|---------------|----------------|
| B4  | A QTY (Odoo) | `=SUM(#REF!)` |
| B5  | B QTY (Physical Lists) | `=SUM(#REF!)` |
| B6  | C QTY (Physical Count) | `=SUM(#REF!)` |
| B7  | A - B | `=SUM(#REF!)` |
| B8  | B - C | `=SUM(#REF!)` |
| B9  | Auction Qty Usable | `=SUM(#REF!)` |
| B10 | Auction Qty Scrap | `=SUM(#REF!)` |
| B11 | Auction Qty Total | `=SUM(#REF!)` |
| B12 | Diff Auction vs Count | `=SUM(#REF!)` |
| B13 | Total Reserve Price | `=SUM(#REF!)` |
| B14 | Total Auction Value | `=SUM(#REF!)` |
| B15 | Payment Rs. | `=SUM(#REF!)` |
| B16 | Lifted Qty | `=SUM(#REF!)` |
| B17 | Lifted KGs | `=SUM(#REF!)` |
| B18 | Balance Qty | `=SUM(#REF!)` |

## Root cause
The notes in this same sheet (cell A32) say: *"Totals feed from helper sheet '_PivotSource'."* That helper sheet — and the referenced `Pivot - Zone/Region/Category/Status` sheets mentioned in A20–A24 — **no longer exist in the workbook**. Current sheets are only: `Database`, `Reconciliation Summary`, `Multan Summary`, `Karachi Summary`, `Islamabad Steel W - Audit`, `Peshawar Steel W - Audit`.

When those source sheets/ranges were deleted, Excel rewrote each `=SUM('_PivotSource'!<col>)` reference to `=SUM(#REF!)` because the target range was destroyed. It is a classic deleted-sheet/deleted-range breakage, not a typo or a circular reference.

## Proposed fix
Re-point each SUM at the corresponding column in the surviving **`Database`** sheet. Database layout: header in row 1, data rows **2 through 11788** (no trailing Subtotal/Total row — verified; the last row 11788 is a real Rice & Pulses item). `SUM` ignores the text/blank "issue" cells automatically, matching the note in A30.

Column map (Database), and the corrected formula for each cell:

| Cell | Database column | Corrected formula |
|------|-----------------|-------------------|
| B4  | E  A QTY (Odoo) | `=SUM(Database!E2:E11788)` |
| B5  | F  B QTY (Physical Lists) | `=SUM(Database!F2:F11788)` |
| B6  | G  C QTY (Physical Count) | `=SUM(Database!G2:G11788)` |
| B7  | H  A - B | `=SUM(Database!H2:H11788)` |
| B8  | I  B - C | `=SUM(Database!I2:I11788)` |
| B9  | J  Auction Qty Usable | `=SUM(Database!J2:J11788)` |
| B10 | K  Auction Qty Scrap | `=SUM(Database!K2:K11788)` |
| B11 | L  Auction Qty Total | `=SUM(Database!L2:L11788)` |
| B12 | M  Diff Auction vs Count | `=SUM(Database!M2:M11788)` |
| B13 | R  Total Reserve Price | `=SUM(Database!R2:R11788)` |
| B14 | Y  Total Auction Value | `=SUM(Database!Y2:Y11788)` |
| B15 | AA Payment Rs. | `=SUM(Database!AA2:AA11788)` |
| B16 | AD Lifted Qty | `=SUM(Database!AD2:AD11788)` |
| B17 | AE Lifted KGs | `=SUM(Database!AE2:AE11788)` |
| B18 | AF Balance Qty | `=SUM(Database!AF2:AF11788)` |

(Use whole-column form `=SUM(Database!E:E)` if the row count may grow — equivalent here since the header is text and is ignored by SUM.)

## Expected grand totals after fix (computed from Database, text/blank treated as 0)
| Label | Value |
|-------|-------|
| A QTY (Odoo) | 2,951,485 |
| B QTY (Physical Lists) | 1,285,595 |
| C QTY (Physical Count) | 1,804,070 |
| A - B | 1,666,174 |
| B - C | -518,475 |
| Auction Qty Usable | 2,138,827 |
| Auction Qty Scrap | 199,799 |
| Auction Qty Total | 1,993,923 |
| Diff Auction vs Count | -758,306 |
| Total Reserve Price | 5,030,464 |
| Total Auction Value | 17,124,785,955 |
| Payment Rs. | 968,295,511 |
| Lifted Qty | 930,063 |
| Lifted KGs | 1,428,183 |
| Balance Qty | 845,605 |

(Total Auction Value looks very large — driven by the FURNITURE and Rice & Pulses categories in Database; it reflects the raw source values, which may themselves warrant a data-quality check, but the formula fix is independent of that.)
