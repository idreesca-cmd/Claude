# Project Updates.xlsx — normalized

Source id `1f_soWkh2MS0AVQ7CMedK6rxUgW0RB661`. Two sheets: **Project Status** and **News**.
These are free-text org/status trackers with no fixed schema. Most rows are just the zone→region→lead org structure with empty status. Only a few rows carry actual status/dates (extracted below). Full normalized array in `project_updates.json`.

## Sheet 1 — Project Status (structure + notable status)

The sheet is a matrix: `Ser No, Zone, Regions, Zone Lead, Team Members`, then per-asset-category `Auction Status` and `Lifting Status` columns (Fixed Assets, IT Eq., Vehicles, R&P, Spices, Own Brd, M&B). Rows with real data:

| Zone | Region(s) | Lead | Status / Date | Note |
|------|-----------|------|---------------|------|
| FSD | Mandi Bahauddin, Khushab, Sargodah, Faisalabad N, Faisalabad S, Jhang | Amim Ikram (Ex-USC), Umar Sheikh (BT) | Advertised 17 Sep 2025 → Auctioned 23 Sep 2025 | Same dates in both Auction and Lifting status columns |
| LHR | Sheikhupura, Sialkot, Gujranwala, Lahore N, Lahore S, Okara, Sahiwal | Taimoor (Ex-USC), Gohar & team | Item-level auction data captured | Rows contain item detail (Atta Stand, Balty Stands, Direction Board, Double Rack, End Rack, Fire Basket…) with Odoo/Physical/Difference sub-columns |

All other zones (Islamabad, Abbottabad, Peshawar, MTN, SKR, KHI, QTA) list only the org structure (region + assigned lead), no status filled in yet. Full lead assignments are in the JSON.

## Sheet 2 — News (dated free-text log)

| Date | Note |
|------|------|
| 2026-07-06 | Lifting is planned in Jhang on 02-07-2026 |
| 2026-07-03 | Lahore Zone Fixed Assets auction is completed. Report shared by Gohar Ali. |

## Normalized field mapping
Each record → `{zone, category, region, lead, status, date, note, source}`. `category` is blank for Project-Status rows (that sheet is region-oriented, not category-oriented); `region` is blank for News rows. Dates left as written where a real date exists, blank where not inferable — nothing fabricated.
