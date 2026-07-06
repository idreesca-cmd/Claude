# USC_Dashboard_Published.html — structure outline

Saved to `scratchpad/USC_Dashboard_Published.html`.

**IMPORTANT — the Drive download is truncated.** `download_file_content` returned only 271,514 bytes; the real file is ~276 KB. The saved copy ends mid-function at `return RAW.filter(row => (!z || row[COL.` — the tail of the JS (getFiltered body, tab-switch wiring, and closing tags) is missing. A fuller local copy exists at `scratchpad/published.html` (276,128 B) but it is truncated at the *same* point. The complete, un-minified source (without the giant embedded CSV) is `scratchpad/index.html` (49,820 B) and DOES contain the tab-switch wiring and full render code. Use `index.html` as the authoritative source for JS logic and `USC_Dashboard_Published.html` for the published layout + embedded data.

## 1. `<head>` libraries / CDNs
| Line | Library | Version | URL |
|------|---------|---------|-----|
| 7 | Tailwind CSS | (play CDN, latest) | https://cdn.tailwindcss.com |
| 8 | Chart.js | 4.4.1 | https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js |
| 9 | PapaParse | 5.4.1 | https://cdn.jsdelivr.net/npm/papaparse@5.4.1/papaparse.min.js |
| 10 | SheetJS (xlsx) | 0.18.5 | https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js |

## 2. CSS `:root` custom properties (lines 13–22)
```
--bt-green:      #00A376
--bt-green-dark: #007F5C
--bt-green-tint: #E5F6F0
--bt-slate:      #2C3136
--bt-slate-soft: #3F454C
--bt-bg:         #F9FAFB
--bt-border:     #E5E7EB
--bt-muted:      #6B7280
--bt-red:        #DC2626
--bt-amber:      #D97706
```

## 3. Tab structure (lines 2220–2224 nav; panes at 2228/2278/2322/2388)
| Button label | `data-tab` | pane `id` | pane line |
|--------------|-----------|-----------|-----------|
| Overall Project Status | `tab1` | `#tab1` | 2228 |
| Quantity Reconciliation | `tab2` | `#tab2` | 2278 |
| Auction Reconciliation | `tab3` | `#tab3` | 2322 |
| Payment & Lifting **`<span class="pending-pill">Pending</span>`** | `tab4` | `#tab4` | 2388 |

Buttons: `class="tab-btn"` (+`active` on tab1). Panes: `class="tab-pane"` (+`active` on tab1); `.tab-pane{display:none}` / `.tab-pane.active{display:block}`.

**switchTab logic** (from `index.html` lines 655–664 — this is the block missing from the truncated published file):
```js
/* ---------- Tab navigation ---------- */
document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById(btn.dataset.tab).classList.add('active');
    renderAll();
  });
});
```

## 4. Embedded data
- Element: `<script id="embeddedCsv" type="text/plain">` — opens **line 89**, closes `</script>` at **line 2141**.
- Loader (lines 2142–2152): on `DOMContentLoaded`, reads `#embeddedCsv`, `Papa.parse(raw,{header:true,skipEmptyLines:true,dynamicTyping:false})`, maps `cleanRow`, filters, sets `RAW`, then `hydrateSlicers(); renderAll();`.
- **Header row (line 90)** — 25 columns:
  `Zone, Region, Item Category, Asset Category, Asset Name (Matched), A QTY (Odoo), B QTY (Physical Lists), C QTY (Physical Count), Auction Status, Auction Date, Auction Qty Usable, Auction Qty Scrap, Auction Qty Total, Auction To Bidder, Reserve Price Usable, Reserve Price Scrape, Auction Price Usable, Auction Price Scrape, Payment Date, Payment Rs., Payment Demand Draft (DD) No., Lifting Date, Lifted Qty, Lifted KGs, Balance Qty to be lifted`
- Data rows: lines 91–2140 = **~2,050 rows** (this embedded set is a SUBSET, not the full 11,785-row Database from the consolidation workbook).
- First 2 data rows (lines 91–92):
  ```
  Quetta Zone,Quetta Region,Fixed Assets,"Furniture, Fixtures and Office Equipment",TYPEWRITER,4,  - ,  - ,Auctioned,,,,0,...
  Quetta Zone,Quetta Region,Fixed Assets,"Furniture, Fixtures and Office Equipment",DUPLICATING MACHINE,1,  - ,  - ,Auctioned,...
  ```
  Note: this embedded header differs from the consolidation `Database` sheet header (it adds `Item Category`, uses different auction/price column names). Column mapping is via a `COL` object in the JS.

## 5. Manual upload flow (to locate/replace)
- File input: `<input id="csvInput" ...>` accept `.xlsx,.xls,.xlsm,.csv,.tsv` (line ~2180). Upload button near line 2176 currently labeled **"Awaiting data"**.
- `change` handler wired at **line 2530** (`document.getElementById('csvInput').addEventListener('change', e => {...})`): derives extension, routes to `handleExcel(file)` for xlsx/xls/xlsm/xlsb/ods or `handleDelimited(file)` for csv/tsv/txt.
- **`function handleDelimited(file)`** — line **2554**; uses `Papa.parse(file,{header:true,skipEmptyLines:true,dynamicTyping:false,...})` (parse call line 2559).
- **`function handleExcel(file)`** — line **2572**; `new FileReader()` (2577), `XLSX.read(new Uint8Array(...), {type:'array', cellDates:true})` (2580), sheet pick + `XLSX.utils.sheet_to_json(...)` (2590 / 2650 / 2657), `readAsArrayBuffer` (2604).
- Helpers: `function parseNumber` (2461), `function parseDate` (2470), `showToast` (2511), `setStatus` (2662).
- These are the functions the main agent should replace/extend to ingest the full consolidation dataset.

## 6. "Pending" / "awaiting data" placeholder markers (exact strings)
- Tab nav pill (line 2224): `Payment & Lifting <span class="pending-pill">Pending</span>`.
- Header/upload area (line 2176): **"Awaiting data"**.
- **Tab 3 (Auction Reconciliation)** card (lines 2371–2382): title **"Pricing & Bidder Details — Still Awaiting"**, body *"Auction Qty is now captured, but the financial-side fields below remain blank in the source. They will populate once the auction is monetised and bidder data is recorded."* Chips: Auction Date, Auction To Bidder, Reserve Price Usable, Reserve Price Scrape, Auction Price Usable, Auction Price Scrape.
- **Tab 4 (Payment & Lifting)** card (lines 2388–2402): `<div class="awaiting-card">`, icon ⏳, title **"Payment & Lifting Reconciliation — Awaiting Data"**, body *"Payment receipts and physical lifting movements will be captured after the auction concludes and successful bidders make payment and collect their assets. This tab will populate automatically once the source workbook is updated."* Chips: Payment Date, Payment Rs., Payment DD No., Lifting Date, Lifted Qty, Lifted KGs, Balance Qty to be lifted.
- CSS classes: `.awaiting-card` / `.awaiting-icon` / `.awaiting-title` / `.awaiting-sub` / `.awaiting-chip` (lines 53–58); `.pending-pill` (line 28).
- Source comment (near end of index.html/published): *"Auction / Payment / Lifting tabs show 'awaiting data' placeholders."*

**NOTE for un-marking:** the consolidation `Database` sheet DOES now contain real Payment (`Payment Rs.` grand total ≈ 968,295,511) and Lifting (`Lifted Qty` ≈ 930,063) data — so tab4's "Pending"/"Awaiting Data" placeholder can be replaced with live content.
