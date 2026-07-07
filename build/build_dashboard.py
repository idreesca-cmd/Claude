#!/usr/bin/env python3
"""
USC Asset Disposal Dashboard — regenerator / build step.

Reads the Drive source workbook (All Zones Consolidation.xlsx, sheet `Database`)
+ normalized Project Updates, bakes the data into the clean app shell, adds the
new "Category Milestone Overview" first tab, un-marks Payment & Lifting, and
writes a self-contained USC_Dashboard_v2.html.

This is the "refresh" mechanism (Decision D1): re-run this after the Drive files
change, then publish the output (GitHub Pages + copy to Drive).

Usage:
    python3 build/build_dashboard.py <consolidation.xlsx> <base_shell.html> \
            <project_updates.json> <out.html>
"""
import sys, json, re, datetime, os
import openpyxl

SRC_XLSX, BASE_HTML, PU_JSON, OUT_HTML = sys.argv[1:5]

# ---------------------------------------------------------------- normalization
def norm_zone(z):
    z = (z or "").strip()
    return "Sukkur" if z.upper() == "SUKKUR" else z

CAT_MAP = {
    "FURNITURE, FIXTURE AND OFFICE": "Furniture, Fixture & Office",
    "MOTOR VEHICLES AND BICYCLES": "Motor Vehicles & Bicycles",
    "SIGN BOARDS": "Sign Boards",
    "COMPUTER AND OFFICE MACHINE/EQUIPMENT": "Computer & Office Equipment",
    "PLANT AND EQUIPMENT": "Plant & Equipment",
    "ERP - COMPUTER & OFFICE MACHINES": "ERP – Computer & Office Machines",
    "ERP - COMPUTER AND IT EQUIPMENT": "ERP – Computer & Office Machines",
    "ERP COMPUTER AND IT EQUIPMENT": "ERP – Computer & Office Machines",
    "ERP": "ERP – Computer & Office Machines",
    "OWN BRAND": "Own Brand",
    "RICE & PULSES": "Rice & Pulses",
    "BRANDED GOODS (BG) NON FOOD": "Branded Goods (Non-Food)",
}
CAT_ORDER = [
    "Furniture, Fixture & Office", "Motor Vehicles & Bicycles", "Sign Boards",
    "Computer & Office Equipment", "Plant & Equipment",
    "ERP – Computer & Office Machines", "Own Brand", "Rice & Pulses",
    "Branded Goods (Non-Food)",
]
FIXED_ASSET_CATS = {
    "Furniture, Fixture & Office", "Motor Vehicles & Bicycles", "Sign Boards",
    "Computer & Office Equipment", "Plant & Equipment",
    "ERP – Computer & Office Machines",
}
def norm_cat(c):
    key = re.sub(r"\s+", " ", (c or "").strip()).upper()
    return CAT_MAP.get(key, (c or "").strip())

def item_bucket(display_cat):
    return "Fixed Assets" if display_cat in FIXED_ASSET_CATS else "Merchandise & Inventory"

def norm_status(s):
    s = re.sub(r"\s+", " ", (s or "").strip())
    u = s.upper()
    if u == "AUCTIONED": return "Auctioned"
    if u in ("NOT AUCTIONED",): return "Not Auctioned"
    if u == "PENDING": return "Pending"
    if u == "CONSIGNMENT": return "Consignment"
    return s or ""

def num(v):
    if isinstance(v, (int, float)): return v
    if v is None: return 0
    try:
        return float(str(v).replace(",", "").replace("Rs", "").strip() or 0)
    except ValueError:
        return 0

# ---------------------------------------------------------------- read Database
wb = openpyxl.load_workbook(SRC_XLSX, data_only=True, read_only=True)
db = wb["Database"]
ci = openpyxl.utils.column_index_from_string
def C(row, letter): return row[ci(letter) - 1]

# fixed emit order -> mirrored in the browser bootstrap
ORDER = ["zone","region","itemCat","assetCat","assetName","aQty","bQty","cQty",
         "status","auctQtyUsable","auctQtyScrap","auctQtyTotal","reservePriceTotal",
         "auctValueTotal","payRs","liftedQty","liftedKg","balanceQty","bidder","auctionDate",
         "ddNo","payDate"]

def _dstr(v):
    if v is None or v == "": return ""
    if hasattr(v, "date"):
        try: return v.date().isoformat()
        except Exception: return str(v).strip()
    return str(v).strip()

rows_out = []
it = db.iter_rows(min_row=2, values_only=True)
for row in it:
    if row is None: continue
    zone = norm_zone(C(row, "A"))
    cat_raw = C(row, "C")
    name = C(row, "D")
    if not (zone or name or cat_raw):  # skip fully-blank
        continue
    cat = norm_cat(cat_raw)
    if not cat:  # drop the 3 blank-category rows
        continue
    rec = [
        zone, (C(row,"B") or "").strip() if isinstance(C(row,"B"),str) else (C(row,"B") or ""),
        item_bucket(cat), cat, (name or ""),
        num(C(row,"E")), num(C(row,"F")), num(C(row,"G")),
        norm_status(C(row,"S")),
        num(C(row,"J")), num(C(row,"K")), num(C(row,"L")),
        num(C(row,"R")), num(C(row,"Y")), num(C(row,"AA")),
        num(C(row,"AD")), num(C(row,"AE")), num(C(row,"AF")),
        _dstr(C(row,"U")), _dstr(C(row,"T")),
        _dstr(C(row,"AB")), _dstr(C(row,"Z")),
    ]
    rows_out.append(rec)

# grand totals (indices into ORDER)
idx = {k: i for i, k in enumerate(ORDER)}
def gsum(k): return sum(r[idx[k]] for r in rows_out)
meta = {
    "refreshDate": datetime.date.today().isoformat(),
    "rowCount": len(rows_out),
    "zones": sorted({r[0] for r in rows_out}),
    "categoryOrder": CAT_ORDER,
    "totals": {k: gsum(k) for k in
               ("aQty","bQty","cQty","auctQtyTotal","reservePriceTotal",
                "auctValueTotal","payRs","liftedQty","liftedKg","balanceQty")},
}
with open(PU_JSON, encoding="utf-8") as f:
    project_updates = json.load(f)

payload = {"cols": ORDER, "rows": rows_out, "meta": meta, "updates": project_updates}
print(f"[build] {len(rows_out):,} rows | Payment Rs {meta['totals']['payRs']:,.0f} "
      f"| Lifted {meta['totals']['liftedQty']:,.0f}")

# ---------------------------------------------------------------- load shell
with open(BASE_HTML, encoding="utf-8") as f:
    html = f.read()

def sub(pattern, repl, flags=0, count=1, label=""):
    global html
    new, n = re.subn(pattern, repl, html, count=count, flags=flags)
    assert n == count, f"anchor NOT found ({n}/{count}): {label or pattern[:60]}"
    html = new

def lit(old, new, count=1, label=""):
    """Plain (non-regex) literal replace with an occurrence assertion."""
    global html
    found = html.count(old)
    assert found == count, f"lit anchor found {found}x (want {count}): {label or old[:50]}"
    html = html.replace(old, new, count)

# 0) Make the page self-contained: inline Chart.js (vendored from npm) so the
#    dashboard works offline / behind firewalls, and drop the now-dead upload CDNs
#    (data is baked in; the optional upload handler already guards missing libs).
with open("node_modules/chart.js/dist/chart.umd.js", encoding="utf-8") as f:
    chartjs = f.read()
chart_tag = ("<script>/* Chart.js 4.4.1 — inlined for a self-contained, offline-safe file */\n"
             + chartjs.replace("</script>", "<\\/script>") + "\n</script>\n  ")
sub(r'<script src="https://cdn\.jsdelivr\.net/npm/chart\.js@4\.4\.1/dist/chart\.umd\.min\.js"></script>\s*',
    lambda m: chart_tag, label="inline chartjs")
sub(r'\s*<script src="https://cdn\.jsdelivr\.net/npm/papaparse@[^"]*"></script>', "",
    label="drop papaparse cdn")
sub(r'\s*<script src="https://cdn\.jsdelivr\.net/npm/xlsx@[^"]*"></script>', "",
    label="drop xlsx cdn")
# Inline compiled Tailwind (only used classes) if present, and drop the Tailwind CDN,
# so the page is fully self-contained (no external requests at all).
if os.path.exists("build/tailwind.gen.css"):
    with open("build/tailwind.gen.css", encoding="utf-8") as f:
        twcss = f.read()
    tw_tag = "<style>/* Tailwind — compiled, only the classes this page uses */\n" + twcss + "\n</style>"
    sub(r'<script src="https://cdn\.tailwindcss\.com"></script>', lambda m: tw_tag,
        label="inline tailwind")
else:
    print("[build] WARN: build/tailwind.gen.css missing — keeping Tailwind CDN")

# 1) palette + fonts (BT Brand Colors) — add brand vars, keep --bt-* scheme
sub(r"--bt-green:#00A376;",
    "--bt-green:#00A376; --bt-lime:#D1EC51; --bt-lime-deep:#8Fb400;"
    " --bt-mint:#DFF5F2; --bt-charcoal:#2C303B;",
    label="palette vars")
sub(r"font-family:'Inter','Segoe UI',system-ui,sans-serif;color:var\(--bt-slate\);",
    "font-family:'Calibri','Segoe UI','Inter',system-ui,sans-serif;color:var(--bt-slate);",
    label="body font")

# 2) new-tab CSS (inject before </style>)
TAB0_CSS = """
 /* Category Milestone Overview */
 .ms-table{border-collapse:separate;border-spacing:0;min-width:920px;width:100%;}
 .ms-table th,.ms-table td{padding:12px 10px;border-bottom:1px solid var(--bt-border);vertical-align:top;}
 .ms-table thead th{position:sticky;top:0;background:var(--bt-charcoal);color:#fff;font-size:12px;font-weight:600;text-align:center;border-bottom:none;}
 .ms-table thead th:first-child{background:var(--bt-charcoal);text-align:left;border-top-left-radius:10px;}
 .ms-table thead th:last-child{border-top-right-radius:10px;}
 .ms-stage{font-weight:700;color:var(--bt-slate);font-size:13px;white-space:nowrap;background:var(--bt-mint);}
 .ms-stage small{display:block;font-weight:500;color:var(--bt-muted);font-size:11px;margin-top:2px;}
 .ms-cell{text-align:center;min-width:120px;}
 .ring{width:64px;height:64px;border-radius:50%;margin:0 auto 6px;display:flex;align-items:center;justify-content:center;position:relative;}
 .ring::after{content:'';position:absolute;inset:8px;background:#fff;border-radius:50%;}
 .ring span{position:relative;z-index:1;font-size:13px;font-weight:700;color:var(--bt-slate);}
 .ms-qty{font-size:11.5px;color:var(--bt-muted);margin-bottom:6px;}
 .ms-qty b{color:var(--bt-slate);}
 .pa-bars{display:flex;gap:5px;justify-content:center;align-items:flex-end;height:34px;}
 .pa-bar{width:16px;border-radius:3px 3px 0 0;}
 .pa-bar.plan{background:var(--bt-charcoal);}
 .pa-bar.act{background:var(--bt-lime-deep);}
 .pa-legend{font-size:9.5px;color:#9CA3AF;margin-top:3px;letter-spacing:.03em;}
 .dq-flag{color:var(--bt-amber);font-weight:700;cursor:help;}
"""
sub(r"</style>", TAB0_CSS + "</style>", label="tab0 css")

# 3) nav — add tab0, drop pending pill, tab0 active
sub(r'<button class="tab-btn active" data-tab="tab1">Overall Project Status</button>',
    '<button class="tab-btn active" data-tab="tab0">Category Milestone Overview</button>\n'
    '        <button class="tab-btn" data-tab="tab1">Overall Project Status</button>',
    label="nav tab0")
sub(r'<button class="tab-btn" data-tab="tab4">Payment &amp; Lifting<span class="pending-pill">Pending</span></button>',
    '<button class="tab-btn" data-tab="tab4">Payment &amp; Lifting</button>',
    label="nav tab4 pill")

# 4) tab0 section (insert before tab1) + make tab1 not active
TAB0_SECTION = """
        <section id="tab0" class="tab-pane active">
          <div class="chart-card">
            <div class="chart-title">Category Milestone Overview</div>
            <div class="chart-sub">Progress by asset category across the four disposal stages. Rings show % achieved; the paired bars compare <b>planned</b> (stage target) vs <b>actual</b> (achieved). Figures respect the global slicers. <span class="dq-flag">&#9650;</span> marks values &gt;100% (source data-quality anomalies, shown unclamped).</div>
            <div id="milestoneMatrix" style="overflow-x:auto;"></div>
          </div>
          <div class="chart-card mt-4">
            <div class="chart-title">Data Quality &amp; Methodology Notes</div>
            <div class="chart-sub">How each stage is measured, and known source anomalies (shown unclamped for the client to reconcile).</div>
            <ul style="font-size:12.5px;color:#374151;line-height:1.7;margin:4px 0 0 18px;list-style:disc;">
              <li><b>Physical Verification</b> = Physical Count (C) &divide; Odoo Qty (A) &middot; <b>Auction</b> = Auctioned Qty &divide; Physical Count &middot; <b>Lifting</b> = Lifted Qty &divide; Auctioned Qty &middot; <b>Payment</b> = Payment Received &divide; Total Auction Value.</li>
              <li>Cells marked <span class="dq-flag">&#9650;</span> exceed 100% &mdash; genuine source inconsistencies (e.g. lifted qty &gt; auctioned qty, payment &gt; booked value, auctioned qty &gt; counted qty). Figures are shown exactly as recorded and are <b>not clamped</b>; they flag records USC should reconcile at source.</li>
              <li>Categories at 0% Payment/Lifting (Motor Vehicles, ERP, Branded Goods) have no monetised auction in the source yet &mdash; this is genuine status, not missing data.</li>
            </ul>
          </div>
        </section>

        <section id="tab1" class="tab-pane">"""
sub(r'<section id="tab1" class="tab-pane active">', TAB0_SECTION, label="tab0 section")

# 5) tab4 — replace awaiting-card with real KPIs + table
TAB4_SECTION = """<section id="tab4" class="tab-pane">
          <div class="grid grid-cols-2 md:grid-cols-5 gap-4 mb-5">
            <div class="kpi-card"><div class="kpi-label">Payment Received (Rs.)</div><div id="kpi4Pay" class="kpi-value">—</div><div class="kpi-sub">Sum of Payment Rs.</div></div>
            <div class="kpi-card"><div class="kpi-label">Total Auction Value (Rs.)</div><div id="kpi4Val" class="kpi-value">—</div><div id="kpi4ValSub" class="kpi-sub">Billed value</div></div>
            <div class="kpi-card"><div class="kpi-label">Lifted Qty</div><div id="kpi4Lift" class="kpi-value">—</div><div id="kpi4LiftSub" class="kpi-sub">Units collected</div></div>
            <div class="kpi-card"><div class="kpi-label">Lifted KGs</div><div id="kpi4Kg" class="kpi-value">—</div><div class="kpi-sub">Weight collected</div></div>
            <div class="kpi-card"><div class="kpi-label">Balance Qty to Lift</div><div id="kpi4Bal" class="kpi-value">—</div><div class="kpi-sub">Outstanding</div></div>
          </div>
          <div class="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-5">
            <div class="chart-card"><div class="chart-title">Payment Received by Zone</div><div class="chart-sub">Sum of Payment Rs. per zone (post-filter)</div><div style="position:relative;height:330px;"><canvas id="chartPayZone"></canvas></div></div>
            <div class="chart-card"><div class="chart-title">Lifted vs Balance by Zone</div><div class="chart-sub">Units lifted vs outstanding balance per zone</div><div style="position:relative;height:330px;"><canvas id="chartLiftZone"></canvas></div></div>
          </div>
          <div class="chart-card"><div class="chart-title">Payment &amp; Lifting Detail — by Zone / Category</div><div class="chart-sub">Auction value, payment received, lifted &amp; balance quantities</div><div id="payLiftTableWrap" class="data-table-wrap"></div></div>
        </section>"""
sub(r'<section id="tab4" class="tab-pane">.*?</section>', TAB4_SECTION,
    flags=re.DOTALL, label="tab4 section")

# 5b) tab3 — replace the stale "Pricing & Bidder — Still Awaiting" card with real coverage KPIs
TAB3_PRICING = ('<div class="grid grid-cols-2 md:grid-cols-4 gap-4 mt-4">\n'
  '            <div class="kpi-card"><div class="kpi-label">Total Reserve Price (Rs.)</div><div id="kpi3Reserve" class="kpi-value">—</div><div class="kpi-sub">Floor value set for lots</div></div>\n'
  '            <div class="kpi-card"><div class="kpi-label">Total Auction Value (Rs.)</div><div id="kpi3Value" class="kpi-value">—</div><div class="kpi-sub">Awarded / booked value</div></div>\n'
  '            <div class="kpi-card"><div class="kpi-label">Records with Winning Bidder</div><div id="kpi3Bidder" class="kpi-value">—</div><div id="kpi3BidderSub" class="kpi-sub">Bidder name recorded</div></div>\n'
  '            <div class="kpi-card"><div class="kpi-label">Records with Auction Date</div><div id="kpi3Date" class="kpi-value">—</div><div id="kpi3DateSub" class="kpi-sub">Auction date recorded</div></div>\n'
  '          </div>\n'
  '          <div class="chart-card mt-4" style="padding:14px 18px;"><div class="chart-sub" style="margin:0;">Pricing &amp; bidder fields are now partially populated in the source (see the corrected workbook). Zero or blank cells indicate lots not yet monetised or awarded; all figures are shown exactly as recorded.</div></div>\n        </section>')
sub(r'<div class="awaiting-card" style="padding:24px;">.*?</section>',
    lambda m: TAB3_PRICING, flags=re.DOTALL, label="tab3 pricing card")

# 5c) new nav buttons for Payments (tab5) + Anomalies (tab6)
lit('<button class="tab-btn" data-tab="tab4">Payment &amp; Lifting</button>',
    '<button class="tab-btn" data-tab="tab4">Payment &amp; Lifting</button>\n'
    '        <button class="tab-btn" data-tab="tab5">Payments &amp; DD Detail</button>\n'
    '        <button class="tab-btn" data-tab="tab6">Anomalies</button>',
    label="nav tab5/6")

# 5d) new sections for Payments + Anomalies (inserted before </main>)
TAB56 = """        <section id="tab5" class="tab-pane">
          <div class="grid grid-cols-2 md:grid-cols-4 gap-4 mb-5">
            <div class="kpi-card"><div class="kpi-label">Total DD / Payment (Rs.)</div><div id="kpi5Amt" class="kpi-value">—</div><div class="kpi-sub">Sum of payments received</div></div>
            <div class="kpi-card"><div class="kpi-label">Payment Line Items</div><div id="kpi5Recs" class="kpi-value">—</div><div class="kpi-sub">Asset lines carrying a payment</div></div>
            <div class="kpi-card"><div class="kpi-label">Unique Bidders</div><div id="kpi5Bidders" class="kpi-value">—</div><div class="kpi-sub">Distinct winning bidders</div></div>
            <div class="kpi-card"><div class="kpi-label">Unique DD Numbers</div><div id="kpi5DDs" class="kpi-value">—</div><div class="kpi-sub">Distinct demand drafts</div></div>
          </div>
          <div class="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-5">
            <div class="chart-card"><div class="chart-title">Top Bidders by Payment</div><div class="chart-sub">Highest total payment received (Rs.)</div><div style="position:relative;height:340px;"><canvas id="chartTopBidders"></canvas></div></div>
            <div class="chart-card"><div class="chart-title">Payment by Asset Category</div><div class="chart-sub">Total payment received per category (Rs.)</div><div style="position:relative;height:340px;"><canvas id="chartPayByCat"></canvas></div></div>
          </div>
          <div class="chart-card"><div class="chart-title">Payments Received — by Winning Bidder</div><div class="chart-sub">Bidder, asset categories won, DD number(s), payment date(s) and total amount. Source records payments at asset-line level, so totals sum all of a bidder's lines — verify against physical DDs where an amount repeats.</div><div id="paymentsTableWrap" class="data-table-wrap"></div></div>
        </section>
        <section id="tab6" class="tab-pane">
          <div class="grid grid-cols-2 md:grid-cols-4 gap-4 mb-5">
            <div class="kpi-card"><div class="kpi-label">Anomalous Records</div><div id="kpi6Total" class="kpi-value">—</div><div class="kpi-sub">Rows failing a consistency check</div></div>
            <div class="kpi-card"><div class="kpi-label">Lifted &gt; Auctioned</div><div id="kpi6Lift" class="kpi-value">—</div><div class="kpi-sub">Over-lifting</div></div>
            <div class="kpi-card"><div class="kpi-label">Payment &gt; Auction Value</div><div id="kpi6Pay" class="kpi-value">—</div><div class="kpi-sub">Over-payment vs booked value</div></div>
            <div class="kpi-card"><div class="kpi-label">Auctioned &gt; Counted</div><div id="kpi6Auct" class="kpi-value">—</div><div class="kpi-sub">Sold more than counted</div></div>
          </div>
          <div class="chart-card"><div class="chart-title">Anomalies by Type</div><div class="chart-sub">Count of records per anomaly category (respects the global slicers)</div><div style="position:relative;height:300px;"><canvas id="chartAnomTypes"></canvas></div></div>
          <div class="chart-card mt-4"><div class="chart-title">Anomaly Detail</div><div class="chart-sub">Every flagged record with the conflicting values (top 500). These are source data-quality issues for USC to reconcile — figures are shown exactly as recorded.</div><div id="anomalyTableWrap" class="data-table-wrap"></div></div>
        </section>
"""
lit('</main>', TAB56 + '        </main>', label="tab5/6 sections")

# 6) COL: add auctValueTotal + reservePriceTotal
sub(r"balanceQty:'Balance Qty to be lifted'",
    "balanceQty:'Balance Qty to be lifted',\n  auctValueTotal:'Total Auction Value', reservePriceTotal:'Total Reserve Price'",
    label="COL keys")

# 7) header status/label wording
sub(r"Awaiting data", "Loading…", label="status text")
sub(r"Upload Excel / CSV", "Replace data (optional)", label="upload label")

# 8) render calls: add tab0 + tab4
sub(r"renderTab3\(data\);",
    "renderTab3(data);\n  renderTab0(data);\n  renderTab4(data);\n  renderAuctionPricing(data);\n  renderTab5(data);\n  renderTab6(data);",
    label="renderAll calls")

# 8b) Brand-align the Chart.js palette (lime accent, matching tab0/tab4)
sub(r"green:'#00A376', greenDark:'#007F5C', greenTint:'#7DCBB3',",
    "green:'#8FB400', greenDark:'#5E7A00', greenTint:'#D1EC51',", label="BT green->lime")
sub(r"palette:\['#00A376','#2C3136','#D97706','#2563EB','#9333EA','#0891B2','#DC2626','#65A30D'\]",
    "palette:['#8FB400','#2C303B','#D97706','#0891B2','#9333EA','#2563EB','#DC2626','#00A376']",
    label="BT palette lime-first")

# 8c) Fix stale caption ("five" asset categories -> nine are in scope)
sub(r"Distribution across the five asset categories in scope",
    "Distribution across the asset categories in scope", label="caption fix")

# 9) cap big detail tables to 500 rows (perf with 11.7k rows)
sub(r"\$\{rows\.map\(r => `\s*\n\s*<tr>\s*\n\s*<td>\$\{escapeHtml\(r\.zone\)\}</td>\s*\n\s*<td>\$\{escapeHtml\(r\.region\)\}</td>\s*\n\s*<td>\$\{escapeHtml\(r\.assetCat\)\}",
    lambda m: m.group(0).replace("${rows.map(r => `", "${rows.slice(0,500).map(r => `", 1),
    flags=re.DOTALL, label="cap variance table", count=1)
sub(r"\$\{rows\.map\(r => `\s*\n\s*<tr>\s*\n\s*<td>\$\{escapeHtml\(r\.zone\)\}</td>\s*\n\s*<td>\$\{escapeHtml\(r\.region\)\}</td>\s*\n\s*<td>\$\{escapeHtml\(r\.cat\)\}",
    lambda m: m.group(0).replace("${rows.map(r => `", "${rows.slice(0,500).map(r => `", 1),
    flags=re.DOTALL, label="cap auction table", count=1)

# 10) inject renderTab0 / renderTab4 before the empty-state block
NEW_JS = r"""
/* ==================== TAB 0 — CATEGORY MILESTONE OVERVIEW ==================== */
const MS_STAGES = [
  { name:'Physical Verification', sub:'Odoo → Physical Count',
    total:g=>sum(g,r=>r[COL.aQty]), ach:g=>sum(g,r=>r[COL.cQty]) },
  { name:'Auction', sub:'Physical Count → Auctioned',
    total:g=>sum(g,r=>r[COL.cQty]), ach:g=>sum(g,r=>r[COL.auctQtyTotal]) },
  { name:'Lifting', sub:'Auctioned → Lifted',
    total:g=>sum(g,r=>r[COL.auctQtyTotal]), ach:g=>sum(g,r=>r[COL.liftedQty]) },
  { name:'Payment', sub:'Auction Value → Received (Rs.)',
    total:g=>sum(g,r=>r[COL.auctValueTotal]), ach:g=>sum(g,r=>r[COL.payRs]) },
];
function ringStyle(pct){
  const deg = Math.max(0, Math.min(100, pct))*3.6;
  const arc = pct>100 ? 'var(--bt-amber)' : 'var(--bt-lime-deep)';
  return `background:conic-gradient(${arc} ${deg}deg, #EDEFF2 ${deg}deg);`;
}
function paBars(total, ach){
  const mx = Math.max(total, ach, 1);
  const hp = Math.round((total/mx)*34), ha = Math.round((ach/mx)*34);
  return `<div class="pa-bars"><div class="pa-bar plan" style="height:${hp}px" title="Planned: ${fmtNum(total)}"></div>`+
         `<div class="pa-bar act" style="height:${ha}px" title="Actual: ${fmtNum(ach)}"></div></div>`+
         `<div class="pa-legend">Plan │ Act</div>`;
}
function renderTab0(data){
  const order = (window.DASH_META && DASH_META.categoryOrder) || [];
  const present = uniqueSorted(data.map(r=>r[COL.assetCat]));
  const cats = order.filter(c=>present.includes(c)).concat(present.filter(c=>!order.includes(c)));
  const groups = {}; cats.forEach(c=> groups[c] = data.filter(r=>r[COL.assetCat]===c));
  let h = '<table class="ms-table"><thead><tr><th>Milestone Stage</th>'+
          cats.map(c=>`<th>${escapeHtml(c)}</th>`).join('')+'</tr></thead><tbody>';
  MS_STAGES.forEach(st=>{
    h += `<tr><td class="ms-stage">${st.name}<small>${st.sub}</small></td>`;
    cats.forEach(c=>{
      const g = groups[c];
      const total = st.total(g), ach = st.ach(g);
      const pct = total>0 ? (ach/total*100) : null;
      const flag = (pct!==null && pct>100) ? ' <span class="dq-flag" title="Exceeds 100% — source data anomaly">▲</span>' : '';
      const label = pct===null ? '—' : pct.toFixed(0)+'%';
      h += `<td class="ms-cell">`+
           `<div class="ring" style="${ringStyle(pct||0)}"><span>${label}</span></div>`+
           `<div class="ms-qty"><b>${fmtNum(ach)}</b> / ${fmtNum(total)}${flag}</div>`+
           paBars(total, ach)+`</td>`;
    });
    h += '</tr>';
  });
  h += '</tbody></table>';
  document.getElementById('milestoneMatrix').innerHTML = h;
}

/* ==================== TAB 4 — PAYMENT & LIFTING ==================== */
function renderTab4(data){
  const pay = sum(data,r=>r[COL.payRs]);
  const val = sum(data,r=>r[COL.auctValueTotal]);
  const lift = sum(data,r=>r[COL.liftedQty]);
  const kg = sum(data,r=>r[COL.liftedKg]);
  const bal = sum(data,r=>r[COL.balanceQty]);
  const set=(id,v)=>{const e=document.getElementById(id); if(e) e.textContent=v;};
  set('kpi4Pay','Rs '+fmtNum(pay)); set('kpi4Val','Rs '+fmtNum(val));
  set('kpi4Lift',fmtNum(lift)); set('kpi4Kg',fmtNum(kg)); set('kpi4Bal',fmtNum(bal));
  const vsub=document.getElementById('kpi4ValSub');
  if(vsub) vsub.textContent = val>0 ? `Payment = ${(pay/val*100).toFixed(1)}% of value` : 'Billed value';

  destroyChart('chartPayZone');
  const gz = groupBy(data, COL.zone); const zl = Object.keys(gz).sort();
  if(data.length && zl.some(z=>sum(gz[z],r=>r[COL.payRs])>0)){
    CHARTS.chartPayZone = new Chart(document.getElementById('chartPayZone'),{ type:'bar',
      data:{ labels:zl, datasets:[{ label:'Payment Rs.', data:zl.map(z=>sum(gz[z],r=>r[COL.payRs])), backgroundColor:BT.slate, borderRadius:6, maxBarThickness:46 }]},
      options:{ responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}, tooltip:{callbacks:{label:ctx=>'Rs '+fmtNum(ctx.parsed.y)}}}, scales:{x:{grid:{display:false}}, y:{beginAtZero:true, grid:{color:'#F3F4F6'}}} } });
  } else noData('chartPayZone');

  destroyChart('chartLiftZone');
  if(data.length && zl.some(z=>sum(gz[z],r=>r[COL.liftedQty])>0 || sum(gz[z],r=>r[COL.balanceQty])>0)){
    CHARTS.chartLiftZone = new Chart(document.getElementById('chartLiftZone'),{ type:'bar',
      data:{ labels:zl, datasets:[
        { label:'Lifted Qty', data:zl.map(z=>sum(gz[z],r=>r[COL.liftedQty])), backgroundColor:'#8FB400', borderRadius:4 },
        { label:'Balance Qty', data:zl.map(z=>sum(gz[z],r=>r[COL.balanceQty])), backgroundColor:BT.amber, borderRadius:4 } ]},
      options:{ responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'bottom', labels:{boxWidth:12,padding:12}}}, scales:{x:{stacked:true, grid:{display:false}}, y:{stacked:true, beginAtZero:true, grid:{color:'#F3F4F6'}}} } });
  } else noData('chartLiftZone');

  const wrap = document.getElementById('payLiftTableWrap');
  const map = new Map();
  data.forEach(r=>{ const k=r[COL.zone]+'||'+r[COL.assetCat];
    const c = map.get(k) || {zone:r[COL.zone], cat:r[COL.assetCat], val:0, pay:0, lift:0, bal:0};
    c.val+=r[COL.auctValueTotal]||0; c.pay+=r[COL.payRs]||0; c.lift+=r[COL.liftedQty]||0; c.bal+=r[COL.balanceQty]||0;
    map.set(k,c); });
  const rows=[...map.values()].filter(r=>r.val||r.pay||r.lift||r.bal).sort((a,b)=>b.pay-a.pay);
  wrap.innerHTML = rows.length===0 ? '<div class="empty-state">No payment or lifting activity under the current filters.</div>'
    : `<table class="data-table"><thead><tr><th>Zone</th><th>Asset Category</th><th style="text-align:right">Auction Value (Rs.)</th><th style="text-align:right">Payment (Rs.)</th><th style="text-align:right">Lifted Qty</th><th style="text-align:right">Balance Qty</th></tr></thead><tbody>`+
      rows.slice(0,500).map(r=>`<tr><td>${escapeHtml(r.zone)}</td><td>${escapeHtml(r.cat)}</td><td style="text-align:right">${fmtNum(r.val)}</td><td style="text-align:right" class="pos">${fmtNum(r.pay)}</td><td style="text-align:right">${fmtNum(r.lift)}</td><td style="text-align:right" class="${r.bal>0?'neg':''}">${fmtNum(r.bal)}</td></tr>`).join('')+`</tbody></table>`;
}

/* ============ TAB 3 — auction pricing & bidder coverage (was "awaiting") ============ */
function renderAuctionPricing(data){
  const rp = sum(data,r=>r[COL.reservePriceTotal]);
  const av = sum(data,r=>r[COL.auctValueTotal]);
  const bid = data.filter(r=>(r[COL.bidder]||'').toString().trim()).length;
  const dat = data.filter(r=>(r[COL.auctionDate]||'').toString().trim()).length;
  const set=(id,v)=>{const e=document.getElementById(id); if(e) e.textContent=v;};
  set('kpi3Reserve','Rs '+fmtNum(rp)); set('kpi3Value','Rs '+fmtNum(av));
  set('kpi3Bidder',fmtNum(bid)); set('kpi3Date',fmtNum(dat));
  const bs=document.getElementById('kpi3BidderSub'); if(bs) bs.textContent = data.length? (bid/data.length*100).toFixed(1)+'% of records' : '—';
  const ds=document.getElementById('kpi3DateSub'); if(ds) ds.textContent = data.length? (dat/data.length*100).toFixed(1)+'% of records' : '—';
}

/* ==================== TAB 5 — PAYMENTS (BIDDER & DD DETAIL) ==================== */
function renderTab5(data){
  const clean = s => { s=(s==null?'':s).toString().trim(); return (s===''||s==='0')?'':s; };
  const pays = data.filter(r => (r[COL.payRs]||0) > 0 || clean(r[COL.ddNo]));
  const amt = sum(pays, r=>r[COL.payRs]);
  const set=(id,v)=>{const e=document.getElementById(id); if(e) e.textContent=v;};
  set('kpi5Amt','Rs '+fmtNum(amt)); set('kpi5Recs',fmtNum(pays.length));
  set('kpi5Bidders',fmtNum(new Set(pays.map(r=>clean(r[COL.bidder])).filter(Boolean)).size));
  set('kpi5DDs',fmtNum(new Set(pays.map(r=>clean(r[COL.ddNo])).filter(Boolean)).size));
  destroyChart('chartTopBidders');
  if(pays.length){
    const byB={}; pays.forEach(r=>{const k=((r[COL.bidder]||'').trim())||'(Unnamed bidder)'; byB[k]=(byB[k]||0)+(r[COL.payRs]||0);});
    const top=Object.entries(byB).sort((a,b)=>b[1]-a[1]).slice(0,12);
    CHARTS.chartTopBidders=new Chart(document.getElementById('chartTopBidders'),{type:'bar',
      data:{labels:top.map(x=>x[0]),datasets:[{label:'Payment Rs.',data:top.map(x=>x[1]),backgroundColor:BT.green,borderRadius:4}]},
      options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>'Rs '+fmtNum(ctx.parsed.x)}}},scales:{x:{beginAtZero:true,grid:{color:'#F3F4F6'}},y:{grid:{display:false}}}}});
  } else noData('chartTopBidders');
  destroyChart('chartPayByCat');
  if(pays.length){
    const g=groupBy(pays,COL.assetCat); const labels=Object.keys(g).sort();
    CHARTS.chartPayByCat=new Chart(document.getElementById('chartPayByCat'),{type:'bar',
      data:{labels,datasets:[{label:'Payment Rs.',data:labels.map(k=>sum(g[k],r=>r[COL.payRs])),backgroundColor:BT.slate,borderRadius:4}]},
      options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>'Rs '+fmtNum(ctx.parsed.y)}}},scales:{x:{grid:{display:false},ticks:{maxRotation:25,font:{size:10}}},y:{beginAtZero:true,grid:{color:'#F3F4F6'}}}}});
  } else noData('chartPayByCat');
  // Payments are recorded at asset-line level in the source; aggregate to one row per
  // winning bidder so the CEO sees "payments received from each bidder", not duplicated lines.
  const wrap=document.getElementById('paymentsTableWrap');
  const agg=new Map();
  pays.forEach(r=>{ const b=clean(r[COL.bidder])||'(Unnamed bidder)';
    const c=agg.get(b)||{bidder:b,total:0,lines:0,cats:new Set(),zones:new Set(),dds:new Set(),dates:new Set()};
    c.total+=r[COL.payRs]||0; c.lines++;
    if(clean(r[COL.assetCat])) c.cats.add(r[COL.assetCat]);
    if(clean(r[COL.zone])) c.zones.add(r[COL.zone]);
    if(clean(r[COL.ddNo])) c.dds.add(clean(r[COL.ddNo]));
    if(clean(r[COL.payDate])) c.dates.add(clean(r[COL.payDate]));
    agg.set(b,c); });
  const rows=[...agg.values()].sort((a,b)=>b.total-a.total);
  const j=s=>[...s].join(', ')||'—';
  wrap.innerHTML = rows.length===0 ? '<div class="empty-state">No payment or DD records under the current filters.</div>' :
    `<table class="data-table"><thead><tr><th>Winning Bidder</th><th>Zone(s)</th><th>Asset Categor(ies) Won</th><th>DD No(s).</th><th>Payment Date(s)</th><th style="text-align:right">Total Payment (Rs.)</th><th style="text-align:right">Lines</th></tr></thead><tbody>`+
    rows.map(r=>`<tr><td>${escapeHtml(r.bidder)}</td><td>${escapeHtml(j(r.zones))}</td><td>${escapeHtml(j(r.cats))}</td><td>${escapeHtml(j(r.dds))}</td><td>${escapeHtml(j(r.dates))}</td><td style="text-align:right" class="pos">${fmtNum(r.total)}</td><td style="text-align:right">${fmtNum(r.lines)}</td></tr>`).join('')+
    `</tbody></table>`;
}

/* ==================== TAB 6 — ANOMALIES ==================== */
function anomaliesOf(r){
  const out=[];
  const lift=r[COL.liftedQty]||0, auc=r[COL.auctQtyTotal]||0, cnt=r[COL.cQty]||0,
        pay=r[COL.payRs]||0, val=r[COL.auctValueTotal]||0, bal=r[COL.balanceQty]||0;
  if(lift>auc && lift>0) out.push(['Lifted > Auctioned','Lifted '+fmtNum(lift)+' vs Auctioned '+fmtNum(auc)]);
  if(pay>val && val>0)   out.push(['Payment > Auction Value','Paid Rs '+fmtNum(pay)+' vs Value Rs '+fmtNum(val)]);
  if(auc>cnt && cnt>0)   out.push(['Auctioned > Counted','Auctioned '+fmtNum(auc)+' vs Count '+fmtNum(cnt)]);
  if(bal<0)              out.push(['Negative Balance','Balance qty '+fmtNum(bal)]);
  return out;
}
function renderTab6(data){
  const flagged=[]; const byType={};
  data.forEach(r=>{ anomaliesOf(r).forEach(([t,d])=>{ flagged.push({zone:r[COL.zone],region:r[COL.region],cat:r[COL.assetCat],asset:r[COL.assetName],type:t,detail:d}); byType[t]=(byType[t]||0)+1; }); });
  const set=(id,v)=>{const e=document.getElementById(id); if(e) e.textContent=v;};
  set('kpi6Total',fmtNum(flagged.length));
  set('kpi6Lift',fmtNum(byType['Lifted > Auctioned']||0));
  set('kpi6Pay',fmtNum(byType['Payment > Auction Value']||0));
  set('kpi6Auct',fmtNum(byType['Auctioned > Counted']||0));
  destroyChart('chartAnomTypes');
  const tl=Object.keys(byType);
  if(tl.length){
    CHARTS.chartAnomTypes=new Chart(document.getElementById('chartAnomTypes'),{type:'bar',
      data:{labels:tl,datasets:[{label:'Records',data:tl.map(k=>byType[k]),backgroundColor:BT.amber,borderRadius:6,maxBarThickness:70}]},
      options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>fmtNum(ctx.parsed.y)+' records'}}},scales:{x:{grid:{display:false}},y:{beginAtZero:true,grid:{color:'#F3F4F6'}}}}});
  } else noData('chartAnomTypes');
  const wrap=document.getElementById('anomalyTableWrap');
  wrap.innerHTML = flagged.length===0 ? '<div class="empty-state">No anomalies under the current filters. &#10003;</div>' :
    `<table class="data-table"><thead><tr><th>Zone</th><th>Region</th><th>Asset Category</th><th>Asset</th><th>Anomaly</th><th>Detail (conflicting values)</th></tr></thead><tbody>`+
    flagged.slice(0,500).map(f=>`<tr><td>${escapeHtml(f.zone)}</td><td>${escapeHtml(f.region)}</td><td>${escapeHtml(f.cat)}</td><td>${escapeHtml(f.asset)}</td><td><span style="color:var(--bt-amber);font-weight:700">${escapeHtml(f.type)}</span></td><td>${escapeHtml(f.detail)}</td></tr>`).join('')+
    `</tbody></table>`;
}

/* ==================== PROJECT UPDATES & NEWS ==================== */
function renderUpdates(){
  const recs = (window.DASH_UPDATES && DASH_UPDATES.records) || [];
  const news = recs.filter(r => (r.source||'')==='News' && (r.note||r.date));
  const nl = document.getElementById('newsList');
  if(nl) nl.innerHTML = news.length ? news.map(r=>
    `<div style="padding:9px 0;border-bottom:1px solid #F3F4F6"><div style="font-size:11px;color:var(--bt-muted);font-weight:600">${escapeHtml(r.date||'')}</div><div style="font-size:13px;color:#374151;margin-top:2px">${escapeHtml(r.note||'')}</div></div>`).join('')
    : '<div class="empty-state">No dated news items in the source.</div>';
  const st = recs.filter(r => (r.source||'')==='Project Status' && (r.zone||r.lead||r.status));
  const sl = document.getElementById('statusList');
  if(sl) sl.innerHTML = st.length ?
    `<table class="data-table"><thead><tr><th>Zone</th><th>Region</th><th>Lead</th><th>Status / Note</th></tr></thead><tbody>`+
    st.map(r=>`<tr><td>${escapeHtml(r.zone||'')}</td><td>${escapeHtml(r.region||'')}</td><td>${escapeHtml(r.lead||'')}</td><td>${escapeHtml(r.status||r.note||'—')}</td></tr>`).join('')+
    `</tbody></table>` : '<div class="empty-state">No status records.</div>';
}

/* ---------- Initial empty state ---------- */"""
sub(r"/\* ---------- Initial empty state ---------- \*/", NEW_JS, label="inject tab0/tab4 js")

# 10b) Project Updates & News panel on the Overall tab (tab1)
PANEL = ('\n          <div class="grid grid-cols-1 lg:grid-cols-2 gap-4 mt-4">\n'
         '            <div class="chart-card"><div class="chart-title">Latest News</div>'
         '<div class="chart-sub">Dated updates from the Project Updates workbook</div>'
         '<div id="newsList"></div></div>\n'
         '            <div class="chart-card"><div class="chart-title">Zone Leads &amp; Field Status</div>'
         '<div class="chart-sub">Assigned leads and status by zone / region</div>'
         '<div id="statusList" class="data-table-wrap" style="max-height:360px;"></div></div>\n'
         '          </div>\n        ')
sub(r'(<canvas id="chartByCat"></canvas></div>\s*</div>)\s*</section>',
    lambda m: m.group(1) + PANEL + '</section>', flags=re.DOTALL, label="updates panel")

# 11) bootstrap from embedded JSON (replace the bootEmpty() call)
BOOT = r"""(function(){
  const meta = JSON.parse(document.getElementById('dashboardData').textContent);
  window.DASH_META = meta.meta; window.DASH_UPDATES = meta.updates;
  const cols = meta.cols;
  RAW = meta.rows.map(a => { const o={}; cols.forEach((c,i)=>{ o[COL[c]] = a[i]; });
    // ensure keys used by render code exist
    o[COL.itemCat]=o[COL.itemCat]||''; return o; });
  hydrateSlicers(); renderAll(); renderUpdates();
  const d = meta.meta.refreshDate || '';
  setStatus(`Live · ${meta.meta.rowCount.toLocaleString()} records · refreshed ${d}`, true);
})();"""
sub(r"bootEmpty\(\);", BOOT, label="bootstrap")

# 12) inject dashboardData JSON before the main app <script>
DATA_TAG = ('<script id="dashboardData" type="application/json">'
            + json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
            + "</script>\n<script>")
_data_repl = DATA_TAG + "\n/* ============================================================\n   USC Asset Disposal"
sub(r"<script>\s*\n/\* =+\s*\n\s*USC Asset Disposal",
    lambda m: _data_repl, flags=re.DOTALL, label="inject data tag")

with open(OUT_HTML, "w", encoding="utf-8") as f:
    f.write(html)
print(f"[build] wrote {OUT_HTML} ({len(html):,} bytes)")
