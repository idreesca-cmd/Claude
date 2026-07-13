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
# NEW SCHEMA (2026-07 restructure): header NAMES are on ROW 2, data starts ROW 3.
# Column C = "Assets Class" (main category), D = "Assets Category" (sub-category).
wb = openpyxl.load_workbook(SRC_XLSX, data_only=True, read_only=True)
db = wb["Database"]
ci = openpyxl.utils.column_index_from_string
def C(row, letter): return row[ci(letter) - 1]

# dashboard key -> source column letter (Row-2 names). Order here == baked array order.
COLMAP = [
    ("zone","A"), ("region","B"), ("assetClass","C"), ("assetCat","D"), ("assetName","E"),
    ("uom","G"), ("aQty","I"), ("bQty","J"), ("cQty","K"), ("dBA","L"), ("eCB","M"),
    ("reservePriceTotal","R"), ("status","S"), ("auctionDate","T"),
    ("auctQtyUsable","U"), ("auctQtyScrap","V"), ("auctQtyTotal","W"), ("auctQtyTotalKg","X"),
    ("diffAuctCount","Y"), ("bidder","Z"), ("auctValueTotal","AE"), ("liftingStatus","AF"),
    ("payDate","AG"), ("payRs","AH"), ("ddNo","AI"), ("liftedQty","AM"), ("liftedKg","AQ"),
    ("qtyDiscrepancy","AR"), ("kgsDiscrepancy","AS"), ("balanceQty","AT"), ("balanceKg","AU"),
    ("liftDate","AJ"),
]
TEXT_KEYS = {"zone","region","assetClass","assetCat","assetName","uom","status",
             "liftingStatus","bidder","ddNo"}
DATE_KEYS = {"auctionDate","payDate","liftDate"}
ORDER = [k for k, _ in COLMAP] + ["itemCat"]   # itemCat (slicer) is derived = assetClass

# Category label normalization (every mapping logged). Keep OB and M&B distinct — here the
# only variant is "Spices" -> "OB - Spices" (both are Own-Brand spices under Inventories;
# no M&B label exists, so nothing is merged across sub-groups).
CAT_NORMALIZE = {"Spices": "OB - Spices"}
_cat_norm_count = 0

def parse_date_iso(v):
    """YYYY-MM-DD or '' — handles Excel datetimes, D-M-YYYY / D/M/YYYY text, and treats
    time-only junk (e.g. '00:00:00') as MISSING."""
    if v is None or v == "": return ""
    if hasattr(v, "year") and hasattr(v, "month"):
        try: return (v.date() if hasattr(v, "date") else v).isoformat()
        except Exception: return ""
    s = str(v).strip()
    if not s: return ""
    if re.fullmatch(r"\d{1,2}:\d{2}(:\d{2})?", s): return ""            # time-only junk
    m = re.fullmatch(r"(\d{1,2})[\-/.](\d{1,2})[\-/.](\d{2,4})", s)      # D-M-YYYY / D/M/YYYY
    if m:
        d, mo, y = (int(x) for x in m.groups())
        if y < 100: y += 2000
        try: return datetime.date(y, mo, d).isoformat()
        except ValueError:
            try: return datetime.date(y, d, mo).isoformat()
            except ValueError: return ""
    try: return datetime.date.fromisoformat(s[:10]).isoformat()
    except Exception: return ""

rows_out = []
for row in db.iter_rows(min_row=3, values_only=True):
    if row is None: continue
    zone = norm_zone(C(row, "A"))
    cls  = str(C(row, "C") or "").strip()
    name = C(row, "E")
    if not (zone or cls or name):  # skip fully-blank
        continue
    if not cls:                    # need an Assets Class to place the row
        continue
    rec = []
    for key, letter in COLMAP:
        v = C(row, letter)
        if key in DATE_KEYS:
            rec.append(parse_date_iso(v))
        elif key == "assetCat":
            cat = str(v or "").strip()
            if cat in CAT_NORMALIZE:
                cat = CAT_NORMALIZE[cat]; _cat_norm_count += 1
            rec.append(cat)
        elif key in TEXT_KEYS:
            rec.append("" if v is None else str(v).strip())
        else:
            rec.append(num(v))
    rec.append(cls)                # itemCat = Assets Class
    rows_out.append(rec)

idx = {k: i for i, k in enumerate(ORDER)}
def gsum(k): return sum(r[idx[k]] for r in rows_out)

# main-category -> sub-category groups, driven by the data (Assets Class -> Assets Category)
import collections as _c
_MAIN_ORDER = ["Fixed Assets", "Inventories", "IT Equipment"]
_pair = _c.Counter((r[idx["assetClass"]], r[idx["assetCat"]]) for r in rows_out)
_mains = _MAIN_ORDER + [m for m in dict.fromkeys(r[idx["assetClass"]] for r in rows_out)
                        if m not in _MAIN_ORDER]
GROUPS = []
for _m in _mains:
    _subs = []
    for (cl, sc), _n in _pair.most_common():
        if cl == _m and sc and sc not in _subs:
            _subs.append(sc)
    if _subs:
        GROUPS.append({"main": _m, "subs": _subs})

SOURCE_PATH_LABEL = (r"G:\Shared drives\Clients (except pvt ltd co)\Advisory\Clients - Idrees"
                      r"\Utility Stores Corporation - FAR Auction\USC Dashboard\USD DB MIK"
                      r"\USC_Dashboard_Data_Source_File.xlsx")
PKT = datetime.timezone(datetime.timedelta(hours=5))  # Pakistan Standard Time, no DST
_now = datetime.datetime.now(PKT)
meta = {
    "refreshDate": _now.strftime("%Y-%m-%d %H:%M PKT"),
    "sourcePath": SOURCE_PATH_LABEL,
    "rowCount": len(rows_out),
    "zones": sorted({r[idx["zone"]] for r in rows_out}),
    "groups": GROUPS,
    "dq": {"categoriesNormalized": _cat_norm_count, "categoryMap": CAT_NORMALIZE},
    "totals": {k: gsum(k) for k in
               ("aQty","bQty","cQty","auctQtyTotal","reservePriceTotal",
                "auctValueTotal","payRs","liftedQty","liftedKg","balanceQty","balanceKg")},
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
with open("node_modules/papaparse/papaparse.min.js", encoding="utf-8") as f:
    papacss = f.read()
pp_tag = ("<script>/* PapaParse 5.4.1 — inlined so 'Refresh Data' works fully offline */\n"
          + papacss.replace("</script>", "<\\/script>") + "\n</script>")
sub(r'<script src="https://cdn\.jsdelivr\.net/npm/papaparse@[^"]*"></script>',
    lambda m: pp_tag, label="inline papaparse")
with open("node_modules/xlsx/dist/xlsx.full.min.js", encoding="utf-8") as f:
    xlsxjs = f.read()
xlsx_tag = ("<script>/* SheetJS (xlsx) 0.18.5 — inlined so 'Refresh Data' works fully offline */\n"
            + xlsxjs.replace("</script>", "<\\/script>") + "\n</script>")
sub(r'<script src="https://cdn\.jsdelivr\.net/npm/xlsx@[^"]*"></script>',
    lambda m: xlsx_tag, label="inline xlsx")
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
 .ms-table{border-collapse:separate;border-spacing:0;min-width:1680px;width:100%;}
 .ms-table th,.ms-table td{padding:11px 9px;border-bottom:1px solid var(--bt-border);vertical-align:top;}
 .ms-table thead th{background:var(--bt-charcoal);color:#fff;font-size:12px;font-weight:600;text-align:center;border-bottom:none;}
 .ms-table thead th.ms-stage-h{text-align:left;border-top-left-radius:10px;vertical-align:bottom;}
 .ms-table thead th.ms-total-h{border-top-right-radius:10px;background:#1b1f26;vertical-align:bottom;}
 .ms-table thead th.ms-total-h small{display:block;font-weight:500;color:#9CA3AF;font-size:10px;margin-top:2px;}
 .ms-main{background:#1b1f26 !important;border-left:3px solid var(--bt-lime-deep);font-size:12.5px;font-weight:700;letter-spacing:.02em;text-transform:uppercase;}
 .ms-sub{font-weight:600;font-size:11.5px;background:var(--bt-charcoal) !important;border-top:1px solid #3a4149;}
 .ms-stage{font-weight:700;color:var(--bt-slate);font-size:13px;white-space:nowrap;background:var(--bt-mint);}
 .ms-stage small{display:block;font-weight:500;color:var(--bt-muted);font-size:11px;margin-top:2px;}
 .ms-cell{text-align:center;min-width:132px;}
 .ms-cell.ms-total{background:var(--bt-mint);border-left:2px solid var(--bt-lime-deep);}
 /* horizontal comparison bars (Reconciliation + Auction rows) */
 .hbars{display:flex;flex-direction:column;gap:7px;text-align:left;}
 .hbar-top{display:flex;justify-content:space-between;align-items:baseline;font-size:10px;line-height:1.2;gap:6px;}
 .hbar-label{color:var(--bt-muted);font-weight:600;white-space:nowrap;}
 .hbar-val{color:var(--bt-slate);font-weight:700;white-space:nowrap;}
 .hbar-track{height:9px;background:#EDEFF2;border-radius:5px;overflow:hidden;margin-top:2px;}
 .hbar-fill{height:100%;border-radius:5px;background:var(--bt-lime-deep);}
 .hbar-fill.ref{background:var(--bt-charcoal);}
 .hbar-fill.anom{background:var(--bt-amber);}
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
TAB1_CSS = """
 /* Overall Project Status — narrative redesign */
 .story-intro{background:linear-gradient(90deg,var(--bt-mint),#fff);border-left:4px solid var(--bt-lime-deep);border-radius:10px;padding:14px 18px;margin-bottom:18px;font-size:13.5px;color:#374151;line-height:1.6;}
 .story-intro b{color:var(--bt-slate);}
 .sort-toggle{display:inline-flex;gap:4px;}
 .sort-btn{font-size:11px;font-weight:600;padding:3px 10px;border-radius:6px;border:1px solid var(--bt-border);background:#fff;color:var(--bt-muted);cursor:pointer;}
 .sort-btn.active{background:var(--bt-charcoal);color:#fff;border-color:var(--bt-charcoal);}
 .comp-legend{display:flex;flex-direction:column;gap:5px;margin-top:12px;}
 .comp-legend .lg-row{display:flex;align-items:center;gap:8px;font-size:12px;}
 .comp-legend .lg-sw{width:12px;height:12px;border-radius:3px;flex:none;}
 .comp-legend .lg-name{flex:1;color:#374151;}
 .comp-legend .lg-val{font-weight:700;color:var(--bt-slate);}
 .comp-legend .lg-pct{color:var(--bt-muted);width:56px;text-align:right;}
 .status-pill{display:inline-block;padding:2px 9px;border-radius:999px;font-size:11px;font-weight:700;color:#fff;white-space:nowrap;}
 .dq-note{background:#FFFBEB;border:1px solid #FDE68A;border-radius:10px;padding:12px 16px;font-size:12px;color:#78350F;line-height:1.7;margin-top:6px;}
 .dq-note b{color:#92400E;}
 /* Database tab */
 .db-toolbar{display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin-bottom:12px;}
 .db-toolbar input,.db-toolbar select{font-size:12px;padding:6px 10px;border:1px solid var(--bt-border);border-radius:7px;background:#fff;color:var(--bt-slate);}
 .db-toolbar input{min-width:230px;flex:1;}
 .db-table{border-collapse:separate;border-spacing:0;width:100%;font-size:12px;}
 .db-table thead th{position:sticky;top:0;z-index:2;background:var(--bt-charcoal);color:#fff;font-weight:600;padding:9px 10px;text-align:left;cursor:pointer;white-space:nowrap;user-select:none;}
 .db-table thead th.num{text-align:right;}
 .db-table thead th .arw{opacity:.45;font-size:10px;margin-left:4px;}
 .db-table thead th.sorted{background:#1b1f26;}
 .db-table thead th.sorted .arw{opacity:1;}
 .db-table td{padding:7px 10px;border-bottom:1px solid var(--bt-border);white-space:nowrap;color:#374151;}
 .db-table tbody tr:nth-child(even){background:#FAFBFC;}
 .db-table td.num{text-align:right;font-variant-numeric:tabular-nums;}
 .db-count{font-size:11.5px;color:var(--bt-muted);margin-top:8px;}
"""
sub(r"</style>", TAB0_CSS + TAB1_CSS + " #sourcePath{white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:520px}</style>", label="tab0/tab1 css")

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
            <div class="chart-sub">Asset categories are grouped under main categories (Fixed Assets, Inventories, IT Equipment); the last column totals all categories. Figures respect the global slicers. <span class="dq-flag">&#9650;</span> marks values &gt;100% (source data-quality anomalies, shown unclamped).</div>
            <div id="milestoneMatrix" style="overflow-x:auto;"></div>
          </div>
          <div class="chart-card mt-4">
            <div class="chart-title">Data Quality &amp; Methodology Notes</div>
            <div class="chart-sub">How each row is measured, and known source anomalies (shown unclamped for the client to reconcile).</div>
            <ul style="font-size:12.5px;color:#374151;line-height:1.7;margin:4px 0 0 18px;list-style:disc;">
              <li><b>Reconciliation</b> (horizontal bars) &mdash; A. Odoo qty (system of record, the reference), B. Physical Lists (% shown vs A. Odoo), C. Physical Count (% shown vs B. Lists).</li>
              <li><b>Auction</b> &mdash; ring = Auctioned Qty &divide; Physical Count; the bars compare <b>Total Reserve Price</b> vs <b>Total Auction Value</b> (Rs), with auction value shown as a % of reserve. <b>Lifting</b> = Lifted Qty &divide; Auctioned Qty. <b>Payment</b> = Payment Received &divide; Total Auction Value.</li>
              <li>Values marked <span class="dq-flag">&#9650;</span> exceed 100% &mdash; genuine source inconsistencies (e.g. count &gt; lists, lifted &gt; auctioned, payment/auction-value &gt; reserve). Figures are shown exactly as recorded and are <b>not clamped</b>; they flag records USC should reconcile at source.</li>
              <li>Main-category grouping (<b>Assets Class</b>) and its sub-categories (<b>Assets Category</b>) are read directly from the source file, so the columns update automatically whenever the client re-classifies assets. Current: <b>Fixed Assets</b> = F&amp;F/Boards &amp; Office Equipment, Vehicles; <b>Inventories</b> = Spices, Rice &amp; Pulses; <b>IT Equipment</b> = General IT, ERP IT.</li>
            </ul>
          </div>
        </section>

        <section id="tab1" class="tab-pane">"""
sub(r'<section id="tab1" class="tab-pane active">', TAB0_SECTION, label="tab0 section")

# 4b) tab1 — REDESIGN as a top-to-bottom narrative: verified base → advertised →
#     auctioned → lifted → stuck (and for how long). Replaces the whole old tab1 markup
#     (the chartStatus / chartByRegion / chartByCat cards). renderTab1 is overridden in NEW_JS.
TAB1_SECTION = """<section id="tab1" class="tab-pane">
          <div class="story-intro">
            The disposal programme end-to-end: from the <b>verified base</b> we physically counted, to what has been
            <b>put to auction</b>, <b>lifted</b> by buyers, and what remains <b>stuck awaiting lifting</b> &mdash; and for how long.
            Figures respect the global Zone / Region slicers. Consignment stock is shown for context but excluded from auction &amp; lifting rates.
          </div>
          <div class="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-4 mb-6">
            <div class="kpi-card"><div class="kpi-label">Verified Base Qty</div><div id="k1Total" class="kpi-value">&mdash;</div><div id="k1TotalSub" class="kpi-sub">Physically counted units</div></div>
            <div class="kpi-card"><div class="kpi-label">Auction in Process</div><div id="k1Proc" class="kpi-value">&mdash;</div><div id="k1ProcSub" class="kpi-sub">Advertised, not concluded</div></div>
            <div class="kpi-card"><div class="kpi-label">Auctioned Qty</div><div id="k1Auct" class="kpi-value">&mdash;</div><div id="k1AuctSub" class="kpi-sub">&mdash; of verified base</div></div>
            <div class="kpi-card"><div class="kpi-label">Lifted Qty</div><div id="k1Lift" class="kpi-value">&mdash;</div><div id="k1LiftSub" class="kpi-sub">&mdash; of auctioned</div></div>
            <div class="kpi-card"><div class="kpi-label">Lifting Pending Qty</div><div id="k1Pend" class="kpi-value">&mdash;</div><div id="k1PendSub" class="kpi-sub">&mdash; of auctioned</div></div>
            <div class="kpi-card"><div class="kpi-label">Consignment Qty</div><div id="k1Cons" class="kpi-value">&mdash;</div><div id="k1ConsSub" class="kpi-sub">Excluded from rates</div></div>
          </div>
          <div class="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-4">
            <div class="chart-card">
              <div class="chart-title">Where the verified base stands today</div>
              <div class="chart-sub">Every counted unit by its derived disposal status (Consignment included for context). Tooltip shows qty &amp; share.</div>
              <div style="position:relative;height:300px;"><canvas id="chartComposition"></canvas></div>
              <div id="compLegend" class="comp-legend"></div>
            </div>
            <div class="chart-card">
              <div class="chart-title">Of what we auctioned, how much is actually lifted?</div>
              <div class="chart-sub">Auctioned quantity split into lifted vs still-to-lift (Consignment excluded).</div>
              <div style="position:relative;height:300px;"><canvas id="chartLiftPie"></canvas></div>
            </div>
          </div>
          <div class="chart-card mb-4">
            <div style="display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:8px;">
              <div><div class="chart-title">Which zones carry the most &mdash; and how far have they moved?</div>
                <div class="chart-sub">Verified base vs auctioned vs lifted quantity per zone.</div></div>
              <div class="sort-toggle" id="zoneSort"><button class="sort-btn active" data-dir="desc">Total &darr;</button><button class="sort-btn" data-dir="asc">Total &uarr;</button></div>
            </div>
            <div style="position:relative;height:340px;"><canvas id="chartZoneCmp"></canvas></div>
          </div>
          <div class="chart-card mb-4">
            <div style="display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:8px;">
              <div><div class="chart-title">Region-level detail, grouped by zone</div>
                <div class="chart-sub">Verified base vs auctioned vs lifted quantity per region. Default order follows each parent zone; labels are prefixed with the zone.</div></div>
              <div class="sort-toggle" id="regionSort"><button class="sort-btn active" data-dir="zone">By zone</button><button class="sort-btn" data-dir="desc">Total &darr;</button><button class="sort-btn" data-dir="asc">Total &uarr;</button></div>
            </div>
            <div style="position:relative;height:560px;"><canvas id="chartRegionCmp"></canvas></div>
          </div>
          <div class="chart-card mb-4">
            <div class="chart-title">How long has auctioned stock been stuck awaiting lifting?</div>
            <div class="chart-sub" id="agingSub">Days since auction date for auctioned-but-not-fully-lifted lots, as of today.</div>
            <div class="grid grid-cols-1 lg:grid-cols-2 gap-4 mt-2">
              <div style="position:relative;height:300px;"><canvas id="chartAging"></canvas></div>
              <div id="agingTableWrap"></div>
            </div>
          </div>
          <div class="chart-card mb-4">
            <div class="chart-title">Top 10 longest-pending lots</div>
            <div class="chart-sub">Auctioned lots with the most days elapsed since auction and still not fully lifted.</div>
            <div id="agingTopWrap" class="data-table-wrap"></div>
          </div>
          <div class="dq-note" id="dqNote"></div>
        </section>"""
sub(r'<section id="tab1" class="tab-pane">.*?</section>', lambda m: TAB1_SECTION,
    flags=re.DOTALL, label="tab1 narrative section")

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
    '        <button class="tab-btn" data-tab="tab6">Anomalies</button>\n'
    '        <button class="tab-btn" data-tab="tab7">Database</button>',
    label="nav tab5/6/7")

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
          <div class="grid grid-cols-2 gap-4 mb-5">
            <div class="kpi-card"><div class="kpi-label">Qty Discrepancy (source col)</div><div id="kpi6QtyDisc" class="kpi-value">—</div><div id="kpi6QtyDiscSub" class="kpi-sub">Source column AR</div></div>
            <div class="kpi-card"><div class="kpi-label">KGs Discrepancy (source col)</div><div id="kpi6KgsDisc" class="kpi-value">—</div><div id="kpi6KgsDiscSub" class="kpi-sub">Source column AS</div></div>
          </div>
          <div class="chart-card"><div class="chart-title">Anomalies by Type</div><div class="chart-sub">Count of records per anomaly category (respects the global slicers)</div><div style="position:relative;height:300px;"><canvas id="chartAnomTypes"></canvas></div></div>
          <div class="chart-card mt-4"><div class="chart-title">Anomaly Detail</div><div class="chart-sub">Every flagged record with the conflicting values (top 500). These are source data-quality issues for USC to reconcile — figures are shown exactly as recorded.</div><div id="anomalyTableWrap" class="data-table-wrap"></div></div>
        </section>
        <section id="tab7" class="tab-pane">
          <div class="chart-card">
            <div class="chart-title">Database &mdash; Full Asset-Line Register</div>
            <div class="chart-sub">Every asset line with its derived disposal status. Type to search, use the drop-downs to filter, and click any column header to sort ascending / descending. Respects the global Zone / Region slicers.</div>
            <div class="db-toolbar">
              <input id="dbSearch" type="text" placeholder="Search zone, region, category, asset, status, bidder&hellip;" oninput="renderTab7(getFiltered())">
              <select id="dbFZone" onchange="renderTab7(getFiltered())"><option value="">All Zones</option></select>
              <select id="dbFRegion" onchange="renderTab7(getFiltered())"><option value="">All Regions</option></select>
              <select id="dbFClass" onchange="renderTab7(getFiltered())"><option value="">All Asset Classes</option></select>
              <select id="dbFCat" onchange="renderTab7(getFiltered())"><option value="">All Categories</option></select>
              <select id="dbFStatus" onchange="renderTab7(getFiltered())"><option value="">All Statuses</option></select>
            </div>
            <div id="dbTableWrap" class="data-table-wrap" style="max-height:640px;"></div>
            <div id="dbCount" class="db-count"></div>
          </div>
        </section>
"""
lit('</main>', TAB56 + '        </main>', label="tab5/6/7 sections")

# 6) NEW SCHEMA — replace COL / NUMERIC_COLS so the manual "Refresh Data" upload (which reads
#    by header NAME) matches the restructured file's Row-2 headers. The baked-JSON boot path
#    maps by array index (handled in the Python extraction above).
NEW_COL = ("const COL = {\n"
    "  zone:'Zone', region:'Region', assetClass:'Assets Class', itemCat:'Item Category', assetCat:'Assets Category', assetName:'Asset Name (Matched)', uom:'Unit of Measure',\n"
    "  aQty:'A QTY (Odoo)', bQty:'B QTY (Physical Lists)', cQty:'C QTY (Physical Count)', dBA:'D=B-A', eCB:'E=C-B',\n"
    "  status:'Auction Status', liftingStatus:'Lifting Status', auctionDate:'Auction Date',\n"
    "  auctQtyUsable:'Auction Qty Usable', auctQtyScrap:'Auction Qty Scrap', auctQtyTotal:'Auction Qty Total', auctQtyTotalKg:'Auction Qty Total KGs', diffAuctCount:'Diff Qty Auction vs Count',\n"
    "  bidder:'Auction To Bidder',\n"
    "  rpUsable:'Total Reserve Price Usable', rpScrap:'Total Reserve Price Scrape', reservePriceTotal:'Overall Reserve Price Total',\n"
    "  apUsable:'Auction Price Usable (Per Unit)', apScrap:'Auction Price Scrape (Per Unit)', auctValueTotal:'Total Auction Value',\n"
    "  payDate:'Payment Date', payRs:'Payment Rs.', ddNo:'Payment Demand Draft (DD) No.',\n"
    "  liftDate:'Lifting Date', liftedQty:'Total Lifted Qty Units', liftedKg:'Total Lifted Qty KGs',\n"
    "  balanceQty:'Balance Qty in Units', balanceKg:'Balance Qty in KGs',\n"
    "  qtyDiscrepancy:'Qty Discrepancy', kgsDiscrepancy:'KGs Discrepancy'\n"
    "}")
sub(r"const COL = \{.*?\}", lambda m: NEW_COL, flags=re.DOTALL, label="COL object")

NEW_NUM = ("const NUMERIC_COLS = [\n"
    "  COL.aQty, COL.bQty, COL.cQty, COL.dBA, COL.eCB,\n"
    "  COL.auctQtyUsable, COL.auctQtyScrap, COL.auctQtyTotal, COL.auctQtyTotalKg, COL.diffAuctCount,\n"
    "  COL.rpUsable, COL.rpScrap, COL.reservePriceTotal, COL.apUsable, COL.apScrap, COL.auctValueTotal,\n"
    "  COL.payRs, COL.liftedQty, COL.liftedKg, COL.balanceQty, COL.balanceKg,\n"
    "  COL.qtyDiscrepancy, COL.kgsDiscrepancy\n"
    "]")
sub(r"const NUMERIC_COLS = \[.*?\]", lambda m: NEW_NUM, flags=re.DOTALL, label="NUMERIC_COLS")

# pickBestSheet: prefer header-signature (row 2 holds names now) over the name-regex,
# and read the header from row 2 (index 1) when row 1 is the marker row.
sub(r"const reconMatch = names\.find\(n => /recon/i\.test\(n\)\);\s*\n\s*if \(reconMatch\) return reconMatch;\s*\n\s*// 2\) header signature match\s*\n\s*const sig = \['Zone','Asset Name \(Matched\)'\];",
    "// header-signature match first (row 2 holds the column names in this workbook)\n"
    "  const sig = ['Zone','Assets Category','Asset Name (Matched)'];",
    label="pickBestSheet: sig")
sub(r"const headerRow = XLSX\.utils\.sheet_to_json\(ws, \{ header:1, defval:'' \}\)\[0\] \|\| \[\];",
    "const _hr = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });\n"
    "    const headerRow = (_hr[1] && _hr[1].some(x => String(x).trim())) ? _hr[1] : (_hr[0] || []);",
    label="pickBestSheet: row2 header")
sub(r"if \(sig\.every\(s => hdrs\.includes\(s\)\)\) return n;\s*\n\s*\}\s*\n\s*// 3\) first non-empty sheet",
    "if (sig.every(s => hdrs.includes(s))) return n;\n"
    "  }\n"
    "  // fall back: a sheet named like \"database\" or \"recon\"\n"
    "  const nm = names.find(n => /database|recon/i.test(n));\n"
    "  if (nm) return nm;\n"
    "  // first non-empty sheet",
    label="pickBestSheet: fallback")

# handleExcel: this workbook has a MARKER row (row 1) above the header row (row 2). Detect it
# and read with range:<hdrRow> so the header is row 2 and data starts at row 3.
sub(r"const rawRows = XLSX\.utils\.sheet_to_json\(wb\.Sheets\[sheetName\], \{ defval:'', raw:true \}\);",
    "const _probe = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header:1, defval:'' });\n"
    "      const _hdrRow = ((_probe[0]||[]).map(x=>String(x).trim()).includes('Zone')) ? 0 : 1;\n"
    "      const rawRows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval:'', raw:true, range:_hdrRow });",
    label="handleExcel: header-row detect")

# cleanRow: trim text fields + derive itemCat (the 'Asset Class' slicer) = Assets Class.
_cleanrow_repl = (
    "[COL.zone, COL.region, COL.assetClass, COL.assetCat, COL.assetName, COL.status, COL.liftingStatus, COL.bidder, COL.ddNo]\n"
    " .forEach(c => { out[c] = (out[c]===null||out[c]===undefined) ? '' : String(out[c]).trim(); });\n"
    " if((out[COL.zone]||'').toUpperCase()==='SUKKUR') out[COL.zone]='Sukkur';\n"
    " if(CAT_NORMALIZE_JS[out[COL.assetCat]]){ out[COL.assetCat]=CAT_NORMALIZE_JS[out[COL.assetCat]]; window.__dqCat=(window.__dqCat||0)+1; }\n"
    " out[COL.itemCat] = out[COL.assetClass] || '';\n"
    " return out;")
sub(r"\[COL\.zone, COL\.region, COL\.itemCat, COL\.assetCat, COL\.assetName, COL\.status, COL\.bidder, COL\.ddNo\]\s*\n\s*\.forEach\(c => \{ out\[c\] = \(out\[c\]===null\|\|out\[c\]===undefined\) \? '' : String\(out\[c\]\)\.trim\(\); \}\);\s*\n\s*return out;",
    lambda m: _cleanrow_repl, label="cleanRow normalization")

# reset the category-normalization DQ counter before each upload rebuild (cleanRow re-counts)
sub(r"RAW = rows\.map\(cleanRow\)\.filter",
    lambda m: "window.__dqCat = 0;\n RAW = rows.map(cleanRow).filter", label="reset dq counter")

# manual-upload status: append a PKT sync timestamp
sub(r"setStatus\(`\$\{RAW\.length\.toLocaleString\(\)\} records loaded from \$\{filename\}\$\{sheetSuffix\}`, true\);",
    lambda m: ("setStatus(`${RAW.length.toLocaleString()} records loaded from ${filename}${sheetSuffix}"
               " · last synced ${new Date(Date.now()+5*3600*1000).toISOString().slice(0,16).replace('T',' ')} PKT`, true);"),
    label="upload status timestamp")

# 6b) Relabel the "Item Category" slicer -> "Asset Class" (it now holds Fixed Assets /
#     Inventories / IT Equipment, the main classes from the source's Assets Class column).
sub(r'<label class="slicer-label">Item Category</label>',
    '<label class="slicer-label">Asset Class</label>', label="slicer label")
sub(r"All Item Categories", "All Asset Classes", count=2, label="slicer all-option")

# 6c) Variance Detail table (Quantity Reconciliation tab): change the two variance columns
#     from "A − B" / "B − C" to "B − A" / "C − B" (compare B vs A, C vs B — matches the
#     Reconciliation row and the source's own D=B-A / E=C-B columns). Keep "C − A".
sub(r"ab: r\.a - r\.b, bc: r\.b - r\.c, ca: r\.c - r\.a",
    lambda m: "ba: r.b - r.a, cb: r.c - r.b, ca: r.c - r.a", label="variance calc")
sub(r'<th class="text-right">A − B</th>',
    lambda m: '<th class="text-right">B − A</th>', label="variance hdr B-A")
sub(r'<th class="text-right">B − C</th>',
    lambda m: '<th class="text-right">C − B</th>', label="variance hdr C-B")
sub(r"""class="\$\{r\.ab<0\?'neg':\(r\.ab>0\?'pos':''\)\}">\$\{fmtSigned\(r\.ab\)\}""",
    lambda m: """class="${r.ba<0?'neg':(r.ba>0?'pos':'')}">${fmtSigned(r.ba)}""",
    label="variance cell B-A")
sub(r"""class="\$\{r\.bc<0\?'neg':\(r\.bc>0\?'pos':''\)\}">\$\{fmtSigned\(r\.bc\)\}""",
    lambda m: """class="${r.cb<0?'neg':(r.cb>0?'pos':'')}">${fmtSigned(r.cb)}""",
    label="variance cell C-B")

# 7) header status/label wording
sub(r"Awaiting data", "Loading…", label="status text")
sub(r'title="Upload \.xlsx, \.xls, or \.csv"', label="upload title",
    repl='title="Download the latest workbook from the USD DB MIK Drive folder, then click here to refresh this dashboard instantly with it."')
sub(r"Upload Excel / CSV", "Refresh Data (upload latest file)", label="upload label")

sub(r'(<div class="text-\[11px\] text-gray-500 leading-tight">Third-Party Validation Dashboard &middot; Baker Tilly</div>)',
    lambda m: m.group(1) + '\n <div id="sourcePath" class="text-[10px] text-gray-400 leading-tight mt-0.5"></div>',
    label="source path element")

HOWTO = ('\n <div class="mt-4 pt-4 text-[11px] text-gray-400 leading-relaxed border-t border-gray-700">\n'
         ' <div class="font-semibold text-gray-300 mb-1">How to refresh</div>\n'
         ' <div>1. Open the &quot;USD DB MIK&quot; Drive folder &rarr; download the latest'
         ' &quot;All Zones Consolidation.xlsx&quot;.<br>2. Click &quot;Refresh Data&quot; above &rarr;'
         ' select the downloaded file.<br>Everything recalculates instantly &mdash; no page reload, no login.</div>\n'
         ' </div>')
sub(r'(<div id="scopeInfo">Upload a CSV to begin\. Slicers cascade and apply to every tab\.</div>\s*\n\s*</div>)',
    lambda m: m.group(1) + HOWTO, label="how-to-refresh note")

# 8) render calls: add tab0 + tab4 … tab7 (Database)
sub(r"renderTab3\(data\);",
    "renderTab3(data);\n  renderTab0(data);\n  renderTab4(data);\n  renderAuctionPricing(data);\n  renderTab5(data);\n  renderTab6(data);\n  renderTab7(data);",
    label="renderAll calls")

# 8b) Brand-align the Chart.js palette (lime accent, matching tab0/tab4)
sub(r"green:'#00A376', greenDark:'#007F5C', greenTint:'#7DCBB3',",
    "green:'#8FB400', greenDark:'#5E7A00', greenTint:'#D1EC51',", label="BT green->lime")
sub(r"palette:\['#00A376','#2C3136','#D97706','#2563EB','#9333EA','#0891B2','#DC2626','#65A30D'\]",
    "palette:['#8FB400','#2C303B','#D97706','#0891B2','#9333EA','#2563EB','#DC2626','#00A376']",
    label="BT palette lime-first")

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
/* Main-category -> sub-category grouping is now DATA-DRIVEN: the source file provides
   "Assets Class" (main) and "Assets Category" (sub), baked into DASH_META.groups. */
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
function ringBlock(pct, ach, total){
  const flag = (pct!==null && pct>100) ? ' <span class="dq-flag" title="Exceeds 100% — source data anomaly">▲</span>' : '';
  const label = pct===null ? '—' : pct.toFixed(0)+'%';
  return `<div class="ring" style="${ringStyle(pct||0)}"><span>${label}</span></div>`+
         `<div class="ms-qty"><b>${fmtNum(ach)}</b> / ${fmtNum(total)}${flag}</div>`;
}
/* one horizontal bar: label + value(+%) above a proportional track */
function hbar(label, qty, widthPct, pctText, kind){
  const w = Math.max(2, Math.min(100, isFinite(widthPct)?widthPct:0));
  return `<div class="hbar-row"><div class="hbar-top"><span class="hbar-label">${label}</span>`+
         `<span class="hbar-val">${fmtNum(qty)}${pctText?' · '+pctText:''}</span></div>`+
         `<div class="hbar-track"><div class="hbar-fill ${kind||''}" style="width:${w}%"></div></div></div>`;
}
function bigPct(p){ return fmtNum(p)+'%'; }
/* Reconciliation cell: A Odoo (reference) -> B Lists (% vs A) -> C Count (% vs B) */
function reconCell(rows){
  const A=sum(rows,r=>r[COL.aQty]), B=sum(rows,r=>r[COL.bQty]), C=sum(rows,r=>r[COL.cQty]);
  const bp = A>0 ? B/A*100 : null, cp = B>0 ? C/B*100 : null;
  return '<div class="hbars">'+
    hbar('A · Odoo', A, 100, 'base', 'ref')+
    hbar('B · Lists', B, bp===null?0:bp, bp===null?'—':bp.toFixed(0)+'% vs A', bp!==null&&bp>100?'anom':'')+
    hbar('C · Count', C, cp===null?0:cp, cp===null?'—':cp.toFixed(0)+'% vs B', cp!==null&&cp>100?'anom':'')+
    '</div>';
}
/* Auction cell: donut (Count -> Auctioned) + horizontal Reserve vs Auction Value (Rs) */
function auctionCell(rows){
  const total=sum(rows,r=>r[COL.cQty]), ach=sum(rows,r=>r[COL.auctQtyTotal]);
  const pct = total>0 ? ach/total*100 : null;
  const rp=sum(rows,r=>r[COL.reservePriceTotal]), av=sum(rows,r=>r[COL.auctValueTotal]);
  const mx=Math.max(rp,av,1);
  const bars='<div class="hbars" style="margin-top:8px">'+
    hbar('Reserve (Rs)', rp, rp/mx*100, '', 'ref')+
    hbar('Auction (Rs)', av, av/mx*100, rp>0?bigPct(av/rp*100)+' of res.':'—', rp>0&&av>rp?'anom':'')+
    '</div>';
  return ringBlock(pct, ach, total)+bars;
}
function liftingCell(rows){
  const total=sum(rows,r=>r[COL.auctQtyTotal]), ach=sum(rows,r=>r[COL.liftedQty]);
  return ringBlock(total>0?ach/total*100:null, ach, total)+paBars(total, ach);
}
function paymentCell(rows){
  const total=sum(rows,r=>r[COL.auctValueTotal]), ach=sum(rows,r=>r[COL.payRs]);
  return ringBlock(total>0?ach/total*100:null, ach, total)+paBars(total, ach);
}
const MS_STAGES = [
  { name:'Reconciliation', sub:'Odoo → Lists → Count', cell:reconCell },
  { name:'Auction', sub:'Reserve vs Auction Value', cell:auctionCell },
  { name:'Lifting', sub:'Auctioned → Lifted', cell:liftingCell },
  { name:'Payment', sub:'Auction Value → Received', cell:paymentCell },
];
function renderTab0(data){
  const groups = (window.DASH_META && DASH_META.groups) || [];
  const subs = groups.flatMap(g=>g.subs);              // flat list of sub-category names
  const byCat = {}; subs.forEach(s=> byCat[s] = data.filter(r=>r[COL.assetCat]===s));
  let h = '<table class="ms-table"><thead><tr><th rowspan="2" class="ms-stage-h">Milestone Stage</th>';
  groups.forEach(g=>{ h += `<th colspan="${g.subs.length}" class="ms-main">${escapeHtml(g.main)}</th>`; });
  h += '<th rowspan="2" class="ms-total-h">Overall Total<small>All categories</small></th></tr><tr>';
  subs.forEach(s=>{ h += `<th class="ms-sub">${escapeHtml(s)}</th>`; });
  h += '</tr></thead><tbody>';
  MS_STAGES.forEach(st=>{
    h += `<tr><td class="ms-stage">${st.name}<small>${st.sub}</small></td>`;
    subs.forEach(s=>{ h += `<td class="ms-cell">${st.cell(byCat[s])}</td>`; });
    h += `<td class="ms-cell ms-total">${st.cell(data)}</td></tr>`;
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
  // Source discrepancy columns (AR/AS) — the user marked them for the dashboard but they are
  // currently empty in the source; show the total if populated, else an "awaiting data" note.
  const qd=sum(data,r=>r[COL.qtyDiscrepancy]), kd=sum(data,r=>r[COL.kgsDiscrepancy]);
  const nq=data.filter(r=>(r[COL.qtyDiscrepancy]||0)!==0).length, nk=data.filter(r=>(r[COL.kgsDiscrepancy]||0)!==0).length;
  set('kpi6QtyDisc', nq? fmtNum(qd) : '—');
  set('kpi6KgsDisc', nk? fmtNum(kd) : '—');
  const qs=document.getElementById('kpi6QtyDiscSub'); if(qs) qs.textContent = nq? `${fmtNum(nq)} rows populated` : 'Awaiting data in source (col AR)';
  const ks=document.getElementById('kpi6KgsDiscSub'); if(ks) ks.textContent = nk? `${fmtNum(nk)} rows populated` : 'Awaiting data in source (col AS)';
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

/* ==================== ZONE SCORECARD (Overall tab) ==================== */
function pctCell(p){
  if(p===null||isNaN(p)) return '<td style="text-align:right;color:#9CA3AF">—</td>';
  const w=Math.max(0,Math.min(100,p)); const anom=p>100;
  const col=anom?'var(--bt-amber)':(p>=66?'#5E7A00':(p>=33?'#8FB400':'#B8860B'));
  return '<td style="text-align:right;position:relative;min-width:78px">'+
    '<div style="position:absolute;left:0;top:3px;bottom:3px;width:'+w+'%;background:var(--bt-mint);opacity:.7;border-radius:3px"></div>'+
    '<span style="position:relative;font-weight:700;color:'+col+'">'+p.toFixed(0)+'%'+(anom?' ▲':'')+'</span></td>';
}
function zStat(rows){
  const a=sum(rows,r=>r[COL.aQty]),c=sum(rows,r=>r[COL.cQty]),au=sum(rows,r=>r[COL.auctQtyTotal]),
        li=sum(rows,r=>r[COL.liftedQty]),val=sum(rows,r=>r[COL.auctValueTotal]),pay=sum(rows,r=>r[COL.payRs]);
  return {n:rows.length, pv:a?c/a*100:null, au:c?au/c*100:null, li:au?li/au*100:null, pay:val?pay/val*100:null, payRs:pay};
}
function renderZoneScorecard(data){
  const el=document.getElementById('zoneScorecard'); if(!el) return;
  const g=groupBy(data,COL.zone); const zones=Object.keys(g).sort();
  const row=(name,s,foot)=>`<tr${foot?' style="background:#F3F4F6;font-weight:700"':''}><td><b>${escapeHtml(name)}</b></td>`+
    `<td style="text-align:right">${fmtNum(s.n)}</td>${pctCell(s.pv)}${pctCell(s.au)}${pctCell(s.li)}${pctCell(s.pay)}`+
    `<td style="text-align:right">${fmtNum(s.payRs)}</td></tr>`;
  let h='<table class="data-table"><thead><tr><th>Zone</th><th style="text-align:right">Records</th>'+
    '<th style="text-align:right">Phys. Verif.</th><th style="text-align:right">Auction</th>'+
    '<th style="text-align:right">Lifting</th><th style="text-align:right">Payment</th>'+
    '<th style="text-align:right">Payment (Rs.)</th></tr></thead><tbody>';
  zones.forEach(z=>{ h+=row(z,zStat(g[z]),false); });
  h+='</tbody><tfoot>'+row('All Zones',zStat(data),true)+'</tfoot></table>';
  el.innerHTML=h;
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

/* ==================== DERIVED DISPOSAL-STATUS MODEL ==================== */
/* One fixed colour per status across every chart + the Database pills. */
const STATUS_ORDER = ['Lifted','Partially Lifted','Lifting Pending','Auction in Process','Not Auctioned','Consignment'];
const STATUS_COLORS = {
  'Lifted':'#8FB400', 'Partially Lifted':'#C9DE7A', 'Lifting Pending':'#D97706',
  'Auction in Process':'#2563EB', 'Not Auctioned':'#6B7280', 'Consignment':'#9333EA'
};
const MEASURE_COLORS = { total:'#2C303B', auction:'#0891B2', lifted:'#8FB400' };
const CAT_NORMALIZE_JS = {"Spices":"OB - Spices"};
function _s(v){ return (v==null?'':String(v)).trim().toLowerCase(); }
/* Precedence: Consignment → Auction-in-Process → Not-Auctioned → (Auctioned:) Lifted /
   Partially Lifted / Lifting Pending. The DERIVED status governs — the Lifting-Status text
   column is never trusted on its own (it contradicts ~itself on some auctioned rows). */
function deriveStatus(r){
  const st=_s(r[COL.status]), lst=_s(r[COL.liftingStatus]);
  const atot=+(r[COL.auctQtyTotal]||0), lifted=+(r[COL.liftedQty]||0);
  if(st==='consignment' || lst==='consignment') return 'Consignment';
  if(/process|progress|advertis|await/.test(st)) return 'Auction in Process';
  if(st!=='auctioned') return 'Not Auctioned';
  if(atot<=0) return lifted>0 ? 'Lifted' : 'Lifting Pending';   // fallback for atot=0 rows
  if(lifted>=atot) return 'Lifted';
  if(lifted>0) return 'Partially Lifted';
  return 'Lifting Pending';
}
function isAuctionedStatus(s){ return s==='Lifted'||s==='Partially Lifted'||s==='Lifting Pending'; }
/* Robust date parse for baked ISO strings, Date objects, Excel serials and D-M-Y text. */
function parseDateJS(v){
  if(v==null||v==='') return null;
  if(v instanceof Date) return isNaN(v.getTime())?null:v;
  if(typeof v==='number'){ if(v>20000&&v<80000){ const d=new Date(Date.UTC(1899,11,30)+v*86400000); return isNaN(d.getTime())?null:d; } return null; }
  const s=String(v).trim(); if(!s) return null;
  if(/^\d{1,2}:\d{2}(:\d{2})?$/.test(s)) return null;                 // time-only junk → missing
  let m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);                      // ISO
  if(m){ const d=new Date(+m[1],+m[2]-1,+m[3]); return isNaN(d.getTime())?null:d; }
  m=s.match(/^(\d{1,2})[\-\/.](\d{1,2})[\-\/.](\d{2,4})/);            // D-M-Y / D/M/Y
  if(m){ let y=+m[3]; if(y<100) y+=2000; const d=new Date(y,+m[2]-1,+m[1]); return isNaN(d.getTime())?null:d; }
  const d=new Date(s); return isNaN(d.getTime())?null:d;
}
function daysSince(d){ return Math.floor((Date.now()-d.getTime())/86400000); }

/* ==================== TAB 1 — OVERALL PROJECT STATUS (narrative) ==================== */
let T1_DATA=[]; let ZONE_SORT='desc'; let REGION_SORT='zone';
function renderTab1(data){
  T1_DATA=data;
  data.forEach(r=>{ r.__st=deriveStatus(r); });
  renderT1KPIs(data); renderComposition(data); renderLiftPie(data);
  renderZoneCmp(data); renderRegionCmp(data); renderAging(data); renderDQNote(data);
  wireT1Toggles();
}
function renderT1KPIs(data){
  const set=(id,v)=>{const e=document.getElementById(id); if(e) e.textContent=v;};
  const q=r=>+(r[COL.cQty]||0);
  const totalBase=sum(data,q);
  const cons=data.filter(r=>r.__st==='Consignment'), consQty=sum(cons,q);
  const proc=data.filter(r=>r.__st==='Auction in Process'), procQty=sum(proc,q);
  const auct=data.filter(r=>isAuctionedStatus(r.__st));
  const auctQty=sum(auct,r=>+(r[COL.auctQtyTotal]||0));
  const liftedQty=sum(auct,r=>+(r[COL.liftedQty]||0));
  const pendQty=sum(auct,r=>Math.max(0,(+(r[COL.auctQtyTotal]||0))-(+(r[COL.liftedQty]||0))));
  set('k1Total',fmtNum(totalBase)); set('k1TotalSub',`${fmtNum(data.length)} counted asset lines`);
  set('k1Proc', procQty>0?fmtNum(procQty):'—'); set('k1ProcSub', proc.length?`${fmtNum(proc.length)} lots advertised`:'Advertised, not concluded');
  set('k1Auct',fmtNum(auctQty)); set('k1AuctSub', totalBase>0?`${(auctQty/totalBase*100).toFixed(1)}% of verified base`:'—');
  set('k1Lift',fmtNum(liftedQty)); set('k1LiftSub', auctQty>0?`${(liftedQty/auctQty*100).toFixed(1)}% of auctioned`:'—');
  set('k1Pend',fmtNum(pendQty)); set('k1PendSub', auctQty>0?`${(pendQty/auctQty*100).toFixed(1)}% of auctioned`:'—');
  set('k1Cons', consQty>0?fmtNum(consQty):'—'); set('k1ConsSub', cons.length?`${fmtNum(cons.length)} lots on consignment`:'Excluded from rates');
}
function renderComposition(data){
  destroyChart('chartComposition');
  const leg=document.getElementById('compLegend');
  const q=r=>+(r[COL.cQty]||0); const by={}; STATUS_ORDER.forEach(s=>by[s]=0);
  data.forEach(r=>{ by[r.__st]=(by[r.__st]||0)+q(r); });
  const labels=STATUS_ORDER.filter(s=>by[s]>0), vals=labels.map(s=>by[s]);
  const tot=vals.reduce((a,b)=>a+b,0);
  if(!data.length||tot<=0){ noData('chartComposition'); if(leg) leg.innerHTML=''; return; }
  CHARTS.chartComposition=new Chart(document.getElementById('chartComposition'),{type:'doughnut',
    data:{labels,datasets:[{data:vals,backgroundColor:labels.map(s=>STATUS_COLORS[s]),borderColor:'#fff',borderWidth:2}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:'62%',
      plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>`${ctx.label}: ${fmtNum(ctx.parsed)} (${(ctx.parsed/tot*100).toFixed(1)}%)`}}}}});
  if(leg) leg.innerHTML=labels.map((s,i)=>`<div class="lg-row"><span class="lg-sw" style="background:${STATUS_COLORS[s]}"></span><span class="lg-name">${s}</span><span class="lg-val">${fmtNum(vals[i])}</span><span class="lg-pct">${(vals[i]/tot*100).toFixed(1)}%</span></div>`).join('');
}
function renderLiftPie(data){
  destroyChart('chartLiftPie');
  const auct=data.filter(r=>isAuctionedStatus(r.__st));
  const auctQty=sum(auct,r=>+(r[COL.auctQtyTotal]||0));
  const lifted=sum(auct,r=>+(r[COL.liftedQty]||0));
  const pend=Math.max(0,auctQty-lifted);
  if(!auct.length||auctQty<=0){ noData('chartLiftPie'); return; }
  const vals=[lifted,pend];
  CHARTS.chartLiftPie=new Chart(document.getElementById('chartLiftPie'),{type:'doughnut',
    data:{labels:['Lifted','Still to lift'],datasets:[{data:vals,backgroundColor:['#8FB400','#D97706'],borderColor:'#fff',borderWidth:2}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:'60%',
      plugins:{title:{display:true,text:'Auctioned base: '+fmtNum(auctQty)+' units',color:'#6B7280',font:{size:12,weight:'600'}},
        legend:{position:'bottom',labels:{boxWidth:12,padding:12}},
        tooltip:{callbacks:{label:ctx=>`${ctx.label}: ${fmtNum(ctx.parsed)} (${(ctx.parsed/auctQty*100).toFixed(1)}% of auctioned)`}}}}});
}
function renderZoneCmp(data){
  destroyChart('chartZoneCmp');
  if(!data.length){ noData('chartZoneCmp'); return; }
  const g=groupBy(data,COL.zone); let zones=Object.keys(g);
  const tot=z=>sum(g[z],r=>+(r[COL.cQty]||0));
  zones.sort((a,b)=> ZONE_SORT==='asc'? tot(a)-tot(b): tot(b)-tot(a));
  CHARTS.chartZoneCmp=new Chart(document.getElementById('chartZoneCmp'),{type:'bar',
    data:{labels:zones,datasets:[
      {label:'Verified base',data:zones.map(z=>sum(g[z],r=>+(r[COL.cQty]||0))),backgroundColor:MEASURE_COLORS.total,borderRadius:4},
      {label:'Auctioned',data:zones.map(z=>sum(g[z],r=>+(r[COL.auctQtyTotal]||0))),backgroundColor:MEASURE_COLORS.auction,borderRadius:4},
      {label:'Lifted',data:zones.map(z=>sum(g[z],r=>+(r[COL.liftedQty]||0))),backgroundColor:MEASURE_COLORS.lifted,borderRadius:4}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'bottom',labels:{boxWidth:12,padding:12}},tooltip:{callbacks:{label:ctx=>`${ctx.dataset.label}: ${fmtNum(ctx.parsed.y)}`}}},
      scales:{x:{grid:{display:false}},y:{beginAtZero:true,grid:{color:'#F3F4F6'}}}}});
}
function renderRegionCmp(data){
  destroyChart('chartRegionCmp');
  if(!data.length){ noData('chartRegionCmp'); return; }
  const g=groupBy(data,COL.region); let regions=Object.keys(g);
  const tot=x=>sum(g[x],r=>+(r[COL.cQty]||0));
  const zoneOf={}; regions.forEach(rg=>{ zoneOf[rg]=g[rg][0][COL.zone]||''; });
  if(REGION_SORT==='zone'){
    const zorder=(window.DASH_META&&DASH_META.zones)||[...new Set(data.map(r=>r[COL.zone]))];
    const zi=z=>{const i=zorder.indexOf(z); return i<0?999:i;};
    regions.sort((a,b)=> (zi(zoneOf[a])-zi(zoneOf[b])) || (tot(b)-tot(a)));
  } else regions.sort((a,b)=> REGION_SORT==='asc'? tot(a)-tot(b): tot(b)-tot(a));
  const labels=regions.map(rg=>`${zoneOf[rg]} · ${rg}`);
  CHARTS.chartRegionCmp=new Chart(document.getElementById('chartRegionCmp'),{type:'bar',
    data:{labels,datasets:[
      {label:'Verified base',data:regions.map(x=>sum(g[x],r=>+(r[COL.cQty]||0))),backgroundColor:MEASURE_COLORS.total,borderRadius:3},
      {label:'Auctioned',data:regions.map(x=>sum(g[x],r=>+(r[COL.auctQtyTotal]||0))),backgroundColor:MEASURE_COLORS.auction,borderRadius:3},
      {label:'Lifted',data:regions.map(x=>sum(g[x],r=>+(r[COL.liftedQty]||0))),backgroundColor:MEASURE_COLORS.lifted,borderRadius:3}]},
    options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',
      plugins:{legend:{position:'bottom',labels:{boxWidth:12,padding:12}},tooltip:{callbacks:{label:ctx=>`${ctx.dataset.label}: ${fmtNum(ctx.parsed.x)}`}}},
      scales:{x:{beginAtZero:true,grid:{color:'#F3F4F6'}},y:{grid:{display:false},ticks:{font:{size:10},autoSkip:false}}}}});
}
function renderAging(data){
  destroyChart('chartAging');
  const pqOf=r=>Math.max(0,(+(r[COL.auctQtyTotal]||0))-(+(r[COL.liftedQty]||0)));
  const rows=data.filter(r=>(r.__st==='Lifting Pending'||r.__st==='Partially Lifted') && pqOf(r)>0);
  const buckets=[{k:'0–30 days',lo:0,hi:30},{k:'31–60 days',lo:31,hi:60},{k:'61–90 days',lo:61,hi:90},{k:'90+ days',lo:91,hi:1e12}];
  buckets.forEach(b=>{b.qty=0;b.val=0;b.n=0;});
  let missing=0, missQty=0, future=0; const pend=[];
  rows.forEach(r=>{
    const pq=Math.max(0,(+(r[COL.auctQtyTotal]||0))-(+(r[COL.liftedQty]||0)));
    const val=+(r[COL.auctValueTotal]||0);
    const d=parseDateJS(r[COL.auctionDate]);
    if(!d){ missing++; missQty+=pq; return; }
    const days=daysSince(d);
    if(days<0){ future++; return; }
    const b=buckets.find(b=>days>=b.lo&&days<=b.hi); if(!b) return;
    b.qty+=pq; b.val+=val; b.n++;
    pend.push({zone:r[COL.zone],region:r[COL.region],cat:r[COL.assetCat],date:d,days,pq,val});
  });
  const bcol=['#FCD34D','#F59E0B','#D97706','#B45309'];
  if(rows.length && buckets.some(b=>b.qty>0)){
    CHARTS.chartAging=new Chart(document.getElementById('chartAging'),{type:'bar',
      data:{labels:buckets.map(b=>b.k),datasets:[{label:'Pending qty',data:buckets.map(b=>b.qty),backgroundColor:buckets.map((b,i)=>bcol[i]),borderRadius:5,maxBarThickness:70}]},
      options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>{const b=buckets[ctx.dataIndex]; return [`Pending qty: ${fmtNum(b.qty)}`,`Auction value: Rs ${fmtNum(b.val)}`,`${fmtNum(b.n)} lots`];}}}},scales:{x:{grid:{display:false}},y:{beginAtZero:true,grid:{color:'#F3F4F6'}}}}});
  } else noData('chartAging');
  const tw=document.getElementById('agingTableWrap');
  if(tw) tw.innerHTML=`<table class="data-table"><thead><tr><th>Age bucket</th><th style="text-align:right">Lots</th><th style="text-align:right">Pending Qty</th><th style="text-align:right">Auction Value (Rs.)</th></tr></thead><tbody>`+
    buckets.map(b=>`<tr><td>${b.k}</td><td style="text-align:right">${fmtNum(b.n)}</td><td style="text-align:right">${fmtNum(b.qty)}</td><td style="text-align:right">${fmtNum(b.val)}</td></tr>`).join('')+`</tbody></table>`;
  const sub=document.getElementById('agingSub');
  if(sub) sub.innerHTML=`Days since auction date for auctioned-but-not-fully-lifted lots, as of ${new Date().toISOString().slice(0,10)}. `+
    `<b>${fmtNum(missing)}</b> lot(s) (${fmtNum(missQty)} pending qty) have a missing / unparseable auction date`+(future?`; <b>${fmtNum(future)}</b> lot(s) carry a future-dated auction and are excluded from aging`:'')+`.`;
  const top=pend.sort((a,b)=>b.days-a.days).slice(0,10);
  const topw=document.getElementById('agingTopWrap');
  if(topw) topw.innerHTML= top.length===0 ? '<div class="empty-state">No auctioned lots are pending lifting under the current filters. ✓</div>' :
    `<table class="data-table"><thead><tr><th>Zone</th><th>Region</th><th>Asset Category</th><th>Auction Date</th><th style="text-align:right">Days Pending</th><th style="text-align:right">Pending Qty</th><th style="text-align:right">Auction Value (Rs.)</th></tr></thead><tbody>`+
    top.map(t=>`<tr><td>${escapeHtml(t.zone)}</td><td>${escapeHtml(t.region)}</td><td>${escapeHtml(t.cat)}</td><td>${t.date.toISOString().slice(0,10)}</td><td style="text-align:right;font-weight:700">${fmtNum(t.days)}</td><td style="text-align:right">${fmtNum(t.pq)}</td><td style="text-align:right">${fmtNum(t.val)}</td></tr>`).join('')+`</tbody></table>`;
}
function renderDQNote(data){
  const el=document.getElementById('dqNote'); if(!el) return;
  const catNorm=(window.__dqCat!==undefined?window.__dqCat:((window.DASH_META&&DASH_META.dq&&DASH_META.dq.categoriesNormalized)||0));
  let contra=0, dmiss=0, auctN=0;
  data.forEach(r=>{
    if(_s(r[COL.status])==='auctioned' && _s(r[COL.liftingStatus])==='not auctioned') contra++;
    if(isAuctionedStatus(deriveStatus(r))){ auctN++; if(!parseDateJS(r[COL.auctionDate])) dmiss++; }
  });
  el.innerHTML=`<b>Data-quality footnotes</b> &mdash; `+
    `Category labels normalized: <b>${fmtNum(catNorm)}</b> (Spices → OB - Spices). `+
    `Status contradictions resolved (Lifting-Status column said &ldquo;Not Auctioned&rdquo; on an auctioned lot; the derived status governs): <b>${fmtNum(contra)}</b>. `+
    `Auction dates missing / unparseable on auctioned lots: <b>${fmtNum(dmiss)}</b> of ${fmtNum(auctN)}. `+
    `All figures are shown exactly as recorded; anomalies are not clamped.`;
}
function wireT1Toggles(){
  const wire=(id,setter,rerender)=>{ const box=document.getElementById(id); if(!box||box.__wired) return; box.__wired=true;
    box.querySelectorAll('.sort-btn').forEach(b=>b.addEventListener('click',()=>{ setter(b.dataset.dir);
      box.querySelectorAll('.sort-btn').forEach(x=>x.classList.toggle('active',x===b)); rerender(T1_DATA); })); };
  wire('zoneSort', v=>ZONE_SORT=v, renderZoneCmp);
  wire('regionSort', v=>REGION_SORT=v, renderRegionCmp);
}

/* ==================== TAB 7 — DATABASE (sortable / searchable register) ==================== */
let DB_SORT={key:'zone',dir:'asc'};
const DB_COLS=[
  {key:'zone',label:'Zone',get:r=>r[COL.zone]},
  {key:'region',label:'Region',get:r=>r[COL.region]},
  {key:'class',label:'Assets Class',get:r=>r[COL.assetClass]},
  {key:'cat',label:'Assets Category',get:r=>r[COL.assetCat]},
  {key:'asset',label:'Asset Name',get:r=>r[COL.assetName]},
  {key:'cQty',label:'Verified Qty',get:r=>+(r[COL.cQty]||0),num:true},
  {key:'status',label:'Status',get:r=>deriveStatus(r)},
  {key:'adate',label:'Auction Date',get:r=>(r[COL.auctionDate]||'').toString().slice(0,10)},
  {key:'aQty',label:'Auction Qty',get:r=>+(r[COL.auctQtyTotal]||0),num:true},
  {key:'lQty',label:'Lifted Qty',get:r=>+(r[COL.liftedQty]||0),num:true},
  {key:'bQty',label:'Balance Qty',get:r=>+(r[COL.balanceQty]||0),num:true},
  {key:'val',label:'Auction Value (Rs.)',get:r=>+(r[COL.auctValueTotal]||0),num:true},
];
function dbSort(key){
  const col=DB_COLS.find(c=>c.key===key)||DB_COLS[0];
  if(DB_SORT.key===key) DB_SORT.dir = DB_SORT.dir==='asc'?'desc':'asc';
  else DB_SORT={key, dir: col.num?'desc':'asc'};
  renderTab7(getFiltered());
}
function fillDbFilters(){
  const rows=(typeof RAW!=='undefined'&&RAW)||[];
  const fill=(id,vals)=>{ const el=document.getElementById(id); if(!el||el.options.length>1) return;
    [...new Set(vals.filter(x=>x!==''&&x!=null))].sort((a,b)=>String(a).localeCompare(String(b)))
      .forEach(v=>{ const o=document.createElement('option'); o.value=v; o.textContent=v; el.appendChild(o); }); };
  fill('dbFZone',rows.map(r=>r[COL.zone])); fill('dbFRegion',rows.map(r=>r[COL.region]));
  fill('dbFClass',rows.map(r=>r[COL.assetClass])); fill('dbFCat',rows.map(r=>r[COL.assetCat]));
  fill('dbFStatus',rows.map(r=>deriveStatus(r)));
}
function renderTab7(data){
  const wrap=document.getElementById('dbTableWrap'); if(!wrap) return;
  fillDbFilters();
  const val=id=>{const e=document.getElementById(id); return e?e.value:'';};
  const q=(val('dbSearch')||'').trim().toLowerCase();
  const fz=val('dbFZone'),fr=val('dbFRegion'),fc=val('dbFClass'),fk=val('dbFCat'),fs=val('dbFStatus');
  let rows=data.filter(r=>{
    if(fz&&r[COL.zone]!==fz) return false;
    if(fr&&r[COL.region]!==fr) return false;
    if(fc&&r[COL.assetClass]!==fc) return false;
    if(fk&&r[COL.assetCat]!==fk) return false;
    const st=deriveStatus(r); if(fs&&st!==fs) return false;
    if(q){ const hay=[r[COL.zone],r[COL.region],r[COL.assetClass],r[COL.assetCat],r[COL.assetName],st,r[COL.bidder]].join(' ').toLowerCase(); if(!hay.includes(q)) return false; }
    return true;
  });
  const col=DB_COLS.find(c=>c.key===DB_SORT.key)||DB_COLS[0];
  rows.sort((a,b)=>{ let x=col.get(a),y=col.get(b);
    if(col.num){ x=+x||0; y=+y||0; return DB_SORT.dir==='asc'?x-y:y-x; }
    x=String(x).toLowerCase(); y=String(y).toLowerCase(); return DB_SORT.dir==='asc'?x.localeCompare(y):y.localeCompare(x); });
  const arw=c=> DB_SORT.key===c.key ? (DB_SORT.dir==='asc'?'▲':'▼') : '↕';
  const head='<tr>'+DB_COLS.map(c=>`<th class="${DB_SORT.key===c.key?'sorted':''}${c.num?' num':''}" onclick="dbSort('${c.key}')">${c.label}<span class="arw">${arw(c)}</span></th>`).join('')+'</tr>';
  const shown=rows.slice(0,1000);
  const body=shown.map(r=>{ const st=deriveStatus(r);
    return '<tr>'+
      `<td>${escapeHtml(r[COL.zone])}</td><td>${escapeHtml(r[COL.region])}</td>`+
      `<td>${escapeHtml(r[COL.assetClass])}</td><td>${escapeHtml(r[COL.assetCat])}</td>`+
      `<td>${escapeHtml(r[COL.assetName])}</td>`+
      `<td class="num">${fmtNum(+(r[COL.cQty]||0))}</td>`+
      `<td><span class="status-pill" style="background:${STATUS_COLORS[st]||'#6B7280'}">${st}</span></td>`+
      `<td>${escapeHtml((r[COL.auctionDate]||'').toString().slice(0,10))||'—'}</td>`+
      `<td class="num">${fmtNum(+(r[COL.auctQtyTotal]||0))}</td>`+
      `<td class="num">${fmtNum(+(r[COL.liftedQty]||0))}</td>`+
      `<td class="num">${fmtNum(+(r[COL.balanceQty]||0))}</td>`+
      `<td class="num">${fmtNum(+(r[COL.auctValueTotal]||0))}</td></tr>`;
  }).join('');
  wrap.innerHTML= rows.length===0 ? '<div class="empty-state">No records match the current search / filters.</div>'
    : `<table class="db-table"><thead>${head}</thead><tbody>${body}</tbody></table>`;
  const cnt=document.getElementById('dbCount');
  if(cnt) cnt.textContent=`Showing ${fmtNum(shown.length)} of ${fmtNum(rows.length)} filtered rows (${fmtNum(data.length)} in scope after global slicers).`;
}

/* ---------- Initial empty state ---------- */"""
sub(r"/\* ---------- Initial empty state ---------- \*/", lambda m: NEW_JS, label="inject tab0/tab4 js")

# 10b) (Zone Scorecard + News panel removed — the Overall tab is now the 7-section
#      narrative built in step 4b / renderTab1; the old chartByCat anchor no longer exists.)

# 11) bootstrap from embedded JSON (replace the bootEmpty() call)
BOOT = r"""(function(){
  const meta = JSON.parse(document.getElementById('dashboardData').textContent);
  window.DASH_META = meta.meta; window.DASH_UPDATES = meta.updates;
  const cols = meta.cols;
  RAW = meta.rows.map(a => { const o={}; cols.forEach((c,i)=>{ o[COL[c]] = a[i]; });
    // default any numeric COL not present in the baked payload to 0 (avoids NaN in sums)
    NUMERIC_COLS.forEach(k=>{ if(typeof o[k] !== 'number') o[k] = Number(o[k]) || 0; });
    o[COL.itemCat] = o[COL.itemCat] || o[COL.assetClass] || '';   // 'Asset Class' slicer
    return o; });
  window.__dqCat = (meta.meta.dq && meta.meta.dq.categoriesNormalized) || 0;
  hydrateSlicers(); renderAll(); renderUpdates();
  const d = meta.meta.refreshDate || '';
  const sp = document.getElementById('sourcePath');
  if(sp){ sp.textContent = 'Source: ' + (meta.meta.sourcePath || ''); sp.title = meta.meta.sourcePath || ''; }
  setStatus(`Live · ${meta.meta.rowCount.toLocaleString()} records · last synced ${d}`, true);
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
