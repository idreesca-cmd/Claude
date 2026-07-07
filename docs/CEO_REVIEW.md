# USC Asset Disposal Dashboard — CEO Review & Coverage Assessment

*Role-play: reviewed as if I were the CEO of Utility Stores Corporation receiving this
Third-Party Validation dashboard from Baker Tilly, then answered as the BT engagement lead.*

The dashboard now has **7 tabs**: Category Milestone Overview · Overall Project Status ·
Quantity Reconciliation · Auction Reconciliation · Payment & Lifting · **Payments & DD Detail** ·
**Anomalies**.

---

## A. What a CEO expects — and how this dashboard answers it

| # | As USC CEO, I expect… | How the dashboard delivers | Verdict |
|---|------------------------|----------------------------|---------|
| 1 | **A bird's-eye view of the whole programme** in 10 seconds | *Category Milestone Overview* (tab 1): every asset category × the 4 stages (Physical Verification → Auction → Lifting → Payment) as % rings + planned-vs-actual bars. *Overall Project Status* adds headline KPIs (9 zones, 70 regions, 11,782 records, 2.95 M Odoo qty). | ✅ Addressed |
| 2 | **Zone- and Region-wise** performance | A **Zone Scorecard** on the Overall tab (every zone × the 4 milestone %s + payment, ranked, with an All-Zones total) gives the one-look view; global **Zone → Region slicers** cascade across *every* tab; plus per-Zone/Region charts. Pick a zone → the whole dashboard (incl. milestone rings) recomputes. | ✅ Addressed (Zone Scorecard + slicers + per-zone/region charts). |
| 3 | **Money recovered vs. expected** | *Payment & Lifting*: Payment received **Rs 968.3 M** vs Total Auction Value **Rs 17.12 bn** booked (5.7%) vs Reserve Rs 5.03 M; Lifted 930,063 units, Balance 845,605. | ✅ Addressed (see also DQ note on the Rs 17.1 bn figure). |
| 4 | **Who bought what, and did they pay?** — bidder name, category won, DD no., DD date, DD amount | **Payments & DD Detail (tab 6)**: per-bidder summary — Winning Bidder · Zone(s) · Categories Won · DD No(s). · Payment Date(s) · **Total Payment (DD amount)** · line count. 24 unique bidders; top bidders & payment-by-category charts. | ✅ Addressed (dedicated tab). ⚠️ DD-number granularity limited by source — see G2. |
| 5 | **Are the numbers trustworthy? Show me the exceptions** | **Anomalies (tab 7)** — a dedicated section: 1,600 flagged records across *Lifted > Auctioned* (983), *Auctioned > Counted* (227), *Payment > Auction Value* (182), *Negative Balance*; by-type chart + line-level detail with the conflicting values. Anomalies are also flagged inline (▲) on the milestone rings and explained in a Data-Quality note. | ✅ Addressed (dedicated tab **+** inline flags). |
| 6 | **A shareable, professional, always-current view** | Self-contained HTML (Baker Tilly palette/fonts, zero external dependencies), regenerated from the Drive workbook; hosts on GitHub Pages as a no-login link; a corrected copy of the workbook (#REF! fixed) is a separate deliverable. | ✅ Addressed |

---

## B. CEO's pointed questions, answered

**Q: "Is it a true bird's-eye view?"**
Yes — tab 1 is the one-screen executive summary. Each of the 9 categories shows exactly how far
it has moved through Verification → Auction → Lifting → Payment, so I can see at a glance that, e.g.,
Physical Verification is well advanced (ERP 88%, Rice 78%) but **Lifting and Payment are the
bottlenecks** (many categories 0–40%, several with no monetised auction yet).

**Q: "Zone and Region wise — can I drill in?"**
Yes. Set Zone = *Multan* (or any zone) and the entire dashboard — KPIs, milestone rings, charts,
tables — recomputes for that zone; the Region slicer then narrows further. Every quantity/auction/
payment chart also has a by-Zone and by-Region breakdown.

**Q: "What about the anomalies you found — is there a separate section?"**
Yes — a dedicated **Anomalies** tab. It doesn't just flag them, it **quantifies and lists** them:
1,600 records where the source contradicts itself (sold more than counted, lifted more than sold,
paid more than booked, negative balances). This is the list my team should reconcile at source.

**Q: "Payments from bidders — the full DD trail?"**
The **Payments & DD Detail** tab gives bidder name, the categories they won, DD number(s), payment
date(s) and total DD amount per bidder (e.g. *Sain Abdul Hakeem* Rs 474.5 M, *Abid Fareed*
Rs 212.4 M). Summary KPIs + top-bidder chart sit above the detail.

---

## C. BT's candour — gaps & recommended next steps

A good validator flags its own limits. These are **source-data** constraints, not dashboard bugs —
the dashboard surfaces them faithfully:

- **G1 — Zone scorecard — ✅ DONE.** Added a Zone Scorecard on the Overall tab: each zone × the 4
  milestone %s + payment (Rs), ranked, with an All-Zones total row and >100% anomaly flags — lagging
  zones (Quetta, Faisalabad) and anomalies (Peshawar auction 762%, Karachi payment 381%) jump out.
- **G2 — DD detail is thin in the source.** The "DD No." column is mostly blank or holds a payment
  *mode* ("Cheque (JS Bank)", "cdr") rather than a draft number, and payments are booked at
  asset-line level. So per-*individual-DD* amount/date can't be cleanly separated — we aggregate to
  per-bidder totals and flag it. *Recommend* USC capture DD No./DD date/DD amount as discrete fields.
- **G3 — Total Auction Value looks inflated** (Rs 17.12 bn vs Reserve Rs 5.03 M and Payment
  Rs 968 M). Likely unit/entry errors in a few high-volume lines (Furniture, Rice). The Anomalies
  tab surfaces the worst offenders; USC should reconcile before this figure is quoted externally.
- **G4 — Completion is shown per-category, not as one blended %.** Deliberate: blending qty across
  vehicles, spices and rice into a single % would mislead. If a headline number is wanted, we can add
  a weighted "value-at-stake" completion gauge.
- **G5 — Snapshot, not trend.** The dashboard reflects the latest workbook; it has no time series.
  Because it regenerates from Drive on demand, we can add period-over-period once dated snapshots exist.

**Bottom line (CEO hat):** the dashboard gives me the executive view, the zone/region drill-down, a
real payments-by-bidder trail, and — importantly for a validation engagement — an honest, dedicated
anomalies register. The open items above are about **USC's source data quality**, which is exactly
what a Third-Party Validation is meant to expose.
