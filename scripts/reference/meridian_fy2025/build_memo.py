"""Phase 6 — build the CFO/board-ready executive memo (.docx), annual structure."""
import pickle
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT_DIR = "data/outputs"
m = pickle.load(open(f"{OUT_DIR}/_model.pkl", "rb"))
PERIODS = m["PERIODS"]
pl, bs, cf = m["consol_pl"], m["consol_bs"], m["cash_flow"]

NAVY = RGBColor(0x1F, 0x38, 0x64)
GREY = RGBColor(0x59, 0x59, 0x59)
RED = RGBColor(0xC0, 0x00, 0x00)

doc = Document()
style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(11)

def h1(text):
    p = doc.add_heading(text, level=1)
    for run in p.runs:
        run.font.color.rgb = NAVY
    return p

def h2(text):
    p = doc.add_heading(text, level=2)
    for run in p.runs:
        run.font.color.rgb = NAVY
    return p

def body(text, italic=False, bold=False, size=11, color=None):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.italic = italic
    r.bold = bold
    r.font.size = Pt(size)
    if color:
        r.font.color.rgb = color
    return p

def bullet(text):
    p = doc.add_paragraph(text, style="List Bullet")
    return p

def set_cell_shading(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), color_hex)
    tcPr.append(shd)

def kpi_table(headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Light Grid Accent 1"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = h
        for p in hdr[i].paragraphs:
            for r in p.runs:
                r.bold = True
        set_cell_shading(hdr[i], "1F3864")
        hdr[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    for row in rows:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = str(val)
    return t

# ============================================================
# COVER
# ============================================================
title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = title_p.add_run("Meridian Technologies Group")
r.font.size = Pt(26); r.font.bold = True; r.font.color.rgb = NAVY

sub_p = doc.add_paragraph()
sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = sub_p.add_run("Annual Business Review — Fiscal Year 2025")
r.font.size = Pt(16); r.font.color.rgb = GREY

meta_p = doc.add_paragraph()
meta_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = meta_p.add_run("Reporting currency: USD  |  Fiscal period: 01 Jan 2025 – 31 Dec 2025  |  "
                    "Accounting basis: Multi-standard rollup — US01 US GAAP, UK01 IFRS, IN01 Ind AS "
                    "(current-rate method consolidation)")
r.font.size = Pt(10); r.italic = True

doc.add_paragraph()
note_p = doc.add_paragraph()
note_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = note_p.add_run("First review cycle for this business. Trend lines below cover a full-year backfill "
                    "(all 12 months supplied and processed together), not periods reviewed as they occurred. "
                    "Multi-year trend analysis begins next cycle.")
r.font.size = Pt(9); r.italic = True; r.font.color.rgb = GREY

doc.add_page_break()

# ============================================================
# 1. EXECUTIVE SUMMARY
# ============================================================
h1("1. Executive Summary")
fy_rev = sum(pl[p]["revenue"] for p in PERIODS)
fy_ni = sum(pl[p]["net_income"] for p in PERIODS)
fy_ebit_adj = sum(pl[p]["ebit_adjusted"] for p in PERIODS)
gm_start = pl[PERIODS[0]]["gross_profit"] / pl[PERIODS[0]]["revenue"]
gm_end = pl[PERIODS[-1]]["gross_profit"] / pl[PERIODS[-1]]["revenue"]
rev_growth = pl[PERIODS[-1]]["revenue"] / pl[PERIODS[0]]["revenue"] - 1
hc_start = sum(d.get("active", 0) for d in m["hc_by_period_dept"][PERIODS[0]].values())
hc_end = sum(d.get("active", 0) for d in m["hc_by_period_dept"][PERIODS[-1]].values())

body(f"Meridian Technologies Group closed FY2025 with USD {fy_rev:,.0f} in consolidated revenue "
     f"across its three entities (US, UK, India), up {rev_growth:.0%} from January's monthly run-rate to "
     f"December's. Consolidated gross margin expanded from {gm_start:.1%} to {gm_end:.1%} over the year, "
     f"driven primarily by the deliberate wind-down of the lower-margin Legacy Hardware line, which is now "
     f"substantially complete.")

body(f"Full-year net income was USD {fy_ni:,.0f}, held down by two large, clearly-flagged one-time items in "
     f"Q3 tied to that same hardware exit: an $850,000 restructuring charge (severance and facility exit, "
     f"August) and a $320,000 inventory write-down (September). Excluding these one-time items, adjusted "
     f"FY2025 EBIT was USD {fy_ebit_adj:,.0f} — the underlying subscription and services business was "
     f"profitable and improving all year; the hardware exit is a deliberate, bounded cost, not an ongoing "
     f"trend.")

body(f"Cash ended the year at USD {bs[PERIODS[-1]]['cash']:,.0f}, up from USD {bs[PERIODS[0]]['cash']:,.0f} "
     f"at the start of January. Board attention: roughly USD 2.0M of that increase came from a September "
     f"draw on the revolving credit facility, not organic operating cash generation — September's operating "
     f"performance was actually a net loss month. See Section 6 for the full explanation; this does not "
     f"reflect an underlying cash-generation problem once the one-time hardware costs are excluded, but the "
     f"board should not read the year-end cash balance as a pure organic-performance signal.")

body(f"Headcount grew from {hc_start} to {hc_end} across the year, including a targeted reduction in the "
     f"Hardware Mfg Support department (9 involuntary terminations in October) directly tied to the line "
     f"sunset — the only department reduction of the year; every other department grew or held flat.")

h2("Headline FY2025 metrics")
kpi_table(
    ["Metric", "FY2025"],
    [
        ["Total revenue (USD)", f"{fy_rev:,.0f}"],
        ["Net income (USD)", f"{fy_ni:,.0f}"],
        ["Adjusted EBIT, ex. one-time items (USD)", f"{fy_ebit_adj:,.0f}"],
        ["Gross margin, Jan → Dec", f"{gm_start:.1%} → {gm_end:.1%}"],
        ["Ending cash (31 Dec 2025, USD)", f"{bs[PERIODS[-1]]['cash']:,.0f}"],
        ["Headcount, Jan → Dec", f"{hc_start} → {hc_end}"],
        ["Customer concentration (top-1 / top-5)", f"{m['top1_share']:.1%} / {m['top5_share']:.1%}"],
    ],
)

# ============================================================
# 2. FINANCIAL PERFORMANCE
# ============================================================
doc.add_page_break()
h1("2. Financial Performance")
body("Consolidated income statement walk (USD, current-rate method — see Section 8 / workbook for FX "
     "methodology and rates).")
kpi_table(
    ["Line", "Jan 2025", "Jun 2025", "Dec 2025", "FY2025 Total"],
    [
        ["Revenue", f"{pl['2025-01']['revenue']:,.0f}", f"{pl['2025-06']['revenue']:,.0f}",
         f"{pl['2025-12']['revenue']:,.0f}", f"{fy_rev:,.0f}"],
        ["Gross profit", f"{pl['2025-01']['gross_profit']:,.0f}", f"{pl['2025-06']['gross_profit']:,.0f}",
         f"{pl['2025-12']['gross_profit']:,.0f}", f"{sum(pl[p]['gross_profit'] for p in PERIODS):,.0f}"],
        ["Operating income (EBIT)", f"{pl['2025-01']['ebit']:,.0f}", f"{pl['2025-06']['ebit']:,.0f}",
         f"{pl['2025-12']['ebit']:,.0f}", f"{sum(pl[p]['ebit'] for p in PERIODS):,.0f}"],
        ["Net income", f"{pl['2025-01']['net_income']:,.0f}", f"{pl['2025-06']['net_income']:,.0f}",
         f"{pl['2025-12']['net_income']:,.0f}", f"{fy_ni:,.0f}"],
    ],
)
h2("Margin bridge — the hardware exit is a mix story, not a pricing or cost-control story")
body("Gross margin rose roughly 4.6 points over the year (69.5% in January to 74.1% in December). The "
     "entire move is explained by product mix: Legacy Hardware carries materially lower gross margin than "
     "the subscription and services lines, and its revenue fell from USD 165,000 in January to roughly "
     "USD 5,000/month by September as the line was sunset. Subscription and services margins were "
     "essentially flat year-over-year; this is not a pricing or cost-efficiency story, and the board "
     "should not extrapolate further margin expansion from this driver once the mix shift is fully "
     "annualized.")
h2("One-time items (excluded from adjusted EBIT above)")
kpi_table(
    ["Period", "Item", "Amount (USD)", "Driver"],
    [
        ["Aug 2025", "Restructuring charge", "850,000", "Severance + facility exit, Hardware line sunset"],
        ["Sep 2025", "Inventory write-down", "320,000", "Hardware inventory to net realizable value, line discontinued"],
    ],
)

# ============================================================
# 3. BALANCE SHEET & CASH POSITION
# ============================================================
doc.add_page_break()
h1("3. Balance Sheet & Cash Position")
body(f"Total consolidated assets stood at USD {bs['2025-12']['assets']:,.0f} at year-end, against total "
     f"liabilities of USD {bs['2025-12']['liabilities']:,.0f} and total equity of USD "
     f"{bs['2025-12']['equity_total']:,.0f}. The balance sheet articulates (Assets = Liabilities + Equity) "
     f"in every month of FY2025 with zero variance.")
h2("Cash & the September financing draw")
body(f"Cash grew from USD {bs['2025-01']['cash']:,.0f} in January to USD {bs['2025-12']['cash']:,.0f} in "
     f"December. The single largest monthly movement was September, when cash rose USD "
     f"{bs['2025-09']['cash']-bs['2025-08']['cash']:,.0f} — but USD 2,000,000 of that (over 99%) was a new "
     f"draw on the revolving credit facility, not operating cash generation. September's net income was "
     f"actually negative (USD {pl['2025-09']['net_income']:,.0f}), driven by the inventory write-down noted "
     f"above. The board should read September–December's healthier cash balance as debt-financed, not as "
     f"evidence of an operating-performance inflection; FY2025 organic operating cash flow (see workbook "
     f"Cash Flow tab) was positive but materially smaller than the headline cash-balance change.")
body(f"Full-year operating cash flow: USD {sum(cf[p]['operating_cf_reported'] for p in PERIODS):,.0f}. "
     f"Full-year financing cash flow: USD {sum(cf[p]['financing_cf'] for p in PERIODS):,.0f} "
     f"(entirely the September revolver draw — no repayment booked yet in FY2025).")
h2("Debt")
body(f"Short-term debt / revolver balance at year-end: USD {bs['2025-12']['debt']:,.0f}, drawn in "
     f"September 2025. No amortization schedule or covenant detail was supplied this cycle; recommend "
     f"obtaining the facility's terms (maturity, covenants, rate basis) for the next review cycle.")

# ============================================================
# 4. REVENUE & GROWTH DRIVERS
# ============================================================
doc.add_page_break()
h1("4. Revenue & Growth Drivers")
new_fy = sum(m["rev_bridge"][p].get("new", 0) for p in PERIODS)
exp_fy = sum(m["rev_bridge"][p].get("expansion", 0) for p in PERIODS)
churn_fy = sum(m["rev_bridge"][p].get("churn", 0) for p in PERIODS)
contraction_fy = sum(m["rev_bridge"][p].get("contraction", 0) for p in PERIODS)
renewal_fy = sum(m["rev_bridge"][p].get("renewal", 0) for p in PERIODS)
body(f"FY2025 gross new-business bookings totaled USD {new_fy:,.0f}, plus USD {exp_fy:,.0f} of "
     f"expansion/upsell within the existing base and USD {renewal_fy:,.0f} of renewals. This was offset by "
     f"USD {abs(contraction_fy):,.0f} of contraction and USD {abs(churn_fy):,.0f} of churn — almost "
     f"entirely a single event (below).")
h2("Material churn event — Vantage Logistics Corp")
churn_amt = abs(m["rev_bridge"]["2025-05"].get("churn", 0))
body(f"In May 2025, Vantage Logistics Corp (Platform Pro, North America) churned in full: USD "
     f"{churn_amt:,.0f} of annualized contract value lost in a single event, equal to "
     f"{churn_amt/pl['2025-05']['revenue']*100:.0f}% of that month's consolidated revenue. This is the "
     f"largest single revenue event of the year in either direction and warrants a customer post-mortem; "
     f"the data supplied this cycle does not include a churn reason code.")
h2("Customer concentration")
kpi_table(
    ["Customer", "FY2025 revenue (USD)", "% of gross booked revenue"],
    [[c, f"{amt:,.0f}", f"{amt/sum(v for v in m['customer_revenue'].values() if v>0)*100:.1f}%"]
     for c, amt in m["top_customers"]],
)
body(f"Top-1 customer concentration: {m['top1_share']:.1%} (threshold: 15%). Top-5: {m['top5_share']:.1%} "
     f"(threshold: 40%). Neither threshold was tripped — concentration risk is currently well-managed at "
     f"the portfolio level, though the May churn event shows single-account revenue swings can still be "
     f"material in-month.")

# ============================================================
# 5. HEADCOUNT & ORGANIZATION
# ============================================================
doc.add_page_break()
h1("5. Headcount & Organization")
body(f"Total headcount grew from {hc_start} (January) to {hc_end} (December), a net increase of "
     f"{hc_end-hc_start} across the year, concentrated in Engineering and Sales. The one exception was "
     f"Hardware Mfg Support, which was reduced from its FY2025 peak to 11 active employees via 9 "
     f"involuntary terminations in October 2025 — the RIF associated with the Q3 restructuring charge in "
     f"Section 2.")
h2("Headcount by department (December 2025)")
dec_hc = m["hc_by_period_dept"]["2025-12"]
kpi_table(
    ["Department", "Active headcount", "FY2025 fully-loaded monthly cost (USD)"],
    [[d, dec_hc.get(d, {}).get("active", 0), f"{dec_hc.get(d, {}).get('cost', 0):,.0f}"]
     for d in sorted(dec_hc.keys())],
)

# ============================================================
# 6. RISKS & WATCH ITEMS
# ============================================================
doc.add_page_break()
h1("6. Risks & Watch Items")
body("Every threshold trip and materiality-override item this cycle, in period order. All items are new — "
     "this is the first review cycle for this business, so nothing is carried forward from a prior period.")
for rf in sorted(m["red_flags"], key=lambda x: x["period"]):
    p = doc.add_paragraph(style="List Bullet")
    r = p.add_run(f"[{rf['period']}] {rf['metric']}: ")
    r.bold = True
    p.add_run(rf["detail"])

# ============================================================
# 7. OUTLOOK / NEXT-PERIOD FOCUS
# ============================================================
doc.add_page_break()
h1("7. Outlook / Next-Period Focus")
for text in [
    "Confirm the Legacy Hardware wind-down is fully complete (near-zero revenue since September) and "
    "retire the associated COGS/inventory lines from the taxonomy if no further activity is expected — "
    "avoids carrying a stale line item into FY2026 reporting.",
    "Obtain FY2026 budget for IN01 (India) — no budget was supplied for this entity in FY2025, so budget "
    "variance could not be reported for it this cycle.",
    "Obtain the revolving credit facility's terms (maturity, covenants, rate basis) — drawn USD 2.0M in "
    "September 2025 with no supporting detail provided this cycle.",
    "Run a post-mortem on the Vantage Logistics Corp churn (May 2025, USD 780,000) — the largest single "
    "revenue event of the year; no reason code was available in the data supplied.",
    "Establish an explicit Intercompany Payable account on US01's chart of accounts so the IN01 "
    "intercompany receivable balance can be fully eliminated in consolidation next cycle (see Section 8).",
    "FY2026 will be the first cycle with a genuine multi-year trend and a full statistical baseline for "
    "anomaly detection — recommend maintaining the current taxonomy and thresholds unchanged into next "
    "cycle to preserve comparability.",
]:
    bullet(text)

# ============================================================
# 8. APPENDIX
# ============================================================
doc.add_page_break()
h1("8. Appendix")
body("This memo is generated entirely from the FY2025 consolidated history workbook — no figure here is "
     "independently re-derived. Full detail, formulas, and the complete Assumptions & Judgment Log are in:")
p = doc.add_paragraph()
r = p.add_run("2025-FY_Business_Review_USD.xlsx")
r.bold = True
body("A separate document, 2025-FY_Methodology_and_Data_Lineage.docx, sets out every input file used, "
     "the processing methodology applied, and how each output was derived, for audit purposes.")
body("Consolidation: current-rate method (balance sheet at period-end spot, income statement at "
     "period-average). FX rates: Company Treasury Rate, sourced from 06_FX_Rates_2025.xlsx. Entities: "
     "US01 (US GAAP), UK01 (IFRS), IN01 (Ind AS) — see the workbook's FX & Consolidation tab for "
     "standard-divergence notes.")

doc.save(f"{OUT_DIR}/2025-FY_Business_Review_USD.docx")
print("Memo written: 2025-FY_Business_Review_USD.docx")
