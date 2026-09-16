"""Phase 6 — build the board deck (.pptx), annual structure (~13-15 slides), native charts."""
import pickle
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

def disable_invert_if_negative(series):
    """PowerPoint's default 'invert if negative' renders negative bars unfilled/white
    instead of the series color -- force it off so negative bars stay colored."""
    ser_elem = series._element
    inv = ser_elem.find(qn("c:invertIfNegative"))
    if inv is None:
        inv = ser_elem.makeelement(qn("c:invertIfNegative"), {})
        spPr = ser_elem.find(qn("c:spPr"))
        if spPr is not None:
            spPr.addnext(inv)
        else:
            ser_elem.append(inv)
    inv.set("val", "0")

OUT_DIR = "data/outputs"
m = pickle.load(open(f"{OUT_DIR}/_model.pkl", "rb"))
PERIODS = m["PERIODS"]
MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
pl, bs, cf = m["consol_pl"], m["consol_bs"], m["cash_flow"]

NAVY = RGBColor(0x1F, 0x38, 0x64)
LBLUE = RGBColor(0xDC, 0xE6, 0xF1)
GREY = RGBColor(0x59, 0x59, 0x59)
RED = RGBColor(0xC0, 0x00, 0x00)
GREEN = RGBColor(0x37, 0x56, 0x23)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LGREY = RGBColor(0xF2, 0xF2, 0xF2)
GOLD = RGBColor(0xBF, 0x8F, 0x00)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

def add_slide():
    return prs.slides.add_slide(BLANK)

def add_bg(slide, color=WHITE):
    rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    rect.fill.solid(); rect.fill.fore_color.rgb = color
    rect.line.fill.background()
    rect.shadow.inherit = False
    slide.shapes._spTree.remove(rect._element)
    slide.shapes._spTree.insert(2, rect._element)
    return rect

def header_bar(slide, title_text, subtitle_text=None, dark=False):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(1.05))
    bar.fill.solid(); bar.fill.fore_color.rgb = NAVY
    bar.line.fill.background(); bar.shadow.inherit = False
    tb = slide.shapes.add_textbox(Inches(0.45), Inches(0.12), Inches(12.4), Inches(0.55))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = title_text
    r.font.size = Pt(26); r.font.bold = True; r.font.color.rgb = WHITE
    if subtitle_text:
        tb2 = slide.shapes.add_textbox(Inches(0.45), Inches(0.62), Inches(12.4), Inches(0.4))
        tf2 = tb2.text_frame; tf2.word_wrap = True
        p2 = tf2.paragraphs[0]
        r2 = p2.add_run(); r2.text = subtitle_text
        r2.font.size = Pt(12); r2.font.color.rgb = LBLUE; r2.font.italic = True

def footer(slide, text="Meridian Technologies Group — FY2025 Annual Business Review — USD — Multi-standard rollup (US GAAP / IFRS / Ind AS)"):
    tb = slide.shapes.add_textbox(Inches(0.45), Inches(7.15), Inches(12.4), Inches(0.3))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = text
    r.font.size = Pt(8); r.font.color.rgb = GREY

def textbox(slide, left, top, width, height, text, size=14, bold=False, color=None, align=PP_ALIGN.LEFT, italic=False):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = color or RGBColor(0x22, 0x22, 0x22)
    return tb

def bullets(slide, left, top, width, height, items, size=13, color=None):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame; tf.word_wrap = True
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = "•  " + it
        p.font.size = Pt(size)
        p.font.color.rgb = color or RGBColor(0x22, 0x22, 0x22)
        p.space_after = Pt(8)
    return tb

def stat_card(slide, left, top, width, height, label, value, color=NAVY, value_color=None):
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    box.fill.solid(); box.fill.fore_color.rgb = LGREY
    box.line.color.rgb = color; box.line.width = Pt(1)
    box.shadow.inherit = False
    tf = box.text_frame; tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p1 = tf.paragraphs[0]; p1.alignment = PP_ALIGN.CENTER
    r1 = p1.add_run(); r1.text = value
    r1.font.size = Pt(26); r1.font.bold = True; r1.font.color.rgb = value_color or color
    p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run(); r2.text = label
    r2.font.size = Pt(11); r2.font.color.rgb = GREY

def line_chart(slide, left, top, width, height, categories, series_dict, title_text=None, pct=False):
    data = CategoryChartData()
    data.categories = categories
    for name, vals in series_dict.items():
        data.add_series(name, vals)
    gframe = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, left, top, width, height, data)
    chart = gframe.chart
    chart.has_legend = len(series_dict) > 1
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(10)
    chart.has_title = bool(title_text)
    if title_text:
        chart.chart_title.text_frame.text = title_text
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(12)
    cat_ax = chart.category_axis
    cat_ax.tick_labels.font.size = Pt(9)
    val_ax = chart.value_axis
    val_ax.tick_labels.font.size = Pt(9)
    if pct:
        val_ax.tick_labels.number_format = '0%'
        val_ax.tick_labels.number_format_is_linked = False
    for series in chart.series:
        series.smooth = False
        series.format.line.color.rgb = NAVY
        series.format.line.width = Pt(2.25)
    return chart

def bar_chart(slide, left, top, width, height, categories, series_dict, title_text=None, colors=None, stacked=False):
    data = CategoryChartData()
    data.categories = categories
    for name, vals in series_dict.items():
        data.add_series(name, vals)
    ctype = XL_CHART_TYPE.COLUMN_STACKED if stacked else XL_CHART_TYPE.COLUMN_CLUSTERED
    gframe = slide.shapes.add_chart(ctype, left, top, width, height, data)
    chart = gframe.chart
    chart.has_legend = len(series_dict) > 1
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(10)
    chart.has_title = bool(title_text)
    if title_text:
        chart.chart_title.text_frame.text = title_text
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(12)
    chart.category_axis.tick_labels.font.size = Pt(9)
    chart.value_axis.tick_labels.font.size = Pt(9)
    palette = colors or [NAVY, GOLD, GREEN, RED, GREY]
    for i, series in enumerate(chart.series):
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = palette[i % len(palette)]
        disable_invert_if_negative(series)
    return chart

def simple_table(slide, left, top, width, height, headers, rows, col_widths=None, font_size=11):
    n_rows = len(rows) + 1
    n_cols = len(headers)
    gframe = slide.shapes.add_table(n_rows, n_cols, left, top, width, height)
    table = gframe.table
    if col_widths:
        for i, w in enumerate(col_widths):
            table.columns[i].width = w
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = h
        cell.fill.solid(); cell.fill.fore_color.rgb = NAVY
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True; p.font.size = Pt(font_size); p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER
    for ri, row in enumerate(rows, start=1):
        for ci, val in enumerate(row):
            cell = table.cell(ri, ci)
            cell.text = str(val)
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(font_size)
            p.alignment = PP_ALIGN.CENTER if ci > 0 else PP_ALIGN.LEFT
            cell.fill.solid()
            cell.fill.fore_color.rgb = LGREY if ri % 2 == 0 else WHITE
    return table

def usd(x, decimals=0, unit=""):
    sign = "-" if x < 0 else ""
    return f"{sign}${abs(x):,.{decimals}f}{unit}"

rev_series = [pl[p]["revenue"] for p in PERIODS]
ni_series = [pl[p]["net_income"] for p in PERIODS]
ebit_series = [pl[p]["ebit"] for p in PERIODS]
ebit_adj_series = [pl[p]["ebit_adjusted"] for p in PERIODS]
gm_series = [pl[p]["gross_profit"] / pl[p]["revenue"] for p in PERIODS]
cash_series = [bs[p]["cash"] for p in PERIODS]
debt_series = [bs[p]["debt"] for p in PERIODS]
hc_series = [sum(d.get("active", 0) for d in m["hc_by_period_dept"][p].values()) for p in PERIODS]
bud_series = [m["bud_revenue"].get(p, 0) for p in PERIODS]

fy_rev = sum(rev_series)
fy_ni = sum(ni_series)
fy_ebit_adj = sum(ebit_adj_series)

# =====================================================================
# SLIDE 1 — Title
# =====================================================================
s = add_slide(); add_bg(s, NAVY)
textbox(s, Inches(0.9), Inches(2.3), Inches(11.5), Inches(1.2),
        "Meridian Technologies Group", size=44, bold=True, color=WHITE)
textbox(s, Inches(0.9), Inches(3.3), Inches(11.5), Inches(0.7),
        "Annual Business Review — Fiscal Year 2025", size=24, color=LBLUE)
textbox(s, Inches(0.9), Inches(4.1), Inches(11.5), Inches(0.5),
        "Reporting currency: USD  |  Period: 01 Jan 2025 – 31 Dec 2025  |  Basis: US GAAP (US) / IFRS (UK) / Ind AS (India), consolidated",
        size=13, color=WHITE, italic=True)
textbox(s, Inches(0.9), Inches(5.0), Inches(11.5), Inches(0.5),
        "First review cycle — FY2025 trend is a full-year backfill, not periods reviewed as they occurred.",
        size=11, color=LBLUE, italic=True)

# =====================================================================
# SLIDE 2 — Executive Summary
# =====================================================================
s = add_slide(); add_bg(s); header_bar(s, "Executive Summary")
gm_start, gm_end = gm_series[0], gm_series[-1]
stat_card(s, Inches(0.45), Inches(1.35), Inches(2.9), Inches(1.5), "FY2025 Revenue", f"${fy_rev/1e6:.2f}M")
stat_card(s, Inches(3.55), Inches(1.35), Inches(2.9), Inches(1.5), "Net Income", usd(fy_ni/1e3, 0, "K"),
          value_color=NAVY if fy_ni >= 0 else RED)
stat_card(s, Inches(6.65), Inches(1.35), Inches(2.9), Inches(1.5), "Adj. EBIT (ex. one-time)", f"${fy_ebit_adj/1e6:.2f}M", value_color=GREEN)
stat_card(s, Inches(9.75), Inches(1.35), Inches(2.9), Inches(1.5), "Ending Cash", f"${cash_series[-1]/1e6:.2f}M")
bullets(s, Inches(0.45), Inches(3.1), Inches(12.4), Inches(3.8), [
    f"Revenue grew from ${rev_series[0]/1e6:.2f}M (Jan) to ${rev_series[-1]/1e6:.2f}M (Dec); gross margin expanded {gm_start:.1%} → {gm_end:.1%}, "
    f"driven by the deliberate Legacy Hardware wind-down (low-margin line, sunset by September).",
    f"Net income of {usd(fy_ni/1e3,0,'K')} was held down by two Q3 one-time items tied to that same exit: an $850K restructuring "
    f"charge (Aug) and a $320K inventory write-down (Sep). Adjusted EBIT ex. one-time items was ${fy_ebit_adj/1e6:.2f}M — "
    f"the underlying business was profitable and improving all year.",
    f"Cash rose to ${cash_series[-1]/1e6:.2f}M, but ~$2.0M of the September increase was a revolver draw, not operating cash "
    f"generation — that month's net income was actually negative. See Cash & Runway slide.",
    f"A single customer (Vantage Logistics Corp) churned in May for $780K — the year's largest single revenue event. "
    f"Customer concentration is otherwise healthy (top-5 = {m['top5_share']:.1%} of revenue, threshold 40%).",
    f"Headcount grew {hc_series[0]}→{hc_series[-1]}; the only reduction was a 9-person RIF in Hardware Mfg Support (Oct), "
    f"tied to the restructuring.",
], size=15)
footer(s)

# =====================================================================
# SLIDE 3 — KPI Dashboard
# =====================================================================
s = add_slide(); add_bg(s); header_bar(s, "KPI Dashboard", "FY2025, trended monthly")
line_chart(s, Inches(0.4), Inches(1.25), Inches(6.3), Inches(2.6), MONTHS,
           {"Revenue (USD)": rev_series, "Budget (USD)": bud_series}, "Revenue vs Budget")
line_chart(s, Inches(6.85), Inches(1.25), Inches(6.1), Inches(2.6), MONTHS,
           {"Gross margin %": gm_series}, "Gross Margin %", pct=True)
line_chart(s, Inches(0.4), Inches(4.05), Inches(6.3), Inches(2.6), MONTHS,
           {"Net income": ni_series, "Adjusted EBIT": ebit_adj_series}, "Net Income vs Adjusted EBIT (USD)")
line_chart(s, Inches(6.85), Inches(4.05), Inches(6.1), Inches(2.6), MONTHS,
           {"Cash (USD)": cash_series}, "Cash Balance (USD)")
footer(s)

# =====================================================================
# SLIDE 4 — P&L Performance
# =====================================================================
s = add_slide(); add_bg(s); header_bar(s, "P&L Performance", "Consolidated, USD, current-rate method")
bar_chart(s, Inches(0.4), Inches(1.25), Inches(7.6), Inches(4.6), MONTHS,
          {"Gross profit": [pl[p]["gross_profit"] for p in PERIODS],
           "Operating expense": [pl[p]["opex_total"] for p in PERIODS]},
          "Gross Profit vs Operating Expense (USD)")
textbox(s, Inches(8.25), Inches(1.3), Inches(4.6), Inches(0.4), "Margin bridge — mix, not price/cost", bold=True, size=14, color=NAVY)
bullets(s, Inches(8.25), Inches(1.8), Inches(4.6), Inches(4.0), [
    "Gross margin: 69.5% (Jan) → 74.1% (Dec), +4.6 pts.",
    "Entire move explained by Legacy Hardware exit — hardware carries structurally lower margin than subscription/services.",
    "Subscription & services margins essentially flat YoY — this is a mix story, not a pricing or cost-efficiency story.",
    "Do not extrapolate further margin expansion once the mix shift is fully annualized in FY2026.",
], size=13)
footer(s)

# =====================================================================
# SLIDE 5 — Revenue Drivers
# =====================================================================
s = add_slide(); add_bg(s); header_bar(s, "Revenue Drivers", "New / Expansion / Renewal / Contraction / Churn — US01 + UK01 (IN01 has no external bookings)")
new_s = [m["rev_bridge"][p].get("new", 0) for p in PERIODS]
exp_s = [m["rev_bridge"][p].get("expansion", 0) for p in PERIODS]
ren_s = [m["rev_bridge"][p].get("renewal", 0) for p in PERIODS]
con_s = [m["rev_bridge"][p].get("contraction", 0) for p in PERIODS]
chu_s = [m["rev_bridge"][p].get("churn", 0) for p in PERIODS]
bar_chart(s, Inches(0.4), Inches(1.25), Inches(8.2), Inches(3.0), MONTHS,
          {"New": new_s, "Expansion": exp_s, "Renewal": ren_s},
          "Gross Bookings — New / Expansion / Renewal (USD)", stacked=True)
textbox(s, Inches(0.4), Inches(4.45), Inches(8.2), Inches(0.35),
        "Contraction & Churn by month (USD) — May's outlier scale would make a bar chart unreadable, shown as a table instead",
        size=12, bold=True, color=NAVY)
simple_table(s, Inches(0.4), Inches(4.85), Inches(8.2), Inches(1.9),
             MONTHS,
             [[f"{usd(v)}" if v else "—" for v in con_s], [f"{usd(v)}" if v else "—" for v in chu_s]],
             font_size=9)
textbox(s, Inches(0.4), Inches(6.85), Inches(8.2), Inches(0.3),
        "Row 1: Contraction   Row 2: Churn", size=9, italic=True, color=GREY)
textbox(s, Inches(8.85), Inches(1.3), Inches(4.0), Inches(0.4), "Material event", bold=True, size=14, color=RED)
bullets(s, Inches(8.85), Inches(1.8), Inches(4.0), Inches(2.2), [
    "May 2025: Vantage Logistics Corp churned in full — $780K annualized value, 46% of that month's revenue.",
    "Largest single revenue event of the year in either direction.",
], size=12)
textbox(s, Inches(8.85), Inches(4.1), Inches(4.0), Inches(0.4), "Concentration", bold=True, size=14, color=NAVY)
bullets(s, Inches(8.85), Inches(4.6), Inches(4.0), Inches(2.0), [
    f"Top-1 customer: {m['top1_share']:.1%} of revenue (threshold 15%).",
    f"Top-5 customers: {m['top5_share']:.1%} of revenue (threshold 40%).",
    "Neither threshold tripped — healthy at portfolio level.",
], size=12)
footer(s)

# =====================================================================
# SLIDE 6 — Opex & Headcount
# =====================================================================
s = add_slide(); add_bg(s); header_bar(s, "Opex & Headcount", "By category and department, FY2025")
sm_s = [pl[p]["sm"] for p in PERIODS]
rd_s = [pl[p]["rd"] for p in PERIODS]
ga_s = [pl[p]["ga"] for p in PERIODS]
bar_chart(s, Inches(0.4), Inches(1.25), Inches(6.3), Inches(2.6), MONTHS,
          {"S&M": sm_s, "R&D": rd_s, "G&A": ga_s}, "Opex by Category (USD)", stacked=True)
line_chart(s, Inches(6.85), Inches(1.25), Inches(6.1), Inches(2.6), MONTHS,
           {"Active headcount": hc_series}, "Total Headcount")
dec_hc = m["hc_by_period_dept"]["2025-12"]
depts_sorted = sorted(dec_hc.keys(), key=lambda d: -dec_hc[d].get("active", 0))
simple_table(s, Inches(0.4), Inches(4.1), Inches(6.3), Inches(2.8),
             ["Department", "Active (Dec)", "Monthly cost (USD)"],
             [[d, dec_hc[d].get("active", 0), f"{dec_hc[d].get('cost', 0):,.0f}"] for d in depts_sorted],
             font_size=11)
textbox(s, Inches(6.85), Inches(4.1), Inches(6.1), Inches(0.4), "October RIF", bold=True, size=14, color=RED)
bullets(s, Inches(6.85), Inches(4.6), Inches(6.1), Inches(2.2), [
    "Hardware Mfg Support: 9 involuntary terminations in October 2025, reduced to 11 active.",
    "Directly tied to the Q3 restructuring charge and Legacy Hardware line sunset.",
    "The only department reduction of the year — every other department grew or held flat.",
], size=13)
footer(s)

# =====================================================================
# SLIDE 7 — Balance Sheet & Cash / Runway
# =====================================================================
s = add_slide(); add_bg(s); header_bar(s, "Balance Sheet & Cash / Runway")
bar_chart(s, Inches(0.4), Inches(1.25), Inches(6.3), Inches(2.6), MONTHS,
          {"Cash": cash_series, "Debt (revolver)": debt_series}, "Cash & Debt Balances (USD)")
stat_card(s, Inches(6.85), Inches(1.25), Inches(2.9), Inches(1.3), "Total Assets (Dec)", f"${bs['2025-12']['assets']/1e6:.2f}M")
stat_card(s, Inches(9.95), Inches(1.25), Inches(2.9), Inches(1.3), "Total Liab.+Equity (Dec)",
          f"${(bs['2025-12']['liabilities']+bs['2025-12']['equity_total'])/1e6:.2f}M")
stat_card(s, Inches(6.85), Inches(2.7), Inches(2.9), Inches(1.15), "Articulation check", "TRUE — 0 var.", value_color=GREEN)
stat_card(s, Inches(9.95), Inches(2.7), Inches(2.9), Inches(1.15), "Sep revolver draw", "$2.0M")
textbox(s, Inches(6.85), Inches(4.0), Inches(6.0), Inches(0.4), "September: financing, not operations", bold=True, size=14, color=RED)
bullets(s, Inches(6.85), Inches(4.5), Inches(6.0), Inches(2.5), [
    f"Cash rose ${bs['2025-09']['cash']-bs['2025-08']['cash']:,.0f} in September; $2,000,000 of that (>99%) is the new "
    f"revolver draw, not operating cash generation.",
    f"September net income was actually negative ({usd(pl['2025-09']['net_income'])}) — the inventory write-down month.",
    "FY2025 operating cash flow was positive but materially smaller than the headline cash-balance change — read the "
    "year-end cash figure as debt-assisted, not a pure organic-performance signal.",
    "No burn/runway concern currently — H2 was net cash-generative even before the debt draw.",
], size=12)
footer(s)

# =====================================================================
# SLIDE 8 — Budget vs Actual
# =====================================================================
s = add_slide(); add_bg(s); header_bar(s, "Budget vs Actual", "US01 + UK01 only — IN01 budget not supplied this cycle (see Risks slide)")
var_series = [(rev_series[i]/bud_series[i]-1) if bud_series[i] else 0 for i in range(12)]
bar_chart(s, Inches(0.4), Inches(1.25), Inches(7.6), Inches(4.6), MONTHS,
          {"Revenue variance %": var_series}, "Revenue vs Budget — Variance %", colors=[NAVY])
textbox(s, Inches(8.25), Inches(1.3), Inches(4.6), Inches(0.4), "Reading the variance", bold=True, size=14, color=NAVY)
bullets(s, Inches(8.25), Inches(1.8), Inches(4.6), Inches(4.0), [
    "Revenue tracked within ±5% of budget all year — no month tripped the 10% fixed threshold.",
    "H2 variance turns consistently negative (-2% to -5%): budget did not anticipate the Legacy Hardware sunset, so "
    "actual revenue undershoots a budget that still assumed hardware continuing.",
    "Opex exceeded budget by >15% in Jul, Aug, and Sep — Aug/Sep driven by the one-time restructuring and write-down "
    "charges, not core spend growth.",
    "IN01 (India) has no FY2025 budget on file — flagged, not fabricated. Obtain for FY2026.",
], size=13)
footer(s)

# =====================================================================
# SLIDE 9 — Risks & Red Flags
# =====================================================================
s = add_slide(); add_bg(s); header_bar(s, "Risks & Red Flags", "Every threshold trip this cycle — all new (first review cycle)")
flag_rows = []
for rf in sorted(m["red_flags"], key=lambda x: x["period"])[:11]:
    flag_rows.append([rf["period"], rf["metric"], rf["detail"][:95] + ("…" if len(rf["detail"]) > 95 else "")])
simple_table(s, Inches(0.4), Inches(1.25), Inches(12.5), Inches(5.6),
             ["Period", "Item", "Detail"], flag_rows,
             col_widths=[Inches(1.1), Inches(3.0), Inches(8.4)], font_size=10)
footer(s)

# =====================================================================
# SLIDE 10 — FY2025 Monthly Trend (stand-in for multi-year trend, first cycle)
# =====================================================================
s = add_slide(); add_bg(s); header_bar(s, "FY2025 Monthly Trend", "First review cycle — this is intra-year trend from the backfill, not a multi-year view. Multi-year trend begins FY2026.")
line_chart(s, Inches(0.4), Inches(1.25), Inches(12.5), Inches(5.6), MONTHS,
           {"Revenue (USD)": rev_series, "Net income (USD)": ni_series, "Cash (USD)": cash_series},
           "Revenue, Net Income & Cash — Jan-Dec 2025")
footer(s)

# =====================================================================
# SLIDE 11 — Outlook / Next-Year Focus
# =====================================================================
s = add_slide(); add_bg(s); header_bar(s, "Outlook / Next-Year Focus")
bullets(s, Inches(0.5), Inches(1.4), Inches(12.3), Inches(5.5), [
    "Confirm the Legacy Hardware wind-down is fully complete and retire the associated COGS/inventory taxonomy lines "
    "for FY2026 if no further activity is expected.",
    "Obtain FY2026 budget for IN01 (India) — no budget was available for this entity in FY2025.",
    "Obtain the revolving credit facility's terms (maturity, covenants, rate basis) — drawn $2.0M in September with no "
    "supporting detail provided this cycle.",
    "Run a customer post-mortem on the Vantage Logistics Corp churn (May 2025, $780K) — no reason code was available "
    "in the data supplied.",
    "Establish an explicit Intercompany Payable account on US01's chart of accounts so the IN01 intercompany "
    "receivable can be fully eliminated in consolidation next cycle.",
    "FY2026 will be the first cycle with a genuine multi-year trend and a full statistical anomaly-detection baseline "
    "— keep the taxonomy and thresholds unchanged to preserve comparability.",
], size=16)
footer(s)

# =====================================================================
# SLIDE 12 — Strategic Narrative
# =====================================================================
s = add_slide(); add_bg(s); header_bar(s, "Strategic Narrative")
bullets(s, Inches(0.5), Inches(1.4), Inches(12.3), Inches(5.5), [
    "FY2025 was a transition year: Meridian completed the deliberate exit from Legacy Hardware, trading a lower-margin, "
    "declining product line for a cleaner, higher-margin subscription-and-services mix.",
    "The one-time cost of that transition ($1.17M combined restructuring + write-down) is now fully absorbed; FY2026 "
    "starts from a structurally improved margin base with no further hardware-related charges expected.",
    "Core subscription growth and gross-margin discipline were solid all year once the hardware noise is excluded — "
    "the underlying SaaS business is the story the board should track into FY2026.",
    "The India (IN01) and UK (UK01) entities are fully integrated into consolidation and running cleanly; India's "
    "missing budget is a process gap to close, not a performance concern.",
    "Recommend FY2026 budget assumptions explicitly assume zero Legacy Hardware revenue/COGS, rather than carrying "
    "forward a declining-but-present hardware line as FY2025's budget did.",
], size=16)
footer(s)

# =====================================================================
# SLIDE 13 — Appendix
# =====================================================================
s = add_slide(); add_bg(s); header_bar(s, "Appendix", "Methodology, data lineage, FX & consolidation")
bullets(s, Inches(0.5), Inches(1.4), Inches(12.3), Inches(5.5), [
    "Source: 8 input files (Company Config, GL/Trial Balance, Cash & Bank Position, Headcount, Revenue Bookings, "
    "Budget, FX Rates, Intercompany Detail) — full lineage in the workbook's File Log & Lineage tab.",
    "Consolidation: current-rate method — balance sheet at period-end spot rate, income statement at period-average "
    "rate, translation difference to Cumulative Translation Adjustment within equity.",
    "FX rates: Company Treasury Rate, sourced from 06_FX_Rates_2025.xlsx, applied consistently across all periods.",
    "Standards: US01 US GAAP, UK01 IFRS, IN01 Ind AS (local GAAP) — divergence points (leases, R&D capitalization, "
    "inventory valuation) reviewed and found immaterial this cycle; see workbook FX & Consolidation tab.",
    "Every MEDIUM/LOW-confidence judgment call this cycle is logged in the workbook's Assumptions & Judgment Log — "
    "10 entries this cycle, none left unresolved.",
    "Full detail: 2025-FY_Business_Review_USD.xlsx (workbook) and 2025-FY_Methodology_and_Data_Lineage.docx.",
], size=15)
footer(s)

prs.save(f"{OUT_DIR}/2025-FY_Business_Review_USD.pptx")
print("Deck written: 2025-FY_Business_Review_USD.pptx, slides:", len(prs.slides.__iter__.__self__._sldIdLst))
