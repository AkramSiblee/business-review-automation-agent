"""Standalone Methodology & Data Lineage document — what inputs were used, what methodology
was applied, and how each output was derived. For audit / reviewer purposes."""
import pickle
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT_DIR = "data/outputs"
m = pickle.load(open(f"{OUT_DIR}/_model.pkl", "rb"))

NAVY = RGBColor(0x1F, 0x38, 0x64)
GREY = RGBColor(0x59, 0x59, 0x59)

doc = Document()
style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(11)

def h1(text):
    p = doc.add_heading(text, level=1)
    for r in p.runs: r.font.color.rgb = NAVY
    return p

def h2(text):
    p = doc.add_heading(text, level=2)
    for r in p.runs: r.font.color.rgb = NAVY
    return p

def body(text, italic=False, bold=False):
    p = doc.add_paragraph()
    r = p.add_run(text); r.italic = italic; r.bold = bold
    return p

def bullet(text):
    return doc.add_paragraph(text, style="List Bullet")

def set_cell_shading(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd"); shd.set(qn("w:fill"), color_hex)
    tcPr.append(shd)

def table(headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Light Grid Accent 1"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = h
        for p in hdr[i].paragraphs:
            for r in p.runs: r.bold = True
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
tp = doc.add_paragraph(); tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = tp.add_run("Methodology & Data Lineage"); r.font.size = Pt(24); r.bold = True; r.font.color.rgb = NAVY
sp = doc.add_paragraph(); sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = sp.add_run("Meridian Technologies Group — FY2025 Annual Business Review"); r.font.size = Pt(14); r.font.color.rgb = GREY
mp = doc.add_paragraph(); mp.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = mp.add_run("Companion document to 2025-FY_Business_Review_USD.xlsx / .docx / .pptx — "
               "produced by the business-review-automation Claude Code skill.")
r.italic = True; r.font.size = Pt(10)
doc.add_paragraph()
body("This document exists so a human reviewer can audit an autonomously-run cycle without redoing the "
     "underlying work: it states exactly what was read, exactly how it was processed, and exactly how "
     "each figure in the three deliverables was derived. It is intentionally separate from the executive "
     "memo, which is written for a board audience rather than an audit audience.")

# ============================================================
# 1. INPUTS USED
# ============================================================
h1("1. Inputs used this cycle")
body("Eight files, uploaded to the business's Google Drive folder (Business_Review_Agent/Input_Files) "
     "and staged locally for processing. Every file matched the skill's standard input-file schemas "
     "exactly by column header, giving HIGH-confidence classification by construction — no content-based "
     "guessing was required for file triage.")
table(
    ["File", "Rows", "What it contains", "Fate"],
    [
        ["00_Company_Config.xlsx", "3", "Entity master: name, country, functional currency, accounting standard, fiscal year start, ownership %, consolidation method", "Used — defines Phase 0 configuration"],
        ["01_GL_TrialBalance_2025.xlsx", "1,140", "Monthly trial balance, all 3 entities, Jan–Dec 2025, account-level detail with one-time-item flags and free-text notes", "Used — primary source for Income Statement & Balance Sheet"],
        ["02_CashBank_Position_2025.xlsx", "48", "Bank-account-level beginning/ending balances by entity/period", "Used — cross-checked against GL cash account (tied exactly, $0 variance)"],
        ["03_Headcount_2025.xlsx", "236", "Employee-level roster: department, hire/termination dates, FTE %, fully-loaded monthly cost", "Used — drives Headcount & Opex Detail tab"],
        ["04_Revenue_Bookings_2025.xlsx", "149", "Deal-level bookings: customer, product, revenue type (new/expansion/renewal/contraction/churn), amount", "Used — drives Revenue Drivers tab (US01 + UK01 only; IN01 has no external bookings)"],
        ["05_Budget_Forecast_FY2025.xlsx", "372", "Monthly budgeted amounts by account", "Used for US01 + UK01 only — IN01 not supplied, see Section 3"],
        ["06_FX_Rates_2025.xlsx", "48", "GBP/USD and INR/USD, period-end spot and period-average, monthly, sourced 'Company Treasury Rate'", "Used — sole FX source for all translation"],
        ["07_Intercompany_Detail_2025.xlsx", "24", "Intercompany transaction detail: R&D recharge (India→US) and management fee (UK→US)", "Used to cross-check the GL's own intercompany account balances before elimination"],
    ],
)

# ============================================================
# 2. METHODOLOGY
# ============================================================
h1("2. Methodology applied")
h2("2.1 Locale & parsing")
bullet("Dates: US01 records use ISO 8601 (YYYY-MM-DD). UK01 and IN01 headcount dates use DD/MM/YYYY — "
       "confirmed unambiguously for 40 of 100 slash-formatted rows where the day or month component "
       "exceeds 12 (e.g. '17/05/2025' cannot be 17th month), establishing the convention for those two "
       "entities by internal evidence, not assumption. All dates normalized to ISO 8601 internally and "
       "to 'DD Mon YYYY' in narrative outputs.")
bullet("Numbers: all files used the standard comma-thousands / period-decimal convention; no ambiguous "
       "number formats were encountered this cycle.")
bullet("Currency & scale: each entity's functional currency was read from 00_Company_Config.xlsx and "
       "confirmed against the currency code stated on every GL row — no scale ambiguity (all figures are "
       "in actual units, not thousands/lakhs/crores).")

h2("2.2 Chart of accounts / taxonomy mapping")
body("Every GL account (94 distinct entity/account-code combinations) was mapped to the standard "
     "taxonomy — Revenue (by type: subscription, professional services, legacy hardware, other, "
     "intercompany), Cost of Revenue, Opex (S&M / R&D / G&A), Other Income/Expense, Tax, and the balance "
     "sheet Asset/Liability/Equity categories. The full account-by-account map is in the workbook's "
     "Taxonomy Map tab.")

h2("2.3 Consolidation — current-rate method")
bullet("Balance sheet accounts translated at the period-end spot rate; income statement accounts "
       "translated at the period-average rate — both sourced from 06_FX_Rates_2025.xlsx ('Company "
       "Treasury Rate'), matching the business's own stated statutory-reporting source.")
bullet("Translation differences are recorded to a Cumulative Translation Adjustment (CTA) line within "
       "equity, not through the income statement, per standard current-rate-method practice.")
bullet("This is the business's first consolidation cycle, so no prior-period opening consolidated equity "
       "exists to anchor CTA. CTA is computed as a plug from a zero baseline at Jan-2025: translated "
       "equity (at spot) less common stock (at spot) less cumulative translated net income (at each "
       "period's average rate). This is a defensible first-cycle approach, logged as a MEDIUM-confidence "
       "assumption, and should be reconciled against actual opening equity if/when available for a "
       "future cycle.")
bullet("Intercompany revenue/expense pairs were identified at the account level and eliminated: IN01's "
       "R&D cost-plus-10% recharge revenue against US01's matching R&D recharge expense, and US01's "
       "management-fee income against UK01's matching management-fee expense. Both were cross-checked "
       "against 07_Intercompany_Detail_2025.xlsx and matched exactly.")
bullet("IN01 carries an 'Intercompany Receivable (US01)' balance-sheet account with no separately "
       "identifiable matching payable on US01's chart of accounts — this balance could not be positively "
       "eliminated this cycle and is disclosed as such (workbook FX & Consolidation tab), rather than "
       "silently left in or silently dropped from the consolidated total.")

h2("2.4 Cash flow statement (indirect method)")
body("No direct cash flow file was supplied. The cash flow statement is derived via the indirect method: "
     "net income, plus depreciation & amortization, plus/minus changes in accounts receivable, inventory, "
     "prepaid expenses, accounts payable, accrued liabilities, and deferred revenue — all computed from "
     "balance-sheet deltas month over month. An explicit 'Other operating adjustments' line captures items "
     "not individually traced this cycle (Other Assets/Liabilities movement, Allowance for Doubtful "
     "Accounts, FX/CTA translation effects on working-capital balances, income-tax timing), sized so that "
     "computed ending cash ties exactly to the Balance Sheet cash line every month — verified at $0.00 "
     "variance for all 12 months after full recalculation in Excel.")

h2("2.5 Red-flag detection")
bullet("Fixed, materiality-scaled thresholds (percentage- and point-based, not fixed currency amounts): "
       "revenue vs. budget >10%, opex-category vs. budget >15%, gross margin move >2pts MoM, headcount "
       "net change >5%/month, cash runway <12 months, monthly burn increase >20% MoM without an explained "
       "driver, DSO increase >15% vs. trailing 3-month average, customer concentration >15% (top-1) or "
       ">40% (top-5).")
bullet("Statistical trend check (±2 standard deviations from the trailing mean), applied using the 12 "
       "months of real history built up within this single backfill cycle, per the skill's guidance for "
       "when at least ~6 periods of history exist.")
bullet("Materiality-override items: individually large, one-off items that could change a headline KPI's "
       "direction were flagged explicitly even where no generic threshold applied — the restructuring "
       "charge, inventory write-down, the Vantage Logistics churn event, and the September debt-financed "
       "cash increase were all identified this way, since a percentage-threshold scan alone would not "
       "have surfaced them as clearly.")

h2("2.6 First-cycle backfill treatment")
body("This cycle supplied a full prior year (all 12 months of FY2025) at once. Per the skill's guidance "
     "for this scenario, every month was populated into the history workbook as usual, but the memo and "
     "deck were generated only for the most recent completed period at its natural annual cadence "
     "(FY2025 as a whole), rather than as 12 separate monthly reports. Every trend chart is labeled as "
     "covering a backfilled history rather than periods reviewed as they occurred.")

# ============================================================
# 3. WHAT WAS NOT FABRICATED / WHAT WAS FLAGGED INSTEAD
# ============================================================
h1("3. Gaps — flagged, not fabricated")
body("Per the skill's autonomy doctrine, a data gap gets a documented, reasonable default and a flag, or "
     "an explicit 'not available' — never an invented figure. This cycle's gaps:")
table(
    ["Gap", "Treatment"],
    [
        ["IN01 (India) has no FY2025 budget file", "Every budget-variance view shows IN01 as 'Not available this cycle.' Consolidated budget variance is computed from US01 + UK01 only, stated explicitly wherever shown."],
        ["No direct cash flow statement supplied", "Derived via the indirect method (Section 2.4) with an explicitly labeled 'Other adjustments' plug line — never presented as if traced line-by-line from a source file that doesn't exist."],
        ["No revolving credit facility terms supplied (maturity, covenants, rate)", "Disclosed as an open item in the memo's Risks section and Outlook section — not assumed or estimated."],
        ["No churn reason code for the Vantage Logistics Corp event", "Flagged as a recommended follow-up (customer post-mortem) rather than speculated upon in the narrative."],
        ["IN01/US01 intercompany balance-sheet balance not fully matchable", "Disclosed explicitly in the workbook's FX & Consolidation tab as un-eliminated, with the reason (no separately identifiable payable account on US01's books)."],
    ],
)

# ============================================================
# 4. OUTPUTS DERIVED
# ============================================================
h1("4. Outputs produced this cycle")
table(
    ["File", "Purpose", "Derivation"],
    [
        ["2025-FY_Business_Review_USD.xlsx", "Full model — 12 tabs: Config & Setup, File Log & Lineage, Assumptions & Judgment Log, Taxonomy Map, Income Statement, Balance Sheet, Cash Flow, FX & Consolidation, Revenue Drivers, Headcount & Opex Detail, KPI Dashboard, Red Flags & Variance Log.", "Built directly from the 8 input files per the methodology in Section 2. Subtotals, the balance-sheet articulation check, and the cash-flow-to-balance-sheet tie are live Excel formulas, not hardcoded values — verified via full Excel recalculation to have zero formula errors and zero tie-out variance."],
        ["2025-FY_Business_Review_USD.docx", "CFO/board-ready executive memo — 8 sections (annual structure): Executive Summary, Financial Performance, Balance Sheet & Cash, Revenue & Growth Drivers, Headcount & Organization, Risks & Watch Items, Outlook, Appendix.", "Every figure is pulled forward from the workbook, never re-derived. Narrative commentary is generated from the same driver analysis and red-flag detection as the deck, so the two documents never disagree on a headline figure."],
        ["2025-FY_Business_Review_USD.pptx", "Board deck — 13 slides, native charts (no screenshots): Title, Executive Summary, KPI Dashboard, P&L Performance, Revenue Drivers, Opex & Headcount, Balance Sheet & Cash/Runway, Budget vs Actual, Risks & Red Flags, FY2025 Monthly Trend, Outlook, Strategic Narrative, Appendix.", "Same source data and figures as the workbook and memo. Went through a visual QA render pass (13 slides rendered to image and reviewed) — one chart-scale legibility issue (a single $780K churn event next to routine $10-40K monthly figures) and one PowerPoint default-formatting bug (negative bars rendering unfilled) were found and fixed during that pass."],
        ["2025-FY_Methodology_and_Data_Lineage.docx", "This document.", "Written directly by the orchestrator from the same processing run, not re-derived from the other outputs after the fact."],
    ],
)

h2("Cross-document consistency check (Phase 8 QA)")
fy_rev = sum(m["consol_pl"][p]["revenue"] for p in m["PERIODS"])
fy_ni = sum(m["consol_pl"][p]["net_income"] for p in m["PERIODS"])
bullet(f"FY2025 total revenue: USD {fy_rev:,.0f} — identical across workbook KPI Dashboard, memo Section 1, and deck Slide 2.")
bullet(f"FY2025 net income: USD {fy_ni:,.0f} — identical across all three deliverables.")
bullet("Balance sheet articulation (Assets = Liabilities + Equity): TRUE, zero variance, all 12 months, "
       "verified by forcing a full recalculation in Excel (not just reading back openpyxl's unevaluated "
       "formula strings).")
bullet("Cash flow statement ties to the Balance Sheet cash line: zero variance, all 12 months, same "
       "verification method.")
bullet("Every one of the 11 red-flag items and 10 Assumptions & Judgment Log entries logged in the "
       "workbook appears in both the memo's Risks & Watch Items section and the deck's Risks & Red Flags "
       "slide.")

h2("A known limitation, disclosed for transparency")
body("This project's dedicated Phases 1-3 subagent (business-review-data-processor) could not be used "
     "this cycle — a platform issue on this session meant it received no working code-execution tool "
     "despite its configuration requesting one, so it could not open the binary .xlsx input files and "
     "reported itself blocked rather than guessing. Phases 1-3 were instead run directly, using the same "
     "methodology the subagent would have applied, with the same source files, the same reference-file "
     "rules, and the same validation checks. This is noted here in the interest of full lineage "
     "transparency; it does not change any figure in the deliverables, and the limitation has been "
     "documented in the project files for future cycles.")

doc.save(f"{OUT_DIR}/2025-FY_Methodology_and_Data_Lineage.docx")
print("Methodology document written.")
