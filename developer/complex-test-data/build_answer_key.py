#!/usr/bin/env python3
"""
Builds the user-facing, plain-English answer key for the Halcyon complex test dataset:

  sample-data/answer-keys/Complex_Test_Answer_Key_FY2025.xlsx

  python developer/complex-test-data/build_answer_key.py

Figures are read from EXPECTED.json (next to this script) so the key never drifts from the dataset. Re-run after
generate_complex_dataset.py. Audience: a non-technical reviewer who wants to know what a cycle produces, how it
gets there, what the test file hides, and whether a given run got it right.
"""
import json
import os

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(os.path.dirname(HERE))
EXP = json.load(open(os.path.join(HERE, "EXPECTED.json"), encoding="utf-8"))
OUT_DIR = os.path.join(REPO, "sample-data", "answer-keys")
OUT = os.path.join(OUT_DIR, "Complex_Test_Answer_Key_FY2025.xlsx")

NAVY = "1F3864"
FILL_H = PatternFill("solid", fgColor=NAVY)
FILL_SEC = PatternFill("solid", fgColor="D9E2F3")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
FILL_WARN = PatternFill("solid", fgColor="FCE4D6")
F_TITLE = Font(name="Calibri", size=16, bold=True, color=NAVY)
F_SUB = Font(name="Calibri", size=11, italic=True, color="595959")
F_H = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
F_B = Font(name="Calibri", size=11, bold=True)
F_INPUT = Font(name="Calibri", size=11, color="0000FF")
WRAP = Alignment(wrap_text=True, vertical="top")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def get(path):
    v = EXP
    for k in path.split("."):
        v = v[k]
    return v


def title(ws, text, sub):
    ws["A1"] = text
    ws["A1"].font = F_TITLE
    ws["A2"] = sub
    ws["A2"].font = F_SUB


def table(ws, row, headers, rows, widths=None):
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.font, cell.fill, cell.alignment, cell.border = F_H, FILL_H, WRAP, BOX
    for r, vals in enumerate(rows, row + 1):
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.alignment, cell.border = WRAP, BOX
    if widths:
        for c, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(c)].width = w
    return row + len(rows)


def paragraphs(ws, start, items, width=110):
    ws.column_dimensions["A"].width = width
    r = start
    for kind, text in items:
        cell = ws.cell(row=r, column=1, value=text)
        cell.alignment = WRAP
        if kind == "h":
            cell.font, cell.fill = F_B, FILL_SEC
        elif kind == "warn":
            cell.font, cell.fill = F_B, FILL_WARN
        r += 1
    return r


def page_setup(ws, landscape=True):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.sheet_view.showGridLines = False


wb = Workbook()

# ---------------------------------------------------------------- Start Here
ws = wb.active
ws.title = "Start Here"
title(ws, "Answer Key: Complex Test Data (Halcyon Industrial Holdings, FY2025)",
      "Fictional company. All figures are synthetic and exist only to test the business review pipeline.")
paragraphs(ws, 4, [
    ("h", "What this file is"),
    ("p", "The test input (sample-data/complex-test-data/Business_Review_Complex_Input_FY2025.xlsx) is a fictional group "
          "of 8 companies in 7 currencies. About 60 problems have been hidden in it on purpose: messy data, tricky "
          "accounting, and warning signs a CFO would want raised. This answer key lists every one of them, what a correct "
          "review should do about each, and the figures it should arrive at."),
    ("warn", "Never copy this answer key into the inputs folder. Copy ONLY the test input workbook. If the review can see "
             "the answers, the test proves nothing."),
    ("h", "How to use it (3 steps)"),
    ("p", "1. Copy Business_Review_Complex_Input_FY2025.xlsx (and nothing else) into the inputs folder."),
    ("p", "2. Ask Claude: \"Run a business review cycle on the complex test data.\" It first produces a Data Validation "
          "Report and pauses for you. Reply to its questions (see tab 'How It Works'), then it builds the workbook, memo and deck."),
    ("p", "3. Open the outputs and work through 'Issues Checklist' (mark each issue Found / Partly / Missed) and type the "
          "review's figures into the yellow cells on 'Key Numbers'. 'Scorecard' then totals the result."),
    ("h", "Tabs in this file"),
    ("p", "What You Get: the five files a review cycle produces, and what is in each."),
    ("p", "How It Works: the review's steps in plain English, including where it stops to ask you."),
    ("p", "The Company: the 8 entities in the test group and what makes each one hard."),
    ("p", "Issues Checklist: every planted issue, where it is, what a correct review does, why it matters, and where to look for it."),
    ("p", "Key Numbers: the figures a correct review should reproduce, with a pass/check column."),
    ("p", "Scorecard: your totals, plus the four things every good run must get right."),
    ("h", "Reading the checklist"),
    ("p", "'Caught automatically' = Yes means the pipeline's built-in checker flags it in the Data Validation Report before "
          "any analysis. 'Needs judgment' means the review has to spot it during analysis, the way an experienced analyst would."),
    ("p", "Confidence levels you will see in the outputs: HIGH = clear evidence, handled silently and logged. MEDIUM = a "
          "reasonable assumption, handled and clearly logged. LOW = genuinely uncertain, so the review asks you rather than guessing."),
    ("h", "Basis of figures"),
    ("p", "Group reporting currency USD. Fiscal year 1 Jan to 31 Dec 2025. Mixed accounting standards (US GAAP, IFRS, Ind AS) "
          "as listed on 'The Company'. P&L figures translated at monthly average rates, balances at period-end rates, "
          "using the rates supplied inside the test file. Small rounding differences are expected."),
])
page_setup(ws, landscape=False)

# ---------------------------------------------------------------- What You Get
ws = wb.create_sheet("What You Get")
title(ws, "What a review cycle produces", "Five files land in the outputs folder, in this order.")
table(ws, 4, ["#", "File", "When", "What it contains", "How it is produced", "What to check on this test"], [
    [1, "<Period>_Data_Validation_Report.xlsx", "First, then the review pauses",
     "What data was expected vs received, what is missing, data-quality problems found (wrong formats, duplicates, "
     "missing months, unbalanced ledgers), and questions for you.",
     "A built-in checker reads every tab of the input file and compares it to the standard layout. Nothing is "
     "consolidated yet.",
     "All issues marked 'Caught automatically = Yes' on the Issues Checklist should appear here. The Turkey hyperinflation "
     "question should be raised."],
    [2, "<Period>_Business_Review_USD.xlsx (the model)", "After you reply to the checkpoint",
     "The full formula-driven model: income statement, balance sheet, cash flow, FX & consolidation, revenue drivers, "
     "headcount, KPI dashboard, red flags, plus extra tabs when triggered (Adjustments bridge, Liquidity & Covenants, "
     "Ownership & Scope, Constant Currency).",
     "All 8 entities are translated to USD, intercompany sales removed, ownership rules applied, then every KPI and "
     "red flag computed. Figures are live Excel formulas, so you can trace them.",
     "Figures on the 'Key Numbers' tab. The balance-sheet check row must read TRUE."],
    [3, "Assumptions & Judgment Log (a tab inside the model)", "With the model",
     "Every assumption and judgment call, with its confidence level and whether it was resolved.",
     "Written as the review goes, so a person can audit the automation without redoing it.",
     "Should list e.g. the Mexico August rate proxy, the earn-out treated as one-time, the Period 13 adjustment."],
    [4, "<Period>_Business_Review_USD.docx (the memo)", "After the model",
     "Executive summary with the top storylines, performance, balance sheet & cash, growth drivers, risks, outlook. "
     "Extra sections when triggered: Liquidity & Going-Concern Indicators, Earnings Quality, Scope & Comparability, FX.",
     "Written from the model. Numbers are pulled from the workbook, never retyped.",
     "Liquidity section placed straight after the executive summary; earnings-quality section shows reported vs adjusted."],
    [5, "<Period>_Business_Review_USD.pptx (the board deck)", "After the memo",
     "10-15 slides: executive summary, KPI scorecard, P&L, revenue bridge, opex & headcount, cash, budget vs actual, "
     "risks, outlook, appendix.",
     "Built from the same model and the same storylines as the memo, then each slide is rendered and visually checked.",
     "Slide 1 must lead with the covenant near-breach and the loan maturing June 2026."],
], widths=[4, 34, 20, 50, 45, 45])
ws.freeze_panes = "A5"
r = 11
for txt in ["Every file states its currency (ISO code), fiscal period and accounting basis on the cover.",
            "Before anything is delivered, automatic quality checks run: Excel recalculation with no formula errors, "
            "no placeholder text, and headline figures matching between memo and deck."]:
    ws.cell(row=r, column=2, value=txt).alignment = Alignment(wrap_text=False)
    r += 1
page_setup(ws)

# ---------------------------------------------------------------- How It Works
ws = wb.create_sheet("How It Works")
title(ws, "How the review gets there", "The steps, in plain English. You are involved at steps 1 and 3 only.")
table(ws, 4, ["Step", "What happens", "Your involvement", "On this test file"], [
    ["0. Setup (first run only)", "Confirms reporting currency, fiscal year, accounting standard per entity, materiality "
     "thresholds and how to consolidate. Later cycles reuse these settings.", "Confirm or correct once.",
     "Expect: USD, calendar year, mixed standards, full consolidation except the 40% India JV (equity method)."],
    ["1. Read the file", "Opens every tab, works out what each tab is, reads dates and numbers using the file's own "
     "evidence (e.g. a day above 12 proves DD/MM), detects currency, units and language.", "None.",
     "Has to cope with German, Portuguese and Spanish account names, thousands-scaled figures and three date styles."],
    ["2. Validate", "Checks ledgers balance, cash ties to the bank, headcount ties to payroll; spots duplicates, gaps, "
     "outliers and Excel errors.", "None.", "Finds the planted data-shape problems (section 'Data shape' on the checklist)."],
    ["3. CHECKPOINT: Data Validation Report", "Writes the report and STOPS. Nothing is consolidated until you reply.",
     "Read the report, answer its questions (batched into one round), say 'go ahead'.",
     "Expect a LOW-confidence question on Turkey's hyperinflation index. A good answer: 'no index available, present "
     "Turkey with a caveat'."],
    ["4. Consolidate", "Translates every entity to USD, removes intercompany trading, applies ownership rules "
     "(minority share, JV share, acquisition/disposal dates).", "None.",
     "Brazil 70%, India JV 40%, Mexico from 1 Jul, UK until 30 Sep."],
    ["5. Analyse drivers", "Explains why numbers moved: price/volume, new vs churned customers, headcount vs cost, "
     "budget variance, cash vs profit, covenant headroom.", "None.", "Recomputes the covenant the company reported."],
    ["6. Red flags & storylines", "Tests every metric against thresholds and trend, ranks the top 5 storylines.", "None.",
     "Covenant + loan maturity should rank first."],
    ["7. Build the deliverables", "Model first, then memo and deck from the model. Depth scales with monthly / "
     "quarterly / annual.", "None.", "Annual (FY2025) depth."],
    ["8. Quality checks", "Recalculates the workbook, scans for errors, renders and inspects every slide, checks the "
     "memo and deck agree.", "None, unless a check fails and cannot be fixed.", "All checks should PASS."],
], widths=[28, 55, 32, 55])
ws.freeze_panes = "A5"
page_setup(ws)

# ---------------------------------------------------------------- The Company
ws = wb.create_sheet("The Company")
title(ws, "The test group: Halcyon Industrial Holdings", "8 entities, 7 currencies, 3 accounting standards.")
ni = get("fy_net_income_native")
table(ws, 4, ["Entity", "Country", "Currency", "Accounting standard", "Ownership / how included", "FY2025 net income (own currency)", "What makes it hard"], [
    ["HQ01", "United States", "USD", "US GAAP", "Parent", ni["HQ01"], "Bank loan with covenants, LIFO inventory, some staff paid in CAD, a EUR bank account"],
    ["DE01", "Germany", "EUR", "IFRS", "100%, fully included", ni["DE01"], "German account names, figures exported as year-to-date totals, buys goods from HQ"],
    ["BR01", "Brazil", "BRL", "IFRS", "70% owned, fully included with 30% minority share", ni["BR01"], "Portuguese names, figures in thousands, April missing"],
    ["TR01", "Turkey", "TRY", "IFRS", "100%, fully included", ni["TR01"], "High-inflation economy, 20% currency drop in March, loss-making"],
    ["MX01", "Mexico", "MXN", "IFRS", "100%, bought 1 Jul 2025", ni["MX01"], "Spanish names, single signed amount column, Jul-Dec only"],
    ["UK01", "United Kingdom", "GBP", "IFRS", "100%, sold 30 Sep 2025", ni["UK01"], "Discontinued operation, Jan-Sep only"],
    ["JV01", "India", "INR", "Ind AS", "40% joint venture, one-line share of profit", ni["JV01"], "Profit & loss only, financial year starts in April"],
    ["HLD1", "Cayman Islands", "USD", "US GAAP", "100%, fully included", ni["HLD1"], "Dormant holding company, balance sheet only"],
], widths=[9, 16, 10, 18, 38, 22, 60])
for r in range(5, 13):
    ws.cell(row=r, column=6).number_format = "#,##0;(#,##0)"
page_setup(ws)

# ---------------------------------------------------------------- Issues Checklist
VR, WB_, MEMO, DECK = "Data Validation Report", "Model workbook", "Memo", "Deck"
ISSUES = [
    # (id, section, what is planted, where in the input, what a correct review does, why it matters, auto, where to look)
    ("2.1", "Data shape", "Germany's figures are cumulative year-to-date, not monthly.", "Main ledger tab, DE01 rows",
     "Convert to monthly amounts.", "Otherwise December would show a whole year of German revenue.", "Yes", VR),
    ("2.2", "Data shape", "Turkey's ledger contains 'Total Revenue' and 'Gross Profit' summary rows.", "Main ledger tab, TR01 rows with no account code",
     "Remove the summary rows.", "Otherwise Turkish revenue is counted twice.", "Yes", VR),
    ("2.3", "Data shape", "Months run across the columns instead of down the rows.", "Tab 09_Operational_KPIs",
     "Reshape to one row per month.", "Needed to calculate trends.", "Yes", VR),
    ("2.4", "Data shape", "Mexico uses one amount column with credits shown as negatives.", "Tab 'MX01 Ledger (post-acq)'",
     "Convert to debit/credit.", "Otherwise Mexico's books do not balance.", "Yes", VR),
    ("2.5", "Data shape", "Brazil's tab is in thousands of BRL.", "Tab 'BR01 GL (BRL 000)', title row",
     "Multiply by 1,000.", "Otherwise Brazil is 1,000x too small.", "Yes", VR),
    ("2.6", "Data shape", "A budget cell contains an Excel error (#REF!).", "05_Budget_Forecast, HQ01, Oct 2025, account 6300",
     "Treat as missing, not zero.", "A zero would create a fake favourable variance.", "Yes", VR),
    ("2.7", "Data shape", "Brazil's April ledger is missing (the bank data has April).", "BR01 ledger",
     "Report the gap; do not estimate it.", "Inventing a number hides a real data problem.", "Yes", VR),
    ("2.8", "Data shape", "Mixed currencies inside one entity: 5 HQ engineers paid in CAD; HQ has a EUR bank account.", "Headcount tab; cash tab",
     "Translate to USD; translate the EUR account at year-end rate before tying cash to the bank.", "Otherwise payroll and cash don't reconcile.", "Yes", VR),
    ("2.9", "Data shape", "Marketing account code changes from 6100 to 6110 in July.", "HQ01 ledger",
     "Map both codes to the same line.", "Otherwise marketing appears to drop to zero mid-year.", "Yes", VR),
    ("2.10", "Data shape", "A 'Period 13' audit adjustment of 85,000 USD (legal accrual).", "HQ01 ledger, period 2025-13",
     "Fold into December and log it.", "Otherwise the year-end total is wrong.", "Yes", VR),
    ("2.11", "Data shape", "A hidden duplicate cash tab.", "Hidden tab 'Copy of 02_CashBank_Position'",
     "Ignore it.", "Otherwise HQ cash is counted twice.", "Yes", VR),
    ("2.12", "Data shape", "Forecast entries dated Feb 2026 and one bookings row dated 15 Jan 2026.", "HQ01 ledger (notes 'FC'); bookings tab",
     "Exclude from FY2025 actuals.", "Forecast mixed into actuals overstates results.", "Yes", VR),
    ("D1", "Data shape", "Date styles differ by country: US MM/DD/YYYY, Europe DD/MM/YYYY and DD.MM.YYYY.", "Headcount and bookings tabs",
     "Detect the style per entity.", "03/04 could be 3 April or 4 March.", "Yes", VR),
    ("D2", "Data shape", "Account names in German, Portuguese and Spanish.", "DE01, BR01, MX01 ledgers",
     "Map to the standard lines.", "Otherwise lines land in the wrong place.", "Needs judgment", WB_ + " (Taxonomy Map)"),
    ("3.1", "Group structure", "Brazil is only 70% owned.", "Events / ownership info",
     "Include 100%, then show the 30% minority share of profit and net assets.", "Otherwise profit belonging to shareholders is overstated.", "Yes", WB_ + " (Ownership & Scope)"),
    ("3.2", "Group structure", "India is a 40% joint venture with profit & loss data only.", "JV01 rows",
     "Show one line: share of JV profit. Leave it out of revenue, headcount and cash. Adjust the 4.0m USD cost on HQ's books.", "Adding JV revenue would inflate the group's size.", "Yes", WB_ + " (Ownership & Scope)"),
    ("3.4", "Group structure", "Mexico was bought on 1 Jul 2025.", "Events tab",
     "Include Jul-Dec only. Calculate provisional goodwill; treat 650k deal costs as one-time; record the earn-out as a liability; show organic vs reported growth.", "Growth would otherwise look much better than it really is.", "Needs judgment", WB_ + " (Ownership & Scope); " + MEMO),
    ("3.5", "Group structure", "The UK business was sold on 30 Sep 2025.", "Events tab",
     "Show as a discontinued operation. The group gain is far smaller than the 3.0m the parent books.", "Overstated gain and misleading KPIs otherwise.", "Needs judgment", WB_ + " (Ownership & Scope); " + MEMO),
    ("3.6", "Group structure", "India's financial year starts in April.", "JV01 rows",
     "Align to calendar year and disclose.", "Otherwise the wrong 12 months are compared.", "Yes", VR),
    ("3.7", "Group structure", "A dormant Cayman holding company with a balance sheet only.", "HLD1 rows",
     "Include its balance sheet; do not expect headcount or budget.", "Avoids false 'missing data' alarms.", "Yes", VR),
    ("G1", "Group structure", "HQ carries 40.1m USD 'Investment in Subsidiaries'.", "HQ01 account 1500",
     "Eliminate against subsidiary equity.", "Otherwise group equity is overstated.", "Needs judgment", WB_ + " (FX & Consolidation)"),
    ("G2", "Group structure", "HQ sells goods to Germany and charges Brazil a management fee; December differs by 150,000 EUR of goods in transit.", "Intercompany tab",
     "Remove internal trading; eliminate the lower amount and disclose the difference.", "Otherwise revenue and costs are inflated.", "Needs judgment", WB_ + " (FX & Consolidation)"),
    ("4.1", "Currency", "Turkey is a hyperinflationary economy, but no price index is supplied.", "TR01; FX tab",
     "Ask you (LOW question); present Turkey with a 'not restated for hyperinflation' caveat.", "Turkish figures are misleading without the caveat.", "Yes", VR + " (question)"),
    ("4.2", "Currency", "Turkish lira drops 20% in March, then 2.5% a month.", "FX tab",
     "Show results at constant currency to separate FX from performance.", "Otherwise currency moves look like business decline.", "Yes", WB_ + " (Constant Currency)"),
    ("4.3", "Currency", "Mexico's August average exchange rate is missing.", "FX tab",
     "Use the August month-end rate as a proxy (MEDIUM) and disclose.", "Shows how gaps are filled transparently.", "Yes", WB_ + " (Assumptions log)"),
    ("4.5", "Currency", "Turkish cash is subject to capital controls.", "Cash tab, TR01 note",
     "Treat as restricted for group liquidity.", "That cash can't be used to repay group debt.", "Needs judgment", WB_ + " (Liquidity)"),
    ("5.1", "Liquidity & debt", "Loan covenant close to breach.", "08_Debt_Covenants",
     "Recompute: about 1.40x vs the 1.50x limit (about 7% headroom). The company reports 1.22x because it leaves the revolver out.", "A breach can make the loan repayable immediately.", "Needs judgment", WB_ + " (Liquidity); " + DECK + " slide 1"),
    ("5.2", "Liquidity & debt", "64.3m USD term loan due 30 Jun 2026, no refinancing in the file.", "08_Debt_Covenants",
     "Flag as a going-concern indicator.", "Due within 12 months with no plan to repay.", "Needs judgment", MEMO + "; " + DECK + " slide 1"),
    ("5.3", "Liquidity & debt", "5.0m USD revolver draw in November.", "Cash tab",
     "Show separately from operating cash flow.", "Borrowing can make cash generation look healthier than it is.", "Needs judgment", WB_ + " (Cash Flow)"),
    ("5.4", "Liquidity & debt", "Two going-concern indicators together (5.1 + 5.2).", "-",
     "Put them on the deck's first slide and in the memo; list facts, no conclusion.", "The board must see this first.", "Needs judgment", DECK + " slide 1; " + MEMO),
    ("5.5", "Liquidity & debt", "HQ stretches supplier payments from 45 to 80 days (Oct-Dec).", "HQ01 balance sheet",
     "Treat the cash benefit as one-off.", "Cash will reverse when suppliers are paid.", "Needs judgment", WB_ + "; " + MEMO),
    ("L1", "Liquidity & debt", "1.5m USD of HQ cash is held as collateral for a letter of credit.", "Cash tab",
     "Exclude from available cash and runway.", "That cash can't be spent.", "Needs judgment", WB_ + " (Liquidity)"),
    ("6.1", "Earnings quality", "One-time items: deal costs, disposal gain, Germany restructuring.", "Ledgers; events tab",
     "Build an Adjusted EBITDA bridge starting from the reported figure.", "Separates underlying performance from one-offs.", "Needs judgment", WB_ + " (Adjustments bridge)"),
    ("6.1b", "Earnings quality", "300k USD earn-out revaluation in December, not tagged as one-time.", "HQ01 account 7920",
     "Treat as one-time per policy and log it.", "Tests whether the review applies policy, not just tags.", "Needs judgment", WB_ + " (Assumptions log)"),
    ("6.2", "Earnings quality", "'One-time' restructuring of 120k USD in Feb, May, Aug and Nov (480k total).", "HQ01 account 7900",
     "Flag as recurring; show Adjusted EBITDA with and without it.", "A cost that happens every quarter is not one-time.", "Needs judgment", WB_ + "; " + MEMO + " (Earnings Quality)"),
    ("6.3", "Earnings quality", "December revenue pulled forward (bill-and-hold, +35%); customers paying slower (45 to 60 days).", "HQ01 revenue and receivables",
     "Flag it; never change reported revenue.", "Next year starts with a revenue hole.", "Needs judgment", WB_ + " (Red Flags); " + MEMO),
    ("6.4", "Earnings quality", "Capitalised development jumps from 150k to 400k USD a month in H2 while R&D staff stay flat.", "HQ01 ledger; headcount",
     "Show costs as reported and with capitalisation added back.", "Costs moved to the balance sheet flatter profit.", "Needs judgment", WB_ + "; " + MEMO + " (Earnings Quality)"),
    ("6.5", "Earnings quality", "Hidden 600k EUR warranty provision release in Germany in November (not disclosed).", "DE01 warranty expense",
     "Find it, quantify it, show margin with and without.", "It props up Germany's profit.", "Needs judgment", WB_ + "; " + MEMO + " (Earnings Quality)"),
    ("6.6", "Earnings quality", "A customer is a related party (Halcyon Family Trust Ltd).", "Bookings tab",
     "Disclose.", "Related-party sales may not be at arm's length.", "Needs judgment", MEMO),
    ("7a", "Revenue & customers", "HQ deferred revenue falls about 50% while service revenue looks flat.", "HQ01 balance sheet",
     "Flag as an early churn warning.", "Future revenue is drying up.", "Needs judgment", WB_ + " (Red Flags); " + MEMO),
    ("7b", "Revenue & customers", "Atlas Automotive is about 16% of bookings.", "Bookings tab",
     "Flag concentration above the 15% threshold.", "Losing one customer would hurt badly.", "Needs judgment", WB_ + " (Red Flags); " + DECK),
    ("7c", "Revenue & customers", "Germany loses Nordwerk GmbH, a 900k EUR customer, in August.", "Bookings tab",
     "Tell the story in the revenue bridge.", "Explains Germany's revenue drop.", "Needs judgment", WB_ + " (Revenue Drivers); " + MEMO),
    ("8a", "Balance sheet", "HQ inventory days rise from 70 to 120 in H2; HQ uses LIFO, others FIFO.", "HQ01 balance sheet",
     "Flag obsolescence risk; note inventory isn't comparable across entities.", "Possible write-down ahead.", "Needs judgment", WB_ + " (Red Flags)"),
    ("8b", "Balance sheet", "Turkey is loss-making with no tax benefit recorded.", "TR01",
     "Flag as an impairment trigger.", "Asset values may need writing down.", "Needs judgment", MEMO),
    ("1a", "Payroll", "Germany's ledger salaries are 1.21x the staff list (employer social charges).", "DE01 ledger vs headcount tab",
     "Ledger wins; disclose the explained difference.", "Shows payroll ties when properly explained.", "Needs judgment", VR + "; " + WB_),
    ("1b", "Payroll", "Turkey's staff list is a January snapshot; ledger salaries rise 4.5% a month.", "TR01 ledger vs headcount tab",
     "Ledger wins.", "Avoids a false payroll mismatch alarm.", "Needs judgment", WB_),
    ("9.1", "Budget", "HQ has an original budget and a July reforecast.", "05_Budget_Forecast",
     "Compare actuals with both.", "Shows whether the reforecast was realistic.", "Needs judgment", WB_ + " (Income Statement); " + DECK),
    ("9.2", "Budget", "Germany's budget is in USD at a fixed 1.05 rate.", "05_Budget_Forecast",
     "Compare at the budget rate and show FX separately.", "Separates FX from performance.", "Needs judgment", WB_),
    ("9.4a", "Budget", "HQ budget still uses the old marketing code 6100 all year.", "05_Budget_Forecast",
     "Map July-Dec 6110 actuals onto it.", "Otherwise a false variance.", "Needs judgment", WB_),
    ("9.4b", "Budget", "UK budget runs to December though the UK was sold in September.", "05_Budget_Forecast",
     "Compare Jan-Sep only.", "Otherwise a fake Oct-Dec shortfall.", "Needs judgment", WB_),
    ("9.4c", "Budget", "Turkey's budget is at January prices despite high inflation.", "05_Budget_Forecast",
     "Say variances are mostly inflation, not performance.", "Avoids misreading Turkey's results.", "Needs judgment", MEMO),
    ("9.4d", "Budget", "No budget for Mexico, the India JV, or Cayman.", "05_Budget_Forecast",
     "Compare like-for-like entities only.", "Avoids fake variances.", "Yes", VR),
    ("B1", "Budget", "HQ sales & marketing salaries were budgeted at only 85% of the staff list.", "05_Budget_Forecast; headcount tab",
     "Report a real overrun of about 18%.", "A genuine budgeting problem, not a data issue.", "Needs judgment", WB_ + "; " + DECK),
]

ws = wb.create_sheet("Issues Checklist")
title(ws, "Every planted issue, and what a correct review does",
      "Fill in the yellow 'Your check' column after a run. The Scorecard tab totals it.")
HDR = ["Ref", "Area", "What's planted in the test file", "Where in the input file", "What a correct review does",
       "Why it matters", "Caught automatically?", "Where to look in the outputs", "Your check", "Notes"]
last = table(ws, 4, HDR, [list(i) + ["Not checked", ""] for i in ISSUES], widths=[6, 16, 42, 28, 46, 32, 14, 30, 13, 28])
dv = DataValidation(type="list", formula1='"Found,Partly,Missed,Not checked"', allow_blank=False)
ws.add_data_validation(dv)
dv.add(f"I5:I{last}")
for r in range(5, last + 1):
    ws.cell(row=r, column=9).fill = FILL_INPUT
    ws.cell(row=r, column=9).font = F_INPUT
    ws.cell(row=r, column=10).fill = FILL_INPUT
rng = f"I5:I{last}"
ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Found"'], fill=PatternFill("solid", fgColor="C6EFCE")))
ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Partly"'], fill=PatternFill("solid", fgColor="FFEB9C")))
ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Missed"'], fill=PatternFill("solid", fgColor="FFC7CE")))
ws.freeze_panes = "C5"
ws.auto_filter.ref = f"A4:J{last}"
page_setup(ws)
ISSUE_LAST = last

# ---------------------------------------------------------------- Key Numbers
NUMS = [
    # (section, label, json path or literal, unit, meaning, tolerance)
    ("Debt & covenants", "Group Adjusted EBITDA, FY2025", "covenants.group_adjusted_ebitda_fy_usd", "USD", "Earnings before interest, tax, depreciation, amortisation and one-offs, as used in the covenant test.", "0.5%"),
    ("Debt & covenants", "Unrestricted cash, 31 Dec 2025", "covenants.unrestricted_cash_dec_usd", "USD", "Cash the group can actually use (excludes collateral and trapped Turkish cash).", "0.5%"),
    ("Debt & covenants", "Term loan outstanding", "term_loan_usd", "USD", "Main bank loan, due 30 Jun 2026.", "0.5%"),
    ("Debt & covenants", "Revolver drawn, 31 Dec 2025", "covenants.revolver_dec_usd", "USD", "Credit line drawn in November.", "0.5%"),
    ("Debt & covenants", "Net debt including revolver", "covenants.net_debt_incl_revolver_usd", "USD", "Total borrowing less usable cash.", "0.5%"),
    ("Debt & covenants", "Net leverage, recomputed", "covenants.net_leverage_recomputed", "x", "Net debt / Adjusted EBITDA. Must be at or below the limit.", "0.02"),
    ("Debt & covenants", "Net leverage limit (covenant)", "covenants.net_leverage_threshold", "x", "Maximum allowed by the loan agreement.", "0.02"),
    ("Debt & covenants", "Headroom to the limit", "covenants.net_leverage_headroom_pct", "%", "How far below the limit. Under 15% is a red flag.", "0.5"),
    ("Debt & covenants", "Net leverage as reported by the company", "covenants.net_leverage_company_reported_q4", "x", "Lower because the company leaves the revolver out of net debt. The review should flag the difference.", "0.02"),
    ("Debt & covenants", "Interest cover, recomputed", "covenants.interest_cover_recomputed", "x", "EBITDA / interest expense.", "0.05"),
    ("Mexico acquisition", "Total price paid", "mx01_acquisition.consideration_usd", "USD", "Cash plus earn-out.", "0.5%"),
    ("Mexico acquisition", "Cash paid", "mx01_acquisition.cash_paid_usd", "USD", "", "0.5%"),
    ("Mexico acquisition", "Earn-out (fair value at deal)", "mx01_acquisition.earnout_fair_value_usd", "USD", "Future payment if targets are met; a liability.", "0.5%"),
    ("Mexico acquisition", "Net assets acquired (at 30 Jun rate)", "mx01_acquisition.net_assets_at_acquisition_usd_jun_spot", "USD", "Book value of what was bought.", "0.5%"),
    ("Mexico acquisition", "Provisional goodwill", "mx01_acquisition.provisional_goodwill_usd", "USD", "Price paid above net assets.", "0.5%"),
    ("Mexico acquisition", "Deal costs (one-time)", "mx01_acquisition.deal_costs_usd", "USD", "Expensed, added back in Adjusted EBITDA.", "0.5%"),
    ("Mexico acquisition", "Earn-out revaluation in December", "mx01_acquisition.earnout_fv_change_dec_usd", "USD", "Treated as one-time.", "0.5%"),
    ("Mexico acquisition", "Mexico net income since acquisition (Jul-Dec)", "mx01_acquisition.fy_net_income_post_acq_usd", "USD", "", "0.5%"),
    ("UK disposal", "Sale proceeds", "uk01_disposal.proceeds_usd", "USD", "", "0.5%"),
    ("UK disposal", "Gain on HQ's own books", "uk01_disposal.parent_book_gain_usd", "USD", "Proceeds less the original investment cost. NOT the group figure.", "0.5%"),
    ("UK disposal", "UK net assets at 30 Sep", "uk01_disposal.net_assets_30sep_usd", "USD", "", "0.5%"),
    ("UK disposal", "Group gain on disposal (before currency recycling)", "uk01_disposal.consolidated_gain_before_cta_recycling_usd", "USD", "Proceeds less UK net assets. Far smaller than the HQ book gain.", "0.5%"),
    ("UK disposal", "UK net income Jan-Sep", "uk01_disposal.jan_sep_net_income_usd", "USD", "Shown under discontinued operations.", "0.5%"),
    ("Brazil (70% owned)", "Brazil net income, full year", "br01_nci.fy_net_income_usd_all_12_months", "USD", "Includes April, which is missing from the ledger export.", "0.5%"),
    ("Brazil (70% owned)", "Minority (30%) share of Brazil profit", "br01_nci.nci_share_30pct_usd_all_12_months", "USD", "Belongs to the other shareholders.", "0.5%"),
    ("Brazil (70% owned)", "Brazil April net income missing from the file", "br01_nci.april_net_income_missing_from_export_usd", "USD", "The review should report the gap, not fill it.", "0.5%"),
    ("India JV (40%)", "Group share of JV profit", "jv01_equity_method.share_of_profit_usd", "USD", "One line in the P&L; no JV revenue in group revenue.", "0.5%"),
    ("HQ items", "Recurring 'one-time' restructuring (4 x 120k)", "hq01.recurring_restructuring_usd", "USD", "", "0.5%"),
    ("HQ items", "December revenue pulled forward", "hq01.dec_bill_and_hold_uplift_usd", "USD", "Flag only; never deducted from revenue.", "0.5%"),
    ("HQ items", "Period 13 audit adjustment", "hq01.period13_adjustment_usd", "USD", "Folded into December.", "0.5%"),
    ("HQ items", "Investment in subsidiaries (eliminated)", "hq01.parent_investment_in_subsidiaries_dec_usd", "USD", "", "0.5%"),
    ("HQ items", "Investment in JV at cost", "hq01.investment_in_jv_at_cost_usd", "USD", "Replaced by equity-method value.", "0.5%"),
    ("Germany items", "Hidden warranty provision release (Nov)", "de01.warranty_provision_release_nov_eur", "EUR", "", "0.5%"),
    ("Germany items", "Intercompany goods in transit (Dec)", "de01.dec_ic_in_transit_eur", "EUR", "", "0.5%"),
    ("Turkey items", "Turkey net income, FY2025", "tr01.fy_net_income_try", "TRY", "Loss.", "0.5%"),
    ("Customers", "Atlas Automotive share of gross bookings", "concentration.top_customer_share_of_gross_bookings_approx", "share", "Above the 15% concentration threshold.", "0.01"),
]
ws = wb.create_sheet("Key Numbers")
title(ws, "Figures a correct review should reproduce",
      "Type the review's figure into the yellow column. Within tolerance = PASS. Small rounding differences are normal.")
hdr = ["Area", "Figure", "Expected", "Unit", "Your run's figure", "Difference", "Result", "Tolerance", "What it means"]
rows = []
for sec, label, path, unit, meaning, tol in NUMS:
    rows.append([sec, label, get(path), unit, None, None, None, tol, meaning])
last = table(ws, 4, hdr, rows, widths=[20, 44, 16, 7, 18, 14, 10, 10, 60])
for r in range(5, last + 1):
    unit = ws.cell(row=r, column=4).value
    tol = ws.cell(row=r, column=8).value
    fmt = {"x": "0.00\"x\"", "%": "0.0\"%\"", "share": "0%"}.get(unit, "#,##0;(#,##0)")
    for c in (3, 5, 6):
        ws.cell(row=r, column=c).number_format = fmt
    e = ws.cell(row=r, column=5)
    e.fill, e.font = FILL_INPUT, F_INPUT
    ws.cell(row=r, column=6).value = f'=IF(E{r}="","",E{r}-C{r})'
    if tol.endswith("%"):
        limit = f"ABS(C{r})*{float(tol[:-1]) / 100}"
    else:
        limit = tol
    ws.cell(row=r, column=7).value = f'=IF(E{r}="","",IF(ABS(E{r}-C{r})<={limit},"PASS","CHECK"))'
rng = f"G5:G{last}"
ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"PASS"'], fill=PatternFill("solid", fgColor="C6EFCE")))
ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"CHECK"'], fill=PatternFill("solid", fgColor="FFC7CE")))
ws.freeze_panes = "C5"
ws.cell(row=last + 2, column=1, value="Tolerance: a % is relative to the expected figure; a plain number is an absolute difference. "
        "Currency amounts use the test file's own exchange rates.").font = F_SUB
page_setup(ws)
NUM_LAST = last

# ---------------------------------------------------------------- Scorecard
ws = wb.create_sheet("Scorecard")
title(ws, "Scorecard", "Calculated from your entries on 'Issues Checklist' and 'Key Numbers'.")
ic = f"'Issues Checklist'!I5:I{ISSUE_LAST}"
ia = f"'Issues Checklist'!G5:G{ISSUE_LAST}"
kn = f"'Key Numbers'!G5:G{NUM_LAST}"
table(ws, 4, ["Planted issues", "Count"], [
    ["Total planted issues", f"=COUNTA({ic})"],
    ["Found", f'=COUNTIF({ic},"Found")'],
    ["Partly", f'=COUNTIF({ic},"Partly")'],
    ["Missed", f'=COUNTIF({ic},"Missed")'],
    ["Not checked yet", f'=COUNTIF({ic},"Not checked")'],
    ["Score (Found = 1, Partly = 0.5) as % of checked", '=IFERROR((B6+0.5*B7)/(B6+B7+B8),"")'],
    ["Auto-detectable issues missed (should be 0)", f'=COUNTIFS({ia},"Yes",{ic},"Missed")'],
], widths=[52, 14])
ws["B10"].number_format = "0%"
table(ws, 13, ["Key numbers", "Count"], [
    ["Figures to check", f"=ROWS({kn})"],
    ["PASS", f'=COUNTIF({kn},"PASS")'],
    ["CHECK (outside tolerance)", f'=COUNTIF({kn},"CHECK")'],
    ["Not entered yet", f'=COUNTBLANK({kn})'],
])
table(ws, 19, ["The four must-haves for a good run", "Met?"], [
    ["1. Every 'Caught automatically = Yes' issue appears in the Data Validation Report", "Not checked"],
    ["2. The Turkey hyperinflation index is raised as a LOW-confidence question at the checkpoint", "Not checked"],
    ["3. Key Numbers match within tolerance", "Not checked"],
    ["4. Deck slide 1 leads with the covenant near-breach and the June 2026 loan maturity", "Not checked"],
])
dv2 = DataValidation(type="list", formula1='"Yes,No,Not checked"', allow_blank=False)
ws.add_data_validation(dv2)
dv2.add("B20:B23")
for r in range(20, 24):
    ws.cell(row=r, column=2).fill, ws.cell(row=r, column=2).font = FILL_INPUT, F_INPUT
ws.conditional_formatting.add("B20:B23", CellIsRule(operator="equal", formula=['"Yes"'], fill=PatternFill("solid", fgColor="C6EFCE")))
ws.conditional_formatting.add("B20:B23", CellIsRule(operator="equal", formula=['"No"'], fill=PatternFill("solid", fgColor="FFC7CE")))
ws["A25"] = "Overall verdict"
ws["A25"].font = F_B
ws["B25"] = ('=IF(COUNTIF(B20:B23,"Not checked")>0,"Incomplete",'
             'IF(AND(COUNTIF(B20:B23,"Yes")=4,B11=0),"GOOD RUN","NEEDS WORK"))')
ws["B25"].font = F_B
ws.column_dimensions["B"].width = 16
page_setup(ws, landscape=False)

os.makedirs(OUT_DIR, exist_ok=True)
wb.save(OUT)
print("Wrote", OUT, "-", len(ISSUES), "issues,", len(NUMS), "figures")
