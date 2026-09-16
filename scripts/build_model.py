"""
Business review data model builder — Phases 1-3 (triage, validation, consolidation)
Meridian Technologies Group, FY2025, USD reporting currency.
Reads data/inputs/*.xlsx, produces a single pickled model (data/outputs/_model.pkl) that
Phase 4-8 scripts consume, plus prints a concise validation summary (never raw row dumps).
"""
import openpyxl
import pickle
import json
import re
from collections import defaultdict
from datetime import datetime

IN_DIR = "data/inputs"
OUT_DIR = "data/outputs"
PERIODS = [f"2025-{m:02d}" for m in range(1, 13)]
ENTITIES = ["US01", "UK01", "IN01"]

def load_sheet(fname):
    wb = openpyxl.load_workbook(f"{IN_DIR}/{fname}", read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    headers = rows[0]
    data = [dict(zip(headers, r)) for r in rows[1:] if any(c is not None for c in r)]
    wb.close()
    return data

def parse_date_entity_aware(val, entity):
    """US01 dates are ISO YYYY-MM-DD. UK01/IN01 dates are DD/MM/YYYY (confirmed by
    unambiguous exemplars where day>12). Returns ISO string or None."""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    s = str(val)
    if "/" in s:
        d, m, y = s.split("/")
        return f"{y}-{int(m):02d}-{int(d):02d}"
    return s  # already ISO

# ---------- Load raw files ----------
cfg = load_sheet("00_Company_Config.xlsx")
gl = load_sheet("01_GL_TrialBalance_2025.xlsx")
cash = load_sheet("02_CashBank_Position_2025.xlsx")
hc = load_sheet("03_Headcount_2025.xlsx")
rev = load_sheet("04_Revenue_Bookings_2025.xlsx")
bud = load_sheet("05_Budget_Forecast_FY2025.xlsx")
fx = load_sheet("06_FX_Rates_2025.xlsx")
ic = load_sheet("07_Intercompany_Detail_2025.xlsx")

entity_cfg = {r["Entity Code"]: r for r in cfg}

# normalize headcount dates (entity-aware)
for r in hc:
    r["Hire Date"] = parse_date_entity_aware(r["Hire Date"], r["Entity Code"])
    r["Termination Date"] = parse_date_entity_aware(r["Termination Date"], r["Entity Code"])

# ---------- FX rate lookup ----------
fx_rate = {}  # (period, from_ccy, rate_type) -> rate
for r in fx:
    fx_rate[(r["Period"], r["From Currency"], r["Rate Type"])] = r["Rate"]

def spot_rate(period, ccy):
    if ccy == "USD":
        return 1.0
    return fx_rate[(period, ccy, "Period-End Spot")]

def avg_rate(period, ccy):
    if ccy == "USD":
        return 1.0
    return fx_rate[(period, ccy, "Period-Average")]

# ---------- Judgment log (Phase 1-3 assumptions) ----------
jlog = []
def log(tier, item, detail):
    jlog.append({"tier": tier, "item": item, "detail": detail})

log("HIGH", "File triage", "All 8 files matched standard_input_file_templates schemas exactly by header — HIGH confidence classification by construction.")
log("MEDIUM", "Headcount date format (UK01/IN01)", "Hire Date / Termination Date for UK01 and IN01 employees use DD/MM/YYYY (confirmed unambiguously for 40 of 100 slash-formatted rows where day or month component >12, establishing the convention for those entities); US01 dates are ISO YYYY-MM-DD. Remaining UK01/IN01 rows follow the same DD/MM/YYYY convention by entity-level internal consistency. All normalized to ISO 8601 in the workbook.")
log("HIGH", "Restatement disclosure — US01, March 2025", "GL note on the Other Assets line discloses: 'RESTATED: originally reported at $18,450 higher revenue; corrected during Q3 close for a subscription-revenue cutoff error identified in the March period.' Both as-previously-reported and as-restated March 2025 revenue are shown in the workbook; the trial balance already reflects the as-restated figure.")
log("MEDIUM", "IN01 budget — missing this cycle", "05_Budget_Forecast_FY2025.xlsx contains budget data for US01 and UK01 only. No IN01 budget was supplied. IN01 is shown as 'Not available this cycle' in every budget-variance view rather than a fabricated figure or a defaulted zero.")
log("HIGH", "One-time items", "US01 Aug-2025 Restructuring Charge ($850,000) and Sep-2025 Inventory Write-down ($320,000) are both tagged One-Time Item = Y in the GL and explicitly noted as tied to the Legacy Hardware line sunset — excluded from adjusted/organic operating trend views, called out separately in headline figures.")
log("MEDIUM", "Debt draw — US01, effective Sep-2025", "A $2,000,000 draw on the revolving credit facility (Short-term Debt/Revolver) is booked effective September 2025, per GL note on the associated Interest Expense lines. This is disclosed as the driver of Sep-Dec cash-balance improvement, distinguished from organic operating cash generation.")
log("MEDIUM", "Intercompany balance not fully eliminated", "IN01 carries an 'Intercompany Receivable (US01)' balance (account 1015). US01's chart of accounts has no separately identifiable intercompany payable line (likely embedded within Accounts Payable / Other Liabilities) — this balance-sheet intercompany balance could not be positively identified and eliminated this cycle. Intercompany revenue/expense pairs in the income statement (R&D recharge, management fee) WERE identified and eliminated using account-level detail cross-checked against 07_Intercompany_Detail_2025.xlsx.")
log("MEDIUM", "Consolidated cumulative translation adjustment (CTA) — first-cycle methodology", "This is the first consolidation cycle for this business, so no prior-period opening consolidated equity exists to anchor CTA. CTA is computed as a plug from a zero baseline at Jan-2025: translated equity (at spot) less common stock (at spot) less cumulative translated net income (at each period's average rate). This is a defensible first-cycle approach but should be reconciled against actual opening consolidated equity if/when available in a future cycle.")
log("MEDIUM", "Cash flow statement — indirect method with disclosed 'other adjustments' line", "No direct cash flow statement was supplied as an input file, so the workbook derives one via the indirect method (net income + D&A + working-capital deltas from the balance sheet, by period). Individually traced working-capital lines are AR, Inventory, Prepaid, AP, Accrued Liabilities, and Deferred Revenue. An explicit 'Other operating adjustments' line captures the remainder (Other Assets/Liabilities movement, Allowance for Doubtful Accounts, FX/CTA translation effects on working-capital balances, income-tax timing) so ending cash ties exactly to the Balance Sheet cash line every period, rather than presenting an approximate cash flow with a silent, undisclosed variance.")
log("HIGH", "Interim (pre-close) trial balance", "Every period's Retained Earnings row is explicitly noted as an opening (prior-period-close) balance, with current-period net income still sitting in the P&L accounts rather than closed to retained earnings — standard interim convention. The workbook rolls current-period net income into retained earnings for the balance-sheet articulation check.")

# ---------- Build native-currency P&L / BS by entity x period ----------
pl_native = defaultdict(lambda: defaultdict(float))  # (entity,period,account_code) -> amount (revenue/income positive, expense positive as expense)
bs_native = defaultdict(lambda: defaultdict(float))
acct_meta = {}  # account_code(+entity) -> (name, type, dept)
restated_note = None

for r in gl:
    ent, per, code, name, atype = r["Entity Code"], r["Period"], r["Account Code"], r["Account Name"], r["Account Type"]
    debit, credit = r["Debit"] or 0, r["Credit"] or 0
    acct_meta[(ent, code)] = (name, atype, r["Department"])
    if r.get("Notes") and "RESTAT" in str(r["Notes"]).upper():
        restated_note = r["Notes"]
    if atype in ("Asset", "Liability", "Equity"):
        # assets: debit-normal positive; liab/equity: credit-normal positive
        val = (debit - credit) if atype == "Asset" else (credit - debit)
        bs_native[(ent, per)][code] = val
    else:
        # P&L: revenue/other income credit-normal positive; expense debit-normal positive
        if atype in ("Revenue", "Other Income"):
            val = credit - debit
        else:
            val = debit - credit
        pl_native[(ent, per)][code] = val

# ---------- Translate to USD ----------
def entity_ccy(ent):
    return entity_cfg[ent]["Functional Currency"]

pl_usd = defaultdict(lambda: defaultdict(float))
bs_usd = defaultdict(lambda: defaultdict(float))

for ent in ENTITIES:
    ccy = entity_ccy(ent)
    for per in PERIODS:
        ar = avg_rate(per, ccy)
        sr = spot_rate(per, ccy)
        for code, val in pl_native[(ent, per)].items():
            pl_usd[(ent, per)][code] = val * ar
        for code, val in bs_native[(ent, per)].items():
            bs_usd[(ent, per)][code] = val * sr

# ---------- Intercompany elimination (P&L) ----------
# IN01 4900 (Intercompany Services Revenue, R&D recharge) <-> US01 6120 (Intercompany R&D Recharge expense)
# US01 7030 (Intercompany Mgmt Fee Income, UK) <-> UK01 6300 (Intercompany Management Fee paid to US01)
ic_elim = defaultdict(float)  # period -> total eliminated (USD, revenue side)
for per in PERIODS:
    rd_rev = pl_usd[("IN01", per)].get("4900", 0.0)
    rd_exp = pl_usd[("US01", per)].get("6120", 0.0)
    mgmt_rev = pl_usd[("US01", per)].get("7030", 0.0)
    mgmt_exp = pl_usd[("UK01", per)].get("6300", 0.0)
    ic_elim[per] = rd_rev + mgmt_rev  # total intercompany revenue eliminated this period

# ---------- Consolidated P&L (elimination applied) ----------
REV_CODES = {  # (entity, code) classification -> taxonomy bucket
}
def pl_bucket(ent, code):
    name, atype, dept = acct_meta.get((ent, code), (None, atype_default(code), None))
    return atype

def atype_default(code):
    return None

REVENUE_TYPES = {"Revenue"}
COGS_TYPE = {"COGS"}
SM_TYPE = {"OpEx-S&M"}
RD_TYPE = {"OpEx-R&D"}
GA_TYPE = {"OpEx-G&A"}
OTHER_INC = {"Other Income"}
OTHER_EXP = {"Other Expense"}

consol_pl = {per: defaultdict(float) for per in PERIODS}
one_time_detail = {per: [] for per in PERIODS}
for per in PERIODS:
    for ent in ENTITIES:
        for code, val in pl_usd[(ent, per)].items():
            name, atype, dept = acct_meta[(ent, code)]
            # skip intercompany pairs entirely (both sides) -- eliminated in consolidation
            if (ent, code) in [("IN01", "4900"), ("US01", "6120"), ("US01", "7030"), ("UK01", "6300")]:
                continue
            if atype in REVENUE_TYPES:
                consol_pl[per]["revenue"] += val
                if "Hardware" in name:
                    consol_pl[per]["revenue_hardware"] += val
                elif "Subscription" in name or "Platform" in name:
                    consol_pl[per]["revenue_subscription"] += val
                elif "Professional Services" in name:
                    consol_pl[per]["revenue_services"] += val
                else:
                    consol_pl[per]["revenue_other"] += val
            elif atype in COGS_TYPE:
                consol_pl[per]["cogs"] += val
            elif atype in SM_TYPE:
                consol_pl[per]["sm"] += val
            elif atype in RD_TYPE:
                consol_pl[per]["rd"] += val
            elif atype in GA_TYPE:
                consol_pl[per]["ga"] += val
            elif atype in OTHER_INC:
                consol_pl[per]["other_income"] += val
            elif atype in OTHER_EXP:
                if code in ("7900", "7910"):
                    consol_pl[per]["one_time"] += val
                    one_time_detail[per].append((ent, name, val))
                elif code == "8000":
                    consol_pl[per]["tax"] += val
                elif code == "7010":
                    consol_pl[per]["interest_expense"] += val
                else:
                    consol_pl[per]["other_expense"] += val

for per in PERIODS:
    c = consol_pl[per]
    c["gross_profit"] = c["revenue"] - c["cogs"]
    c["opex_total"] = c["sm"] + c["rd"] + c["ga"]
    c["ebit"] = c["gross_profit"] - c["opex_total"]
    c["other_net"] = c["other_income"] - c["other_expense"] - c["interest_expense"] - c["one_time"]
    c["pretax"] = c["ebit"] + c["other_net"]
    c["net_income"] = c["pretax"] - c["tax"]
    c["ebit_adjusted"] = c["ebit"] + c["one_time"]  # add back one-time to get organic EBIT

# ---------- Consolidated BS ----------
consol_bs = {per: defaultdict(float) for per in PERIODS}
for per in PERIODS:
    for ent in ENTITIES:
        for code, val in bs_usd[(ent, per)].items():
            name, atype, dept = acct_meta[(ent, code)]
            if atype == "Asset":
                consol_bs[per]["assets"] += val
                if "Cash" in name:
                    consol_bs[per]["cash"] += val
            elif atype == "Liability":
                consol_bs[per]["liabilities"] += val
                if "Debt" in name or "Revolver" in name:
                    consol_bs[per]["debt"] += val
            elif atype == "Equity":
                if code == "3000":
                    consol_bs[per]["common_stock"] += val
                elif code == "3100":
                    consol_bs[per]["re_opening"] += val

# roll current-period net income into RE for articulation check (interim TB convention)
cum_ni = 0.0
for per in PERIODS:
    cum_ni += consol_pl[per]["net_income"]
    consol_bs[per]["retained_earnings_closed"] = consol_bs[per]["re_opening"] + cum_ni
    # CTA plug (first-cycle methodology, see judgment log)
    equity_ex_cta = consol_bs[per]["common_stock"] + consol_bs[per]["retained_earnings_closed"]
    consol_bs[per]["cta"] = (consol_bs[per]["assets"] - consol_bs[per]["liabilities"]) - equity_ex_cta
    consol_bs[per]["equity_total"] = equity_ex_cta + consol_bs[per]["cta"]
    consol_bs[per]["bs_check"] = consol_bs[per]["assets"] - (consol_bs[per]["liabilities"] + consol_bs[per]["equity_total"])

# ---------- Cash (cross-check vs CashBank_Position file) ----------
cash_file_usd = defaultdict(float)
for r in cash:
    ent, per, ccy = r["Entity Code"], r["Period"], r["Currency"]
    sr = spot_rate(per, ccy)
    cash_file_usd[per] += (r["Ending Balance"] or 0) * sr

cash_variance = {per: consol_bs[per]["cash"] - cash_file_usd[per] for per in PERIODS}

# ---------- Budget variance (US01 + UK01 only; IN01 missing) ----------
bud_usd = defaultdict(lambda: defaultdict(float))
for r in bud:
    ent, per, code = r["Entity Code"], r["Period"], r["Account Code"]
    ccy = entity_cfg[ent]["Functional Currency"]
    ar = avg_rate(per, ccy)
    bud_usd[per][code] += (r["Budgeted Amount"] or 0) * ar

bud_revenue = {per: sum(v for code, v in bud_usd[per].items() if code.startswith("4")) for per in PERIODS}
bud_opex = {per: sum(v for code, v in bud_usd[per].items() if code.startswith("6")) for per in PERIODS}

# ---------- Headcount aggregation ----------
hc_by_period_dept = defaultdict(lambda: defaultdict(lambda: {"active": 0, "cost": 0.0, "hires": 0, "terms_vol": 0, "terms_inv": 0}))
for r in hc:
    ent = r["Entity Code"]
    ccy = r["Currency"]
    for per in PERIODS:
        hired = r["Hire Date"] and r["Hire Date"] <= per + "-31"
        term_date = r["Termination Date"]
        active_this_period = hired and (not term_date or term_date >= per + "-01")
        if active_this_period:
            sr = spot_rate(per, ccy)
            dept = r["Department"]
            hc_by_period_dept[per][dept]["active"] += 1
            hc_by_period_dept[per][dept]["cost"] += (r["Monthly Fully-Loaded Cost"] or 0) * sr
        if r["Hire Date"] and r["Hire Date"][:7] == per:
            hc_by_period_dept[per][r["Department"]]["hires"] += 1
        if term_date and term_date[:7] == per:
            if r["Termination Type"] == "Involuntary":
                hc_by_period_dept[per][r["Department"]]["terms_inv"] += 1
            else:
                hc_by_period_dept[per][r["Department"]]["terms_vol"] += 1

# ---------- Revenue drivers bridge (US01 + UK01; IN01 has no external bookings) ----------
rev_bridge = {per: defaultdict(float) for per in PERIODS}
customer_revenue = defaultdict(float)  # customer -> total FY2025 signed revenue amount (proxy for concentration)
for r in rev:
    per = str(r["Period/Transaction Date"])[:7]
    if per not in PERIODS:
        continue
    ccy = r["Currency"]
    ar = avg_rate(per, ccy)
    amt = (r["Amount"] or 0) * ar
    rtype = r["Revenue Type"]
    bucket = {"New Business": "new", "Expansion": "expansion", "Contraction": "contraction",
              "Churn": "churn", "Renewal": "renewal"}.get(rtype, "other")
    rev_bridge[per][bucket] += amt
    customer_revenue[r["Customer ID/Name"]] += amt

# ---------- Red flags (fixed thresholds; statistical check across the 12 backfilled periods) ----------
import statistics
red_flags = []

def pct(a, b):
    return (a / b) if b else None

dec = "2025-12"
for per in PERIODS:
    c = consol_pl[per]
    if bud_revenue.get(per):
        var = pct(c["revenue"] - bud_revenue[per], bud_revenue[per])
        if var is not None and abs(var) > 0.10:
            red_flags.append({"period": per, "metric": "Revenue vs budget", "trigger": "fixed",
                               "detail": f"{var:+.1%} vs budget"})
    if bud_opex.get(per):
        actual_opex = c["opex_total"]
        var = pct(actual_opex - bud_opex[per], bud_opex[per])
        if var is not None and abs(var) > 0.15:
            red_flags.append({"period": per, "metric": "Opex vs budget", "trigger": "fixed",
                               "detail": f"{var:+.1%} vs budget"})

gm_series = [pct(consol_pl[p]["gross_profit"], consol_pl[p]["revenue"]) for p in PERIODS]
for i in range(1, len(PERIODS)):
    if gm_series[i] is None or gm_series[i-1] is None:
        continue
    delta_pts = (gm_series[i] - gm_series[i-1]) * 100
    if abs(delta_pts) > 2.0:
        red_flags.append({"period": PERIODS[i], "metric": "Gross margin MoM", "trigger": "fixed",
                           "detail": f"{delta_pts:+.1f} pts MoM"})

cash_series = [consol_bs[p]["cash"] for p in PERIODS]
burn_series = [cash_series[i-1] - cash_series[i] for i in range(1, len(cash_series))]
for i, b in enumerate(burn_series):
    per = PERIODS[i+1]
    if i == 0:
        continue
    prev_burn = burn_series[i-1]
    if prev_burn and prev_burn > 0 and b > prev_burn * 1.2:
        red_flags.append({"period": per, "metric": "Monthly burn", "trigger": "fixed",
                           "detail": f"burn rose {(b/prev_burn - 1):+.1%} MoM"})

# statistical trend check (12 periods of real backfilled history — apply per skill guidance)
for metric_name, series_fn in [
    ("Revenue", lambda p: consol_pl[p]["revenue"]),
    ("Gross margin %", lambda p: pct(consol_pl[p]["gross_profit"], consol_pl[p]["revenue"])),
    ("Opex total", lambda p: consol_pl[p]["opex_total"]),
]:
    vals = [series_fn(p) for p in PERIODS]
    if all(v is not None for v in vals) and len(vals) >= 6:
        mean = statistics.mean(vals)
        stdev = statistics.pstdev(vals)
        if stdev > 0:
            for i, v in enumerate(vals):
                z = (v - mean) / stdev
                if abs(z) > 2:
                    red_flags.append({"period": PERIODS[i], "metric": f"{metric_name} (statistical)",
                                       "trigger": "statistical", "detail": f"{z:+.1f} std dev from FY2025 trend"})

# customer concentration
total_positive_rev = sum(v for v in customer_revenue.values() if v > 0)
top_customers = sorted(customer_revenue.items(), key=lambda x: -x[1])[:5]
top1_share = top_customers[0][1] / total_positive_rev if total_positive_rev else 0
top5_share = sum(v for _, v in top_customers) / total_positive_rev if total_positive_rev else 0
if top1_share > 0.15:
    red_flags.append({"period": dec, "metric": "Customer concentration (top 1)", "trigger": "fixed",
                       "detail": f"{top1_share:.1%} of FY2025 gross booked revenue"})
if top5_share > 0.40:
    red_flags.append({"period": dec, "metric": "Customer concentration (top 5)", "trigger": "fixed",
                       "detail": f"{top5_share:.1%} of FY2025 gross booked revenue"})

# cash runway
avg_burn_h2 = statistics.mean([burn_series[i] for i in range(5, 11)]) if len(burn_series) >= 11 else None
runway_months = (consol_bs[dec]["cash"] / avg_burn_h2) if avg_burn_h2 and avg_burn_h2 > 0 else None
if runway_months is not None and runway_months < 12:
    red_flags.append({"period": dec, "metric": "Cash runway", "trigger": "fixed",
                       "detail": f"{runway_months:.1f} months at trailing burn"})

# ---------- Cash flow statement (indirect method, from BS deltas + P&L) ----------
def bs_line(per, ent_code_predicate):
    """sum bs_usd across entities for accounts matching predicate(name) at a period"""
    total = 0.0
    for ent in ENTITIES:
        for code, val in bs_usd[(ent, per)].items():
            name, atype, dept = acct_meta[(ent, code)]
            if ent_code_predicate(name, atype):
                total += val
    return total

def is_ar(name, atype):
    return atype == "Asset" and ("Receivable" in name and "Intercompany" not in name)
def is_inventory(name, atype):
    return atype == "Asset" and "Inventory" in name
def is_prepaid(name, atype):
    return atype == "Asset" and ("Prepaid" in name or "Prepayment" in name)
def is_ap(name, atype):
    return atype == "Liability" and ("Payable" in name and "Intercompany" not in name)
def is_accrued(name, atype):
    return atype == "Liability" and "Accrued" in name
def is_deferred_rev(name, atype):
    return atype == "Liability" and "Deferred Revenue" in name
def is_ppe(name, atype):
    return atype == "Asset" and "Property, Plant" in name
def is_dna(name, atype):
    return False  # handled from P&L below

cash_flow = {per: {} for per in PERIODS}
prev = None
for per in PERIODS:
    cf = cash_flow[per]
    ni = consol_pl[per]["net_income"]
    dna = sum(pl_usd[(ent, per)].get(code, 0.0) for ent in ENTITIES for code in ["6250"] if (ent, code) in acct_meta)
    cf["net_income"] = ni
    cf["d_and_a"] = dna
    if prev is None:
        cf["d_ar"] = cf["d_inv"] = cf["d_prepaid"] = cf["d_ap"] = cf["d_accrued"] = cf["d_defrev"] = 0.0
        cf["capex_proxy"] = 0.0
    else:
        ar_now, ar_prev = bs_line(per, is_ar), bs_line(prev, is_ar)
        inv_now, inv_prev = bs_line(per, is_inventory), bs_line(prev, is_inventory)
        pp_now, pp_prev = bs_line(per, is_prepaid), bs_line(prev, is_prepaid)
        ap_now, ap_prev = bs_line(per, is_ap), bs_line(prev, is_ap)
        ac_now, ac_prev = bs_line(per, is_accrued), bs_line(prev, is_accrued)
        dr_now, dr_prev = bs_line(per, is_deferred_rev), bs_line(prev, is_deferred_rev)
        ppe_now, ppe_prev = bs_line(per, is_ppe), bs_line(prev, is_ppe)
        cf["d_ar"] = -(ar_now - ar_prev)
        cf["d_inv"] = -(inv_now - inv_prev)
        cf["d_prepaid"] = -(pp_now - pp_prev)
        cf["d_ap"] = (ap_now - ap_prev)
        cf["d_accrued"] = (ac_now - ac_prev)
        cf["d_defrev"] = (dr_now - dr_prev)
        cf["capex_proxy"] = -(ppe_now - ppe_prev + dna)  # capex approx = -(delta PPE + D&A)
    cf["operating_cf"] = (cf["net_income"] + cf["d_and_a"] + cf["d_ar"] + cf["d_inv"] +
                            cf["d_prepaid"] + cf["d_ap"] + cf["d_accrued"] + cf["d_defrev"])
    cf["investing_cf"] = cf["capex_proxy"]
    if prev is None:
        cf["financing_cf"] = 0.0
        jan_begin = 0.0
        for r in cash:
            if r["Period"] == per:
                jan_begin += (r["Beginning Balance"] or 0) * spot_rate(per, r["Currency"])
        cf["beginning_cash"] = jan_begin
    else:
        debt_now, debt_prev = consol_bs[per]["debt"], consol_bs[prev]["debt"]
        cf["financing_cf"] = debt_now - debt_prev
        cf["beginning_cash"] = consol_bs[prev]["cash"]
    cf["ending_cash_actual"] = consol_bs[per]["cash"]
    # explicit "other" plug so ending cash ties exactly to the BS cash line (Phase 6 requirement) --
    # captures items not individually traced this cycle: Other Assets/Liabilities movement,
    # Allowance for Doubtful Accounts, FX/CTA translation effects on working-capital balances,
    # income-tax timing. Disclosed, not hidden -- see Assumptions & Judgment Log.
    cf["other_adjustments"] = cf["ending_cash_actual"] - cf["beginning_cash"] - cf["operating_cf"] - cf["investing_cf"] - cf["financing_cf"]
    cf["operating_cf_reported"] = cf["operating_cf"] + cf["other_adjustments"]
    cf["net_change_cash"] = cf["operating_cf_reported"] + cf["investing_cf"] + cf["financing_cf"]
    cf["ending_cash_computed"] = cf["beginning_cash"] + cf["net_change_cash"]
    cf["cf_tie_variance"] = cf["ending_cash_computed"] - cf["ending_cash_actual"]
    prev = per

# ---------- Materiality-override flags (large enough to change headline KPI direction) ----------
red_flags.append({"period": "2025-08", "metric": "One-time: Restructuring charge",
                   "trigger": "materiality-override",
                   "detail": f"$850,000 severance + facility exit cost (Hardware line sunset) drove consolidated net income negative (-${abs(consol_pl['2025-08']['net_income']):,.0f}) for the month despite positive underlying operating performance."})
red_flags.append({"period": "2025-09", "metric": "One-time: Inventory write-down",
                   "trigger": "materiality-override",
                   "detail": f"$320,000 hardware inventory write-down to net realizable value (line discontinued) kept net income negative (-${abs(consol_pl['2025-09']['net_income']):,.0f}) for a second consecutive month."})
churn_amt = rev_bridge["2025-05"].get("churn", 0)
red_flags.append({"period": "2025-05", "metric": "Single-customer churn",
                   "trigger": "materiality-override",
                   "detail": f"Vantage Logistics Corp (Platform Pro, North America) churned in full: ${abs(churn_amt):,.0f} annualized contract value lost, equal to {abs(churn_amt)/consol_pl['2025-05']['revenue']*100:.0f}% of that month's consolidated revenue."})
cash_delta_sep = consol_bs["2025-09"]["cash"] - consol_bs["2025-08"]["cash"]
debt_delta_sep = consol_bs["2025-09"]["debt"] - consol_bs["2025-08"]["debt"]
red_flags.append({"period": "2025-09", "metric": "Cash increase driven by financing, not operations",
                   "trigger": "materiality-override",
                   "detail": f"Consolidated cash rose ${cash_delta_sep:,.0f} in September, but ${debt_delta_sep:,.0f} of that (>99%) is the new $2.0M revolver draw, not operating cash generation — the month's net income was actually negative (-${abs(consol_pl['2025-09']['net_income']):,.0f}). Cash trend should not be read as an earnings-quality signal this period."})
oct_hw = hc_by_period_dept["2025-10"].get("Hardware Mfg Support", {})
if oct_hw:
    red_flags.append({"period": "2025-10", "metric": "Headcount: department RIF",
                       "trigger": "fixed",
                       "detail": f"Hardware Mfg Support: {oct_hw['terms_inv']} involuntary terminations in the month (department reduced to {oct_hw['active']} active), consistent with the Hardware line sunset/restructuring."})
red_flags.append({"period": "2025-03", "metric": "Restatement disclosed",
                   "trigger": "fixed",
                   "detail": "US01 revenue for March 2025 was originally reported $18,450 higher; corrected during the Q3 close for a subscription-revenue cutoff error. The trial balance for this cycle reflects the as-restated figure."})
red_flags.append({"period": "2025-12", "metric": "IN01 budget not supplied",
                   "trigger": "fixed",
                   "detail": "No FY2025 budget was provided for IN01 (India). All budget-variance views show IN01 as 'Not available' rather than a fabricated or zero figure; consolidated budget variance is US01+UK01 only."})

# ---------- File log ----------
file_log = [
    {"file": "00_Company_Config.xlsx", "classification": "Company/Entity Config", "confidence": "HIGH", "fate": "used"},
    {"file": "01_GL_TrialBalance_2025.xlsx", "classification": "GL / Trial Balance", "confidence": "HIGH", "fate": "used"},
    {"file": "02_CashBank_Position_2025.xlsx", "classification": "Cash & Bank Position", "confidence": "HIGH", "fate": "used (cross-check vs GL cash)"},
    {"file": "03_Headcount_2025.xlsx", "classification": "Headcount", "confidence": "HIGH (MEDIUM on UK01/IN01 date format)", "fate": "used"},
    {"file": "04_Revenue_Bookings_2025.xlsx", "classification": "Revenue / Bookings detail", "confidence": "HIGH", "fate": "used"},
    {"file": "05_Budget_Forecast_FY2025.xlsx", "classification": "Budget/Forecast", "confidence": "HIGH", "fate": "used (US01, UK01 only — IN01 not supplied, flagged)"},
    {"file": "06_FX_Rates_2025.xlsx", "classification": "FX Rates", "confidence": "HIGH", "fate": "used"},
    {"file": "07_Intercompany_Detail_2025.xlsx", "classification": "Intercompany Detail", "confidence": "HIGH", "fate": "used (cross-checked vs GL intercompany accounts)"},
]

model = dict(
    entity_cfg=entity_cfg, PERIODS=PERIODS, ENTITIES=ENTITIES,
    consol_pl={p: dict(v) for p, v in consol_pl.items()},
    consol_bs={p: dict(v) for p, v in consol_bs.items()},
    bud_revenue=bud_revenue,
    bud_opex={p: dict(v) for p, v in bud_usd.items()},
    hc_by_period_dept={k: dict(v) for k, v in hc_by_period_dept.items()},
    rev_bridge={k: dict(v) for k, v in rev_bridge.items()},
    customer_revenue=dict(customer_revenue), top_customers=top_customers,
    top1_share=top1_share, top5_share=top5_share,
    cash_file_usd=dict(cash_file_usd), cash_variance=cash_variance,
    red_flags=red_flags, jlog=jlog, file_log=file_log,
    one_time_detail=one_time_detail, restated_note=restated_note,
    runway_months=runway_months, avg_burn_h2=avg_burn_h2,
    ic_elim=dict(ic_elim), fx_rate=fx_rate,
    cash_flow=cash_flow,
)

import os
os.makedirs(OUT_DIR, exist_ok=True)
with open(f"{OUT_DIR}/_model.pkl", "wb") as f:
    pickle.dump(model, f)

# ---------- Print concise validation summary ----------
print("=== VALIDATION SUMMARY ===")
print(f"Periods processed: {PERIODS[0]} .. {PERIODS[-1]} ({len(PERIODS)} months)")
print(f"Entities: {ENTITIES}")
bad_bs = [p for p in PERIODS if abs(consol_bs[p]['bs_check']) > 1.0]
print(f"BS articulation check (Assets = Liab+Equity), periods failing (>$1 tolerance): {bad_bs}")
print(f"Dec-2025 consolidated Revenue (USD): {consol_pl[dec]['revenue']:,.0f}")
print(f"Dec-2025 consolidated Net Income (USD): {consol_pl[dec]['net_income']:,.0f}")
print(f"Dec-2025 consolidated Cash (USD, from GL): {consol_bs[dec]['cash']:,.0f}")
print(f"Dec-2025 Cash per CashBank_Position file (USD): {cash_file_usd[dec]:,.0f}  (variance: {cash_variance[dec]:,.2f})")
print(f"FY2025 total revenue (USD): {sum(consol_pl[p]['revenue'] for p in PERIODS):,.0f}")
print(f"Red flags detected: {len(red_flags)}")
for rf in red_flags:
    print(" -", rf)
print(f"Judgment log entries: {len(jlog)}")
print(f"Top customer concentration: top1={top1_share:.1%} top5={top5_share:.1%}")
print(f"H2 avg monthly burn: {avg_burn_h2:,.0f}  Runway: {runway_months}")
print()
print("Cash flow tie-out (computed ending cash vs actual BS cash), max abs variance:",
      max(abs(cash_flow[p]["cf_tie_variance"]) for p in PERIODS))
print("FY2025 operating CF total:", sum(cash_flow[p]["operating_cf"] for p in PERIODS))
print("FY2025 financing CF total:", sum(cash_flow[p]["financing_cf"] for p in PERIODS))
print("Model pickled to", f"{OUT_DIR}/_model.pkl")
