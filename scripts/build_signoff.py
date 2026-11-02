#!/usr/bin/env python3
"""
Phase 9 -- Reconciliation & Sign-off file: the real-data counterpart of a test answer key.

Build (after Phase 8 QA passes):
  python scripts/build_signoff.py --spec <outputs>/_work/<Period>/signoff_inputs.json [--outputs DIR]

Read back the client's filled-in copy:
  python scripts/build_signoff.py --check <outputs>/<Period>_Reconciliation_and_Signoff.xlsx

The spec JSON (written by the cycle's build scripts, figures read FROM the model workbook, never retyped):
  {
    "period": "2026-Q1", "period_text": "Q1 FY2026 (1 Jan - 31 Mar 2026)", "currency": "USD",
    "basis": "US GAAP (HQ01); IFRS (DE01)", "company": "Example Group",
    "workbook": "2026-Q1_Business_Review_USD.xlsx",
    "figures": [   # control totals the client can check against systems they already trust
      {"section": "Entity (own currency)", "entity": "DE01", "figure": "Revenue", "unit": "EUR",
       "value": 1234567, "source": "FX & Consolidation!D14", "where_to_find": "DE01 accounting system P&L",
       "tolerance": "1", "client_value": null}
    ],
    "checks":    [{"check": "Balance sheet balances, every period", "result": "PASS", "detail": "..."}],
    "judgments": [{"ref": "J3", "tier": "MEDIUM", "item": "...", "treatment": "...", "impact": "..."}],
    "open_items":[{"item": "...", "why_open": "...", "needed": "..."}]
  }
tolerance: "0.5%" = relative to the review's figure; a plain number = absolute difference. Default "0.1%".
If "judgments" is omitted, MEDIUM/LOW rows are read from the workbook's Assumptions & Judgment Log tab.

Rebuilding the same period keeps whatever the client already typed (their figures, responses, comments,
owners, statuses), matched by entity+figure / judgment item / open item; the sign-off decision is reset
because the deliverables changed.

--check exit codes: 0 = approved, every figure PASS, no corrections or open items; 1 = items need action;
2 = not filled in yet / file unreadable.
"""
import argparse
import json
import os
import re
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import brlib  # noqa: E402

from openpyxl import Workbook, load_workbook  # noqa: E402
from openpyxl.formatting.rule import CellIsRule  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402
from openpyxl.worksheet.datavalidation import DataValidation  # noqa: E402

NAVY = "1F3864"
FILL_H = PatternFill("solid", fgColor=NAVY)
FILL_SEC = PatternFill("solid", fgColor="D9E2F3")
FILL_IN = PatternFill("solid", fgColor="FFF2CC")
GREEN, AMBER, RED = (PatternFill("solid", fgColor=c) for c in ("C6EFCE", "FFEB9C", "FFC7CE"))
F_TITLE = Font(size=16, bold=True, color=NAVY)
F_SUB = Font(size=11, italic=True, color="595959")
F_H = Font(bold=True, color="FFFFFF")
F_B = Font(bold=True)
F_IN = Font(color="0000FF")
WRAP = Alignment(wrap_text=True, vertical="top")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

S_START, S_FIG, S_CHK, S_JDG, S_OPEN, S_SIGN = ("Start Here", "1 Control Totals", "2 Built-in Checks",
                                               "3 Judgment Calls", "4 Open Items", "5 Sign-off")
FIG_HDR = ["Ref", "Section", "Entity", "Figure", "Unit", "Review's figure", "Where it comes from (model)",
           "Your figure", "Difference", "Result", "Tolerance", "Where to find your figure", "Your comment"]
JDG_HDR = ["Ref", "Confidence", "Item", "What the review did", "Effect on the numbers", "Your response",
           "Your correction / comment"]
OPEN_HDR = ["Ref", "Item", "Why it is still open", "What is needed to close it", "Owner", "Status", "Comment"]
JDG_CHOICES = ("Agree", "Correct this", "Need more info", "Not reviewed")
OPEN_CHOICES = ("Open", "Resolved", "Accepted as is")
DECISIONS = ("Approved", "Approved with comments", "Not approved", "Pending")
SIGN_LABELS = ["Reviewed by", "Role", "Date reviewed", "Decision", "Comments"]


# ------------------------------------------------------------------ helpers
def num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace(" ", "")
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    try:
        return -float(s) if neg else float(s)
    except ValueError:
        return None


def tol_limit(tol, review_value):
    tol = str(tol or "0.1%").strip()
    if tol.endswith("%"):
        return abs(review_value or 0) * float(tol[:-1]) / 100, f"*{float(tol[:-1]) / 100}"
    return float(tol), None


def header(ws, row, headers, widths):
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.font, cell.fill, cell.alignment, cell.border = F_H, FILL_H, WRAP, BOX
        ws.column_dimensions[get_column_letter(c)].width = widths[c - 1]


def body(ws, row, values, inputs=()):
    for c, v in enumerate(values, 1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.alignment, cell.border = WRAP, BOX
        if c in inputs:
            cell.fill, cell.font = FILL_IN, F_IN


def titled(wb, name, title, sub):
    ws = wb.create_sheet(name)
    ws["A1"], ws["A2"] = title, sub
    ws["A1"].font, ws["A2"].font = F_TITLE, F_SUB
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


def table_rows(ws, ncols, start=5):
    """Yield (row_number, [values]) for a data table whose header sits on row start-1."""
    for r in range(start, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, ncols + 1)]
        if vals[0] in (None, ""):
            break
        yield r, vals


def judgments_from_workbook(path):
    """Fallback: MEDIUM/LOW rows from the model's Assumptions & Judgment Log tab, whatever its columns."""
    if not path or not os.path.exists(path):
        return []
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = next((w for w in wb.worksheets if "judg" in w.title.lower()), None)
        if ws is None:
            return []
        rows = list(ws.iter_rows(values_only=True))
        hi = next((i for i, r in enumerate(rows[:15])
                   if any(isinstance(v, str) and re.search(r"tier|confidence", v, re.I) for v in r)), None)
        if hi is None:
            return []
        hdr = [str(v or "").lower() for v in rows[hi]]

        def col(*keys):
            return next((i for i, h in enumerate(hdr) if any(k in h for k in keys)), None)
        ct = col("tier", "confidence")
        ci = col("item", "area", "topic", "subject", "category")
        cd = col("treatment", "decision", "detail", "assumption", "description", "rationale")
        cm = col("impact", "effect")
        out = []
        for r in rows[hi + 1:]:
            tier = str(r[ct] or "").upper() if ct is not None and ct < len(r) else ""
            if not any(t in tier for t in ("MEDIUM", "LOW")):
                continue
            g = lambda i: (str(r[i]) if i is not None and i < len(r) and r[i] is not None else "")  # noqa: E731
            out.append({"tier": "LOW" if "LOW" in tier else "MEDIUM", "item": g(ci) or g(cd)[:80],
                        "treatment": g(cd), "impact": g(cm)})
        return out
    finally:
        wb.close()


def read_existing(path):
    """Client entries from a previous build of the same period's file, for carry-over."""
    keep = {"fig": {}, "jdg": {}, "open": {}}
    if not os.path.exists(path):
        return keep
    try:
        wb = load_workbook(path, data_only=True)
    except Exception as e:  # unreadable / mid-sync: rebuild clean, but say so
        brlib.warn(f"Could not read existing {os.path.basename(path)} for carry-over ({type(e).__name__}); "
                   "client entries in it were not carried forward.")
        return keep
    if S_FIG in wb.sheetnames:
        for _, v in table_rows(wb[S_FIG], len(FIG_HDR)):
            keep["fig"][f"{v[2]}|{v[3]}"] = (v[7], v[12])
    if S_JDG in wb.sheetnames:
        for _, v in table_rows(wb[S_JDG], len(JDG_HDR)):
            keep["jdg"][str(v[2])] = (v[5], v[6])
    if S_OPEN in wb.sheetnames:
        for _, v in table_rows(wb[S_OPEN], len(OPEN_HDR)):
            keep["open"][str(v[1])] = (v[4], v[5], v[6])
    return keep


# ------------------------------------------------------------------ build
def build(spec, outputs):
    period, ccy = spec["period"], spec["currency"]
    path = os.path.join(outputs, f"{period}_Reconciliation_and_Signoff.xlsx")
    keep = read_existing(path)
    workbook_name = spec.get("workbook", f"{period}_Business_Review_{ccy}.xlsx")
    judgments = spec.get("judgments")
    if judgments is None:
        judgments = judgments_from_workbook(os.path.join(outputs, workbook_name))
    figures, checks, opens = spec.get("figures", []), spec.get("checks", []), spec.get("open_items", [])
    if not figures:
        brlib.warn("No control figures in the spec -- the client has nothing to reconcile against.")

    wb = Workbook()
    wb.remove(wb.active)

    # Start Here
    ws = titled(wb, S_START, f"{period} Reconciliation & Sign-off",
                f"{spec.get('company', '')}  |  Reporting currency {ccy}  |  {spec.get('period_text', period)}  |  "
                f"Accounting basis: {spec.get('basis', 'see Config & Setup in the model')}")
    ws.column_dimensions["A"].width = 115
    lines = [
        ("h", "What this file is for"),
        ("p", "A test file comes with an answer key. Real data doesn't, so this file plays that role: it compares "
              "the review with figures you already trust, shows the review's own checks, lists every judgment call "
              "it made, and records your sign-off. Only the yellow cells need filling in."),
        ("h", "How to complete it (about 15-30 minutes)"),
        ("p", f"1. '{S_FIG}': type the figure from your own system (accounting system, bank statement, HR/payroll, "
              "bank compliance certificate) into 'Your figure'. 'Result' shows PASS or CHECK. Leave a row blank if "
              "you have no independent source for it."),
        ("p", f"2. '{S_CHK}': read only. These are the checks the review ran on itself."),
        ("p", f"3. '{S_JDG}': for each call the review made, choose Agree, Correct this, or Need more info. "
              "If you choose Correct this, say what is right in the comment."),
        ("p", f"4. '{S_OPEN}': assign an owner and a status to anything the review could not resolve."),
        ("p", f"5. '{S_SIGN}': enter your name, role and date, and choose a decision."),
        ("h", "What happens next"),
        ("p", "Save this file where it is (same name, same folder) and tell Claude: \"Check the sign-off for "
              f"{period}.\" Anything marked CHECK, Correct this or Need more info is investigated. If the "
              "deliverables need fixing they are rebuilt, and this file is regenerated with your entries kept "
              "(only the decision is reset, because the numbers may have changed)."),
        ("h", "Files covered by this sign-off"),
        ("p", f"{workbook_name} (model), plus the memo and deck for {period} built from it. Prepared "
              f"{date.today():%d %b %Y}."),
    ]
    for i, (kind, text) in enumerate(lines, 4):
        c = ws.cell(row=i, column=1, value=text)
        c.alignment = WRAP
        if kind == "h":
            c.font, c.fill = F_B, FILL_SEC

    # 1 Control Totals
    ws = titled(wb, S_FIG, "Control totals: the review's figures vs your own systems",
                "Type your figure in the yellow column. PASS = within tolerance. Blank = not checked.")
    header(ws, 4, FIG_HDR, [6, 18, 10, 30, 8, 17, 26, 17, 14, 9, 9, 32, 30])
    r = 4
    for i, f in enumerate(figures, 1):
        r = 4 + i
        key = f"{f.get('entity', '')}|{f['figure']}"
        prior_val, prior_comment = keep["fig"].get(key, (None, None))
        client = f.get("client_value") if f.get("client_value") is not None else prior_val
        client = num(client) if num(client) is not None else client
        tol = str(f.get("tolerance", "0.1%"))
        body(ws, r, [f"C{i}", f.get("section", ""), f.get("entity", ""), f["figure"], f.get("unit", ccy),
                     f.get("value"), f.get("source", ""), client, None, None, tol, f.get("where_to_find", ""),
                     prior_comment], inputs=(8, 13))
        abs_lim, pct = tol_limit(tol, num(f.get("value")))
        limit = f"ABS(F{r}){pct}" if pct else f"{abs_lim}"
        ws.cell(row=r, column=9, value=f'=IF(H{r}="","",H{r}-F{r})')
        ws.cell(row=r, column=10, value=f'=IF(H{r}="","",IF(ABS(H{r}-F{r})<={limit},"PASS","CHECK"))')
        fmt = "0.00" if f.get("unit") in ("x", "ratio") else ("0.0%" if f.get("unit") == "%" else "#,##0;(#,##0)")
        for c in (6, 8, 9):
            ws.cell(row=r, column=c).number_format = fmt
    fig_last = max(r, 5)
    ws.conditional_formatting.add(f"J5:J{fig_last}", CellIsRule(operator="equal", formula=['"PASS"'], fill=GREEN))
    ws.conditional_formatting.add(f"J5:J{fig_last}", CellIsRule(operator="equal", formula=['"CHECK"'], fill=RED))
    ws.freeze_panes = "E5"
    ws.cell(row=fig_last + 2, column=2, value="Tolerance: a % is relative to the review's figure; a plain number "
            "is an absolute difference. Group figures in the reporting currency can differ from your own "
            "consolidation if different exchange rates were used; the review's rates are on the model's "
            "FX & Consolidation tab.").font = F_SUB

    # 2 Built-in Checks
    ws = titled(wb, S_CHK, "Checks the review ran on itself", "Read only. These prove the review is internally "
                "consistent; your control totals prove it matches reality.")
    header(ws, 4, ["Ref", "Check", "Result", "Detail"], [6, 48, 10, 90])
    r = 4
    for i, c in enumerate(checks, 1):
        r = 4 + i
        body(ws, r, [f"K{i}", c["check"], c.get("result", ""), c.get("detail", "")])
    chk_last = max(r, 5)
    for word, fill in (("PASS", GREEN), ("WARN", AMBER), ("FAIL", RED)):
        ws.conditional_formatting.add(f"C5:C{chk_last}", CellIsRule(operator="equal", formula=[f'"{word}"'], fill=fill))

    # 3 Judgment Calls
    ws = titled(wb, S_JDG, "Judgment calls the review made", "Every MEDIUM and LOW confidence decision from the "
                "Assumptions & Judgment Log. Choose a response for each.")
    header(ws, 4, JDG_HDR, [6, 11, 30, 60, 30, 16, 40])
    r = 4
    for i, j in enumerate(judgments, 1):
        r = 4 + i
        resp, comment = keep["jdg"].get(str(j.get("item", "")), (None, None))
        body(ws, r, [j.get("ref", f"J{i}"), j.get("tier", ""), j.get("item", ""), j.get("treatment", ""),
                     j.get("impact", ""), resp or "Not reviewed", comment], inputs=(6, 7))
    jdg_last = max(r, 5)
    dv = DataValidation(type="list", formula1='"' + ",".join(JDG_CHOICES) + '"')
    ws.add_data_validation(dv)
    dv.add(f"F5:F{jdg_last}")
    for word, fill in (("Agree", GREEN), ("Need more info", AMBER), ("Correct this", RED)):
        ws.conditional_formatting.add(f"F5:F{jdg_last}", CellIsRule(operator="equal", formula=[f'"{word}"'], fill=fill))
    ws.freeze_panes = "D5"

    # 4 Open Items
    ws = titled(wb, S_OPEN, "Open items", "Things the review could not resolve from the data. Assign an owner.")
    header(ws, 4, OPEN_HDR, [6, 40, 45, 45, 18, 16, 35])
    r = 4
    for i, o in enumerate(opens, 1):
        r = 4 + i
        owner, status, comment = keep["open"].get(str(o["item"]), (None, None, None))
        body(ws, r, [f"O{i}", o["item"], o.get("why_open", ""), o.get("needed", ""), owner,
                     status or "Open", comment], inputs=(5, 6, 7))
    open_last = max(r, 5)
    dv = DataValidation(type="list", formula1='"' + ",".join(OPEN_CHOICES) + '"')
    ws.add_data_validation(dv)
    dv.add(f"F5:F{open_last}")

    # 5 Sign-off
    ws = titled(wb, S_SIGN, "Sign-off", f"{period}  |  {ccy}  |  {spec.get('basis', '')}")
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 50, 40
    fr, jr, orng = f"'{S_FIG}'!J5:J{fig_last}", f"'{S_JDG}'!F5:F{jdg_last}", f"'{S_OPEN}'!F5:F{open_last}"
    summary = [
        ("Control totals checked", f'=COUNTIF({fr},"PASS")+COUNTIF({fr},"CHECK")'),
        ("  PASS", f'=COUNTIF({fr},"PASS")'),
        ("  CHECK (outside tolerance)", f'=COUNTIF({fr},"CHECK")'),
        ("Built-in checks failed", f"=COUNTIF('{S_CHK}'!C5:C{chk_last},\"FAIL\")"),
        ("Judgment calls: agreed", f'=COUNTIF({jr},"Agree")'),
        ("Judgment calls: to correct / need more info", f'=COUNTIF({jr},"Correct this")+COUNTIF({jr},"Need more info")'),
        ("Judgment calls: not reviewed", f'=COUNTIF({jr},"Not reviewed")'),
        ("Open items still open", f'=COUNTIF({orng},"Open")'),
    ]
    header(ws, 4, ["Summary", "Count"], [50, 40])
    for i, (label, formula) in enumerate(summary, 5):
        body(ws, i, [label, formula])
    r0 = 5 + len(summary) + 1
    header(ws, r0, ["Sign-off", "Your entry"], [50, 40])
    for i, label in enumerate(SIGN_LABELS, r0 + 1):
        body(ws, i, [label, "Pending" if label == "Decision" else None], inputs=(2,))
    dec_row = r0 + 1 + SIGN_LABELS.index("Decision")
    dv = DataValidation(type="list", formula1='"' + ",".join(DECISIONS) + '"')
    ws.add_data_validation(dv)
    dv.add(f"B{dec_row}")
    ws.cell(row=r0 + 1 + SIGN_LABELS.index("Date reviewed"), column=2).number_format = "dd mmm yyyy"
    vr = r0 + len(SIGN_LABELS) + 2
    ws.cell(row=vr, column=1, value="Status").font = F_B
    ws.cell(row=vr, column=2, value=(
        f'=IF(B{dec_row}="Pending","Awaiting review",IF(AND(B7=0,B8=0,B10=0,B12=0,'
        f'LEFT(B{dec_row},8)="Approved"),"Signed off","Signed off with items to action"))')).font = F_B

    return brlib.safe_save(wb.save, path)


# ------------------------------------------------------------------ check
def check(path):
    try:
        wb = load_workbook(path, data_only=True)
    except Exception as e:
        print(f"[ERROR] Cannot read {path}: {type(e).__name__}: {e}")
        return 2
    actions, passed, blank = [], 0, 0
    for _, v in table_rows(wb[S_FIG], len(FIG_HDR)):
        review, client = num(v[5]), num(v[7])
        label = f"{v[0]} {v[2]} {v[3]}".replace("  ", " ")
        if client is None:
            blank += 1
            continue
        lim, _ = tol_limit(v[10], review)
        if review is not None and abs(client - review) <= lim + 1e-9:
            passed += 1
        else:
            rv = f"{review:,.2f}" if review is not None else "(none)"
            actions.append(f"CHECK  {label}: review {rv} vs yours {client:,.2f} "
                           f"(diff {client - (review or 0):,.2f}, tolerance {v[10]})"
                           + (f" -- comment: {v[12]}" if v[12] else ""))
    for _, v in table_rows(wb[S_CHK], 4):
        if str(v[2]).upper() == "FAIL":
            actions.append(f"FAILED CHECK  {v[0]} {v[1]}: {v[3]}")
    unreviewed = judged = 0
    for _, v in table_rows(wb[S_JDG], len(JDG_HDR)):
        judged += 1
        resp = str(v[5] or "Not reviewed")
        if resp in ("Correct this", "Need more info"):
            actions.append(f"{resp.upper()}  {v[0]} {v[2]}" + (f" -- {v[6]}" if v[6] else " -- (no comment given)"))
        elif resp == "Not reviewed":
            unreviewed += 1
    for _, v in table_rows(wb[S_OPEN], len(OPEN_HDR)):
        if str(v[5] or "Open") == "Open":
            actions.append(f"OPEN  {v[0]} {v[1]}" + (f" (owner: {v[4]})" if v[4] else " (no owner)"))
    sign = {}
    ws = wb[S_SIGN]
    for r in range(1, ws.max_row + 1):
        if ws.cell(row=r, column=1).value in SIGN_LABELS:
            sign[ws.cell(row=r, column=1).value] = ws.cell(row=r, column=2).value
    decision = str(sign.get("Decision") or "Pending")
    if hasattr(sign.get("Date reviewed"), "strftime"):
        sign["Date reviewed"] = sign["Date reviewed"].strftime("%d %b %Y")

    print(f"Sign-off file: {os.path.basename(path)}")
    print(f"  Control totals: {passed} PASS, {sum(a.startswith('CHECK') for a in actions)} CHECK, {blank} not entered")
    print(f"  Judgment calls not reviewed: {unreviewed}")
    print(f"  Decision: {decision}  |  Reviewed by: {sign.get('Reviewed by') or '-'}  |  "
          f"Date: {sign.get('Date reviewed') or '-'}")
    if sign.get("Comments"):
        print(f"  Reviewer comments: {sign['Comments']}")
    for a in actions:
        print(f"  [ACTION] {a}")
    entered = (passed + sum(a.startswith("CHECK") for a in actions)) > 0
    if decision == "Pending" and not entered and unreviewed == judged:
        print("NOT FILLED IN YET")
        return 2
    if decision.startswith("Approved") and not actions:
        print("SIGNED OFF")
        return 0
    print(f"ITEMS TO ACTION ({len(actions)})" if actions else f"NOT SIGNED OFF (decision: {decision})")
    return 1


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--spec", help="signoff_inputs.json written by the cycle")
    g.add_argument("--check", help="Read back a client-completed sign-off file")
    ap.add_argument("--outputs", help="Override the outputs folder")
    args = ap.parse_args()
    if args.check:
        sys.exit(check(args.check))
    with open(args.spec, encoding="utf-8") as fh:
        spec = json.load(fh)
    outputs = args.outputs or brlib.resolve_paths()[1]
    written = build(spec, outputs)
    print(f"Wrote {written}")


if __name__ == "__main__":
    main()
