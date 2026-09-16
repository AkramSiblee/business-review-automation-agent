"""Phase 8 QA — force Excel to recalculate the workbook and scan every cell for formula errors."""
import win32com.client as win32
import os, sys

path = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else "data/outputs/2025-FY_Business_Review_USD.xlsx")

excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(path)
excel.CalculateFullRebuild()
excel.CalculateUntilAsyncQueriesDone()

errors_found = []
ERROR_VALS = {"#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NULL!", "#NUM!"}

for ws in wb.Worksheets:
    used = ws.UsedRange
    vals = used.Value
    if vals is None:
        continue
    rows = vals if isinstance(vals, tuple) and isinstance(vals[0], tuple) else ((vals,),)
    for ri, row in enumerate(rows, start=1):
        for ci, v in enumerate(row, start=1):
            if isinstance(v, str) and v in ERROR_VALS:
                errors_found.append((ws.Name, ri, ci, v))

# specific checks: BS articulation and CF tie-out rows read TRUE/0
bs_ws = wb.Worksheets("Balance Sheet")
cf_ws = wb.Worksheets("Cash Flow")

def find_row(ws, label_substr):
    used = ws.UsedRange
    for r in range(1, used.Rows.Count + 1):
        v = ws.Cells(r, 1).Value
        if v and label_substr in str(v):
            return r
    return None

bs_check_row = find_row(bs_ws, "Articulation check")
cf_tie_row = find_row(cf_ws, "Tie-out variance")

bs_max_abs = 0.0
if bs_check_row:
    for c in range(2, 14):
        v = bs_ws.Cells(bs_check_row, c).Value
        if isinstance(v, (int, float)):
            bs_max_abs = max(bs_max_abs, abs(v))

cf_max_abs = 0.0
if cf_tie_row:
    for c in range(2, 14):
        v = cf_ws.Cells(cf_tie_row, c).Value
        if isinstance(v, (int, float)):
            cf_max_abs = max(cf_max_abs, abs(v))

wb.Save()
wb.Close(False)
excel.Quit()

print("Formula errors found:", len(errors_found))
for e in errors_found[:30]:
    print(" ", e)
print(f"Balance Sheet articulation max abs variance across FY2025: {bs_max_abs}")
print(f"Cash Flow tie-out max abs variance across FY2025: {cf_max_abs}")
print("PASS" if not errors_found and bs_max_abs < 1.0 and cf_max_abs < 1.0 else "FAIL")
