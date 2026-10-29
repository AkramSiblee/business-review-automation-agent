#!/usr/bin/env python3
"""
Phase 8 QA -- force Excel to fully recalculate a workbook, scan every sheet for formula errors, and
check the articulation rows (Balance Sheet "Articulation check", Cash Flow "Tie-out variance") are ~0.

  python scripts/verify_workbook.py <path-to-workbook.xlsx> [--no-save] [--tolerance 1.0]

Runs in a private Excel instance (never touches workbooks you have open). Saving after recalculation
(the default) stores computed values in the file, so openpyxl-based readers and phone previews see
numbers rather than empty formula cells.
Exit code 0 = PASS, 1 = FAIL, 3 = Excel automation unavailable (non-Windows / no desktop Excel) --
do the recalculation check manually in that case and say so in the QA summary.
"""
import argparse
import os
import sys

ERROR_VALS = {"#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NULL!", "#NUM!", "#SPILL!", "#CALC!"}
# Excel COM returns errors as negative ints (e.g. -2146826281 = #DIV/0!), not strings.
COM_ERROR_CODES = {-2146826281, -2146826246, -2146826259, -2146826288, -2146826252, -2146826265,
                   -2146826273}
CHECK_ROWS = {"Balance Sheet": "Articulation check", "Cash Flow": "Tie-out variance"}


def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("path")
    ap.add_argument("--no-save", action="store_true")
    ap.add_argument("--tolerance", type=float, default=1.0)
    args = ap.parse_args()
    path = os.path.abspath(args.path)
    if not os.path.exists(path):
        print(f"[FAIL] Not found: {path}")
        sys.exit(1)
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        print("[SKIP] pywin32 not available -- Excel recalculation QA cannot run on this machine.")
        sys.exit(3)

    pythoncom.CoInitialize()
    excel = wb = None
    errors, check_results = [], {}
    try:
        try:
            excel = win32.DispatchEx("Excel.Application")
        except Exception as e:
            print(f"[SKIP] Desktop Excel not available via COM ({e}).")
            sys.exit(3)
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        wb = excel.Workbooks.Open(path, UpdateLinks=0, ReadOnly=args.no_save)
        excel.CalculateFullRebuild()

        for ws in wb.Worksheets:
            vals = ws.UsedRange.Value
            if vals is None:
                continue
            if not isinstance(vals, tuple):
                vals = ((vals,),)
            r0, c0 = ws.UsedRange.Row, ws.UsedRange.Column
            for ri, row in enumerate(vals):
                for ci, v in enumerate(row):
                    if (isinstance(v, str) and v in ERROR_VALS) or (isinstance(v, int) and v in COM_ERROR_CODES):
                        errors.append(f"{ws.Name}!{col_letter(c0 + ci)}{r0 + ri}")
                label = row[0] if row else None
                if ws.Name in CHECK_ROWS and isinstance(label, str) and CHECK_ROWS[ws.Name] in label:
                    nums = [abs(v) for v in row[1:] if isinstance(v, (int, float)) and not isinstance(v, bool)]
                    bools = [v for v in row[1:] if isinstance(v, bool)]
                    check_results[ws.Name] = (max(nums) if nums else 0.0, all(bools) if bools else True)

        if not args.no_save:
            wb.Save()
    finally:
        if wb is not None:
            wb.Close(False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()

    ok = not errors
    print(f"Formula errors: {len(errors)}")
    for e in errors[:40]:
        print(f"  {e}")
    for sheet, label in CHECK_ROWS.items():
        if sheet not in check_results:
            print(f"[WARN] {sheet} '{label}' row not found -- articulation not verified")
            continue
        max_abs, all_true = check_results[sheet]
        passed = max_abs < args.tolerance and all_true
        ok &= passed
        print(f"{sheet} {label}: max abs variance {max_abs:,.2f} -> {'OK' if passed else 'FAIL'}")
    print("PASS" if ok else "FAIL")
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
