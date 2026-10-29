"""
Shared helpers for every business-review script: path resolution, input discovery, tab loading,
category matching, period/date detection, and safe file writes.

Every script imports this instead of hardcoding "data/inputs" / "data/outputs", so the pipeline
works from any working directory and always honours config/data-paths.json.
"""
import json
import os
import re
import sys
import tempfile
from collections import Counter
from datetime import date, datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(ROOT, "config", "data-paths.json")

INPUT_EXTS = {".xlsx", ".xlsm", ".xls", ".csv", ".docx", ".pdf", ".pptx"}
# Excel lock files, Drive/OS cruft, and half-synced temp files must never be treated as inputs.
IGNORE_PREFIXES = ("~$", ".", "_")
IGNORE_NAMES = {"desktop.ini", "thumbs.db"}

# Standard categories from references/standard-input-file-templates.md.
# keywords: matched against the tab/file name; columns: the contract headers.
CATEGORIES = {
    "00_Company_Config": {
        "keywords": ["config", "entity", "entities", "setup", "company"],
        "columns": ["Entity Code", "Entity Name", "Country", "Functional Currency", "Accounting Standard",
                    "Fiscal Year Start Month", "Ownership %", "Consolidation Method", "Reporting Currency"],
        "recurring": False,
    },
    "01_GL_TrialBalance": {
        "keywords": ["gl", "trial", "tb", "ledger", "p&l", "pnl", "income", "balance sheet"],
        "columns": ["Entity Code", "Period", "Account Code", "Account Name", "Account Type", "Department",
                    "Debit", "Credit", "Currency", "One-Time Item", "Notes"],
        "recurring": True,
    },
    "02_CashBank_Position": {
        "keywords": ["cash", "bank", "treasury"],
        "columns": ["Entity Code", "Period", "Bank/Account Name", "Account Number (last 4)", "Currency",
                    "Beginning Balance", "Ending Balance", "Restricted Cash", "Notes"],
        "recurring": True,
    },
    "03_Headcount": {
        "keywords": ["headcount", "hc", "employee", "staff", "payroll", "people"],
        "columns": ["Entity Code", "Employee ID", "Department", "Job Title", "Employment Status", "Hire Date",
                    "Termination Date", "Termination Type", "FTE %", "Monthly Fully-Loaded Cost", "Currency",
                    "Country/Location"],
        "recurring": True,
    },
    "04_Revenue_Bookings": {
        "keywords": ["revenue", "bookings", "sales", "deals", "contracts", "arr"],
        "columns": ["Entity Code", "Period/Transaction Date", "Customer ID/Name", "Deal/Contract ID",
                    "Revenue Type", "Product/Segment", "Region", "Amount", "Currency", "Contract Start Date",
                    "Contract End Date", "Recurring"],
        "recurring": True,
    },
    "05_Budget_Forecast": {
        "keywords": ["budget", "forecast", "plan", "reforecast"],
        "columns": ["Entity Code", "Fiscal Year", "Period", "Account Code", "Account Name", "Account Type",
                    "Budgeted Amount", "Currency", "Version"],
        "recurring": True,
    },
    "06_FX_Rates": {
        "keywords": ["fx", "rates", "exchange", "currency"],
        "columns": ["Period", "From Currency", "To Currency", "Rate Type", "Rate", "Source"],
        "recurring": True,
        "only_if": "multi_currency",
    },
    "07_Intercompany_Detail": {
        "keywords": ["intercompany", "interco", "ic ", "related party"],
        "columns": ["Period", "Entity (From)", "Entity (To)", "Transaction Type", "Amount", "Currency",
                    "Description"],
        "recurring": True,
        "only_if": "multi_entity",
    },
    # Optional tabs: used when present, never reported as missing.
    "08_Debt_Covenants": {
        "keywords": ["debt", "covenant", "facility", "facilities", "loan", "borrowing"],
        "columns": ["Entity Code", "Period", "Facility Name", "Facility Type", "Limit", "Drawn", "Currency",
                    "Maturity Date", "Interest Rate", "Covenant Name", "Covenant Definition", "Threshold",
                    "Comparison", "Reported Value"],
        "recurring": True,
        "optional": True,
    },
    "09_Operational_KPIs": {
        "keywords": ["kpi", "operational", "metrics", "ops"],
        "columns": ["Entity Code", "Period", "Segment", "Metric", "Value", "Unit", "Notes"],
        "recurring": True,
        "optional": True,
    },
    "10_Adjustments_Events": {
        "keywords": ["adjustment", "events", "one-time", "one off", "m&a"],
        "columns": ["Entity Code", "Period", "Event Type", "Description", "Amount", "Currency", "One-Time"],
        "recurring": True,
        "optional": True,
    },
}

# Account Type values (01_GL_TrialBalance) that are P&L activity rather than balance-sheet positions.
PNL_TYPE_HINTS = ("revenue", "cogs", "cost", "opex", "expense", "income")
BS_TYPE_HINTS = ("asset", "liabilit", "equity")

# A GL export with one signed column instead of Debit/Credit (playbook 2.4).
SIGNED_AMOUNT_HEADERS = {"amount", "balance", "value", "net", "net amount", "closing balance", "ytd",
                         "ytd balance"}
EXCEL_ERRORS = {"#REF!", "#DIV/0!", "#N/A", "#VALUE!", "#NAME?", "#NUM!", "#NULL!", "#SPILL!", "#CALC!"}
SUBTOTAL_RE = re.compile(r"\b(sub)?total\b|^\s*net (income|profit|loss)\b|\bgross (profit|margin)\b|\bebitda\b",
                         re.I)
UNIT_SCALE_RE = re.compile(
    r"\b(in )?(thousands|millions|billions|lakhs?|lacs?|crores?)\b|\(\s*0{3}s?\s*\)|'0{3}s?\b|\b0{3}s\b|"
    r"(?-i:\b[A-Z]{3}\s?(k|m|mn|bn|K|M|MN|BN)\b)", re.I)  # "USD m", "EURk" -- case-sensitive, or "Item" hits
SCRATCH_TAB_RE = re.compile(r"^(sheet\d*|copy of\b|old\b|draft\b|backup\b|temp\b|test\b|scratch\b)|\(\d+\)\s*$",
                            re.I)
ADJ_PERIOD_RE = re.compile(r"^\s*\d{4}[-/.](1[3-6])\s*$|\badj(ustment)?s?\b|\bp(eriod)?\s?1[3-6]\b", re.I)


# ---------------------------------------------------------------- paths

def load_config():
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


def resolve_paths(create=False):
    """Return (inputs_dir, outputs_dir, using_drive). Falls back to data/ when Drive isn't configured
    or the configured Drive root isn't mounted (Drive for Desktop not running)."""
    cfg = load_config()
    root = (cfg.get("google_drive_root") or "").strip()
    inputs_sub = cfg.get("inputs_subfolder") or "inputs"
    outputs_sub = cfg.get("outputs_subfolder") or "outputs"
    if root and not root.startswith("REPLACE_ME") and os.path.isdir(root):
        paths = (os.path.normpath(os.path.join(root, inputs_sub)),
                 os.path.normpath(os.path.join(root, outputs_sub)), True)
    else:
        if root and not root.startswith("REPLACE_ME"):
            warn(f"google_drive_root '{root}' is not reachable (is Google Drive for Desktop running?) -- "
                 f"falling back to local data/inputs and data/outputs.")
        paths = (os.path.join(ROOT, "data", "inputs"), os.path.join(ROOT, "data", "outputs"), False)
    if create:
        os.makedirs(paths[0], exist_ok=True)
        os.makedirs(paths[1], exist_ok=True)
    return paths


def list_input_files(inputs_dir):
    """Real input files only -- skips Excel lock files (~$...), hidden/underscore files, OS cruft."""
    files = []
    if not os.path.isdir(inputs_dir):
        return files
    for name in sorted(os.listdir(inputs_dir)):
        full = os.path.join(inputs_dir, name)
        if not os.path.isfile(full) or name.lower() in IGNORE_NAMES or name.startswith(IGNORE_PREFIXES):
            continue
        if os.path.splitext(name)[1].lower() in INPUT_EXTS:
            files.append(full)
    return files


def warn(msg):
    print(f"[WARN] {msg}", file=sys.stderr)


# ---------------------------------------------------------------- loading

def _norm(s):
    return re.sub(r"[^a-z0-9%]+", " ", str(s).lower()).strip()


def _find_header_row(rows, scan=15):
    """Exports often have title rows above the real header. Pick the first row, within the first
    `scan` rows, with the most non-empty text cells."""
    best_i, best_n = 0, -1
    for i, r in enumerate(rows[:scan]):
        n = sum(1 for c in r if isinstance(c, str) and c.strip())
        if n > best_n:
            best_i, best_n = i, n
    return best_i


def load_tables(path):
    """Yield (source_label, tab_name, headers, list_of_row_dicts, meta) for every tab of an Excel file
    (or a CSV). meta = {"hidden": bool, "preamble": text of any title rows above the header} -- the
    preamble is where exports usually state unit scale ("USD '000") and currency.
    Non-tabular formats (docx/pdf/pptx) yield nothing here -- they need Claude-driven extraction."""
    ext = os.path.splitext(path)[1].lower()
    base = os.path.basename(path)
    if ext == ".csv":
        import csv
        with open(path, encoding="utf-8-sig", newline="") as f:
            rows = [tuple(r) for r in csv.reader(f)]
        yield from _rows_to_table(base, rows)
        return
    if ext not in (".xlsx", ".xlsm"):
        return
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True))
            label = base if len(wb.worksheets) == 1 else f"{base} :: {ws.title}"
            hidden = getattr(ws, "sheet_state", "visible") != "visible"
            for t in _rows_to_table(label, rows, tab_name=ws.title, hidden=hidden):
                yield t
    finally:
        wb.close()


def _rows_to_table(label, rows, tab_name=None, hidden=False):
    rows = [r for r in rows if r is not None]
    if not rows:
        yield (label, tab_name or label, [], [], {"hidden": hidden, "preamble": ""})
        return
    hi = _find_header_row(rows)
    headers = [str(h).strip() if h is not None else f"col{i+1}" for i, h in enumerate(rows[hi])]
    preamble = " | ".join(str(c) for r in rows[:hi] for c in r if c not in (None, ""))
    data = []
    for r in rows[hi + 1:]:
        if any(c not in (None, "") for c in r):
            data.append(dict(zip(headers, r)))
    yield (label, tab_name or label, headers, data, {"hidden": hidden, "preamble": preamble})


# ---------------------------------------------------------------- classification

def classify(name, headers):
    """Return (category or None, confidence 'HIGH'|'MEDIUM'|'LOW', header_coverage 0-1, column_map).
    column_map maps contract column -> actual header found in the tab."""
    hn = {_norm(h): h for h in headers}
    nname = _norm(name)
    best = (None, 0.0, 0, {}, 0.0)
    for cat, spec in CATEGORIES.items():
        colmap = {c: hn[_norm(c)] for c in spec["columns"] if _norm(c) in hn}
        matched = len(colmap)
        if cat == "01_GL_TrialBalance" and "Debit" not in colmap and "Credit" not in colmap and \
                SIGNED_AMOUNT_HEADERS & set(hn):
            matched += 2  # a signed amount column stands in for the Debit/Credit pair
        coverage = matched / len(spec["columns"])
        prefix = cat.split("_")[0] + "_"
        name_hit = 2 if nname.replace(" ", "_").startswith(prefix) or _norm(cat[3:]) in nname else (
            1 if any(_norm(k) in nname for k in spec["keywords"]) else 0)
        score = coverage + 0.15 * name_hit
        if score > best[1]:
            best = (cat, score, name_hit, colmap, coverage)
    if best[0] is None:
        return None, "LOW", 0.0, {}
    cat, _, name_hit, colmap, coverage = best
    if coverage >= 0.8 or (coverage >= 0.6 and name_hit):
        conf = "HIGH"
    elif coverage >= 0.4 or (name_hit and coverage >= 0.2):
        conf = "MEDIUM"
    else:
        return None, "LOW", coverage, colmap
    return cat, conf, coverage, colmap


# ---------------------------------------------------------------- periods & dates

_MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], start=1)}


def to_period(val, day_first=None):
    """Best-effort conversion of a cell to 'YYYY-MM'. Returns None if it can't be read.
    day_first: True for DD/MM/YYYY, False for MM/DD/YYYY, None = guess (DD/MM if unambiguous)."""
    if val is None or val == "":
        return None
    if isinstance(val, (datetime, date)):
        return f"{val.year:04d}-{val.month:02d}"
    s = str(val).strip()
    m = re.match(r"^(\d{4})[-/.](\d{1,2})(?:[-/.]\d{1,2})?", s)
    if m:
        # Period 13-16 (audit/adjustment periods) is not a calendar month -- leave it to ADJ_PERIOD_RE.
        return f"{m.group(1)}-{int(m.group(2)):02d}" if 1 <= int(m.group(2)) <= 12 else None
    m = re.match(r"^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})", s)
    if m:
        a, b, y = int(m.group(1)), int(m.group(2)), m.group(3)
        if day_first is None:
            day_first = a > 12 or b <= 12
        month = b if day_first else a
        return f"{y}-{month:02d}" if 1 <= month <= 12 else None
    m = re.match(r"^([A-Za-z]{3})[A-Za-z]*[-\s']*(\d{2,4})$", s)
    if m and m.group(1).lower() in _MONTHS:
        y = m.group(2)
        y = ("20" + y) if len(y) == 2 else y
        return f"{y}-{_MONTHS[m.group(1).lower()]:02d}"
    return None


def month_range(first, last):
    """Every 'YYYY-MM' from first to last inclusive."""
    y, m = int(first[:4]), int(first[5:7])
    out = []
    while f"{y:04d}-{m:02d}" <= last:
        out.append(f"{y:04d}-{m:02d}")
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return out


def fiscal_year(period, fy_start_month=1):
    """Fiscal year a 'YYYY-MM' falls in, labelled by the calendar year it ends in."""
    y, m = int(period[:4]), int(period[5:7])
    return y if fy_start_month == 1 or m < fy_start_month else y + 1


def date_order_evidence(values):
    """Inspect string dates like 03/04/2025. Returns 'DD/MM', 'MM/DD', 'CONFLICT', 'AMBIGUOUS', or None
    (no slash-style string dates present). Native Excel dates are never ambiguous."""
    dd = mm = n = 0
    for v in values:
        if not isinstance(v, str):
            continue
        m = re.match(r"^\s*(\d{1,2})[/.-](\d{1,2})[/.-]\d{2,4}", v)
        if not m:
            continue
        n += 1
        a, b = int(m.group(1)), int(m.group(2))
        dd += a > 12
        mm += b > 12
    if n == 0:
        return None
    if dd and mm:
        return "CONFLICT"
    if dd:
        return "DD/MM"
    if mm:
        return "MM/DD"
    return "AMBIGUOUS"


def period_label(periods, fy_start_month=1):
    """Derive the cycle label from the set of 'YYYY-MM' periods covered: '2025-FY', '2026-Q1', '2026-03'."""
    ps = sorted(p for p in periods if p)
    if not ps:
        return "UNKNOWN-PERIOD"
    last = ps[-1]
    y, mth = int(last[:4]), int(last[5:7])
    fiscal_idx = (mth - fy_start_month) % 12 + 1  # 1..12 within the fiscal year
    fy = y if fy_start_month == 1 else (y + 1 if mth >= fy_start_month else y)
    if len(ps) >= 12 and fiscal_idx == 12:
        return f"{fy}-FY"
    if fiscal_idx % 3 == 0 and len(ps) >= 3:
        return f"{fy}-Q{fiscal_idx // 3}"
    return last


# ---------------------------------------------------------------- safe writes

def safe_save(save_fn, path):
    """Write via a temp file in the same folder, then atomically replace -- so Drive never syncs a
    half-written file. If the target is locked (open in Excel/Word/PowerPoint), save alongside it with
    a timestamp suffix instead of failing the whole cycle. Returns the path actually written."""
    folder = os.path.dirname(path) or "."
    os.makedirs(folder, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(path))
    fd, tmp = tempfile.mkstemp(prefix=f"_{stem}.", suffix=ext, dir=folder)
    os.close(fd)
    try:
        save_fn(tmp)
        try:
            os.replace(tmp, path)
            return path
        except PermissionError:
            alt = os.path.join(folder, f"{stem}_{datetime.now():%Y%m%d-%H%M%S}{ext}")
            os.replace(tmp, alt)
            warn(f"{os.path.basename(path)} is open/locked -- saved as {os.path.basename(alt)} instead. "
                 f"Close the file and re-run to write the canonical name.")
            return alt
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


def counter_summary(values, limit=8):
    c = Counter(v for v in values if v not in (None, ""))
    items = [str(k) for k, _ in sorted(c.items(), key=lambda kv: str(kv[0]))]
    return ", ".join(items[:limit]) + (f" (+{len(items) - limit} more)" if len(items) > limit else "")
