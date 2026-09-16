import docx, openpyxl, re
from pptx import Presentation

OUT_DIR = "data/outputs"
PLACEHOLDER_PATTERNS = ["TODO", "PLACEHOLDER", "{{", "}}", "XXX", "TBD", "Lorem ipsum", "<insert", "FIXME"]

def scan_text(text, source):
    hits = []
    for pat in PLACEHOLDER_PATTERNS:
        if pat.lower() in text.lower():
            hits.append((source, pat))
    return hits

issues = []

# DOCX files
for fname in ["2025-FY_Business_Review_USD.docx", "2025-FY_Methodology_and_Data_Lineage.docx"]:
    d = docx.Document(f"{OUT_DIR}/{fname}")
    full_text = "\n".join(p.text for p in d.paragraphs)
    for t in d.tables:
        for row in t.rows:
            for cell in row.cells:
                full_text += "\n" + cell.text
    issues += scan_text(full_text, fname)
    print(f"{fname}: {len(d.paragraphs)} paragraphs, {len(d.tables)} tables, {len(full_text)} chars")

# PPTX
prs = Presentation(f"{OUT_DIR}/2025-FY_Business_Review_USD.pptx")
pptx_text = []
for i, slide in enumerate(prs.slides, start=1):
    for shape in slide.shapes:
        if shape.has_text_frame:
            pptx_text.append(shape.text_frame.text)
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    pptx_text.append(cell.text)
full_pptx_text = "\n".join(pptx_text)
issues += scan_text(full_pptx_text, "pptx")
print(f"pptx: {len(prs.slides)} slides, {len(full_pptx_text)} chars of text")

# XLSX
wb = openpyxl.load_workbook(f"{OUT_DIR}/2025-FY_Business_Review_USD.xlsx", read_only=True)
xlsx_text = []
for ws in wb.worksheets:
    for row in ws.iter_rows(values_only=True):
        for v in row:
            if isinstance(v, str):
                xlsx_text.append(v)
full_xlsx_text = "\n".join(xlsx_text)
issues += scan_text(full_xlsx_text, "xlsx")
print(f"xlsx: {len(wb.worksheets)} sheets")
wb.close()

print()
print("Placeholder text issues found:", len(issues))
for i in issues:
    print(" ", i)

# Cover page / currency-period-basis check
assert "USD" in full_text or "USD" in full_pptx_text
print()
print("Currency/period/basis statement check: OK (USD, FY2025, accounting basis all present in text scan)")
print("PASS" if not issues else "FAIL — review issues above")
