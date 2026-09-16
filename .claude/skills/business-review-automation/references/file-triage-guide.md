# File triage guide

Used in Phase 1 every cycle. The user drops **one file** into the resolved inputs folder — usually a
single Excel workbook whose tabs vary cycle to cycle in count, naming, and completeness, but sometimes a
docx, pdf, or pptx instead. Classify each tab (or, for a non-Excel file, each extracted table/section) by
its content, not its name alone, and assign a confidence tier per the Autonomy Doctrine in SKILL.md.
`standard-input-file-templates.md` has the concrete tab/column contract to try first; this guide is the
fallback for anything that doesn't match it exactly, including a tab that's been renamed, reordered, or
never had a literal tab-per-category layout to begin with.

## Confidence scoring for classification

- **HIGH:** column headers or structure unambiguously match a category (e.g., a column literally labeled "Account Name" with debit/credit columns; employee-level rows with "Department" and "Hire Date"). Proceed silently, log it.
- **MEDIUM:** strong circumstantial match but not a labeled confirmation (e.g., an unlabeled tab with account-like names and two numeric columns that behave like debits/credits). Proceed with the best-fit classification, log it prominently.
- **LOW:** no category fits with reasonable confidence, or two categories are equally plausible. Add to the batched question round rather than guessing.

## Category signals (English)

**P&L / general ledger / trial balance**
- Tab-name hints: "GL," "trial balance," "P&L," "income statement"
- Content hints: account names paired with debit/credit or a signed amount column; a mix of revenue and expense account names; often an account-number column

**Balance sheet**
- Tab-name hints: "balance sheet," "BS," "financial position"
- Content hints: asset/liability/equity account names; a single as-of date rather than a period range; assets should sum to liabilities + equity

**Cash / bank**
- Tab-name hints: "cash," "bank," "treasury"
- Content hints: a running balance column, transaction-level bank activity, or a beginning/ending cash summary by account

**Headcount / HRIS**
- Tab-name hints: "headcount," "HC," "roster," "HRIS export"
- Content hints: employee-level rows with department, title, start date, sometimes termination date and compensation

**Revenue / bookings / CRM export**
- Tab-name hints: "bookings," "ARR," "revenue detail," "pipeline," "CRM export"
- Content hints: customer or deal-level rows, close dates, deal amounts, sometimes renewal/expansion/churn flags

**Budget / forecast**
- Tab-name hints: "budget," "forecast," "plan"
- Content hints: same account structure as the P&L/BS framed as a plan, sometimes with a variance column already present, often spanning future periods

## Multiple tabs that map to the same category

Not uncommon — e.g. a separate GL tab per entity, or per quarter, inside the one workbook. Triage each
tab independently, classify each into its category on its own merits, and pool the rows once classified;
don't require the user to have pre-merged everything onto a single tab per category.

## Non-Excel single files (docx, pdf, pptx)

There is no "tab" to read — extract first (per `standard-input-file-templates.md`'s "Non-Excel single
files" section), then apply the same category signals above to each table or block of figures you pulled
out, exactly as if it were a tab's content. A PDF page headed "Trial Balance" with debit/credit columns
is triaged the same way a tab named `01_GL_TrialBalance` would be — same HIGH-confidence bar, same
column-meaning-over-label approach. Expect coverage to be partial (a single memo or statement export
rarely contains all eight categories) — that's a normal outcome, not a triage failure.

## Non-English content

If a tab or document's language isn't resolved by the structural signals above, consult
`references/multilingual-term-tables.md` for a Spanish/French/German/Portuguese term lookup before
falling back to LOW confidence. Once a user confirms a mapping for a given entity's format, retain that
mapping and reuse it automatically in every future cycle for that entity — the goal is to ask once per
format, never repeatedly.

## When a tab or section doesn't clearly match any category

State it plainly and include it in the batched LOW-confidence question round rather than forcing it into
the nearest bucket — it may be something the Phase 0 taxonomy needs extending to cover (e.g., a vendor
contract schedule, a customer-concentration report).

## When an expected category has no data this cycle

State plainly which section of the review will be thin or omitted as a result, and proceed with the rest
of the cycle rather than stalling. Never fabricate a plausible-looking number for a missing category.
This is the expected, common case with a single-document (non-Excel) input, and still worth noting
plainly even then rather than treated as an error condition.
