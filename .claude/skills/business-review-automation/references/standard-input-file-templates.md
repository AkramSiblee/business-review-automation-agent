# Standard input workbook contract

The user provides **one file per cycle**, dropped into the resolved inputs folder. The expected and
best-supported case is a **single Excel workbook with multiple tabs** — one tab per data category below.
If a docx, pdf, or pptx arrives instead, see "Non-Excel single files" at the end of this document; the
same target categories and column contract still apply, just extracted rather than read directly.

This is the concrete tab/column contract that Phase 1 (File Triage) should try to match first, before
falling back to the general heuristics in `file-triage-guide.md`. If a user's workbook already uses
these tab names and columns, classification confidence is HIGH by construction and mapping is nearly
mechanical. If it doesn't match exactly (different tab names, different column order, extra columns),
fall back to content-based triage — never reject a tab, or the file as a whole, just because it doesn't
match this template verbatim. The sample workbook in `sample-data/sample-test-data/Business_Review_Sample_Input_FY2025.xlsx`
uses this exact tab layout and is the easiest way to see the contract in a real file — copy it as a
starting template if the user wants a blank version to fill in.

## Setup tab (one-time, or revisited on structural change)

**`00_Company_Config`** — one row per entity/subsidiary
`Entity Code, Entity Name, Country, Functional Currency, Accounting Standard, Fiscal Year Start Month, Ownership %, Consolidation Method, Reporting Currency`

## Recurring tabs (every cycle)

**`01_GL_TrialBalance`** — covers P&L and balance sheet accounts in one export
`Entity Code, Period, Account Code, Account Name, Account Type, Department, Debit, Credit, Currency, One-Time Item, Notes`
- `Account Type` is the routing column: Revenue / COGS / OpEx-S&M / OpEx-R&D / OpEx-G&A / Other Income / Other Expense / Asset / Liability / Equity.
- If the user only has separate Income Statement and Balance Sheet data rather than a full trial balance, that's fine — treat each as this same schema with `Account Type` inferred from which section it came from, whether that's two tabs or one, rather than requiring them to reformat into a single trial balance.

**`02_CashBank_Position`**
`Entity Code, Period, Bank/Account Name, Account Number (last 4), Currency, Beginning Balance, Ending Balance, Restricted Cash, Notes`

**`03_Headcount`** — one row per employee
`Entity Code, Employee ID, Department, Job Title, Employment Status, Hire Date, Termination Date, Termination Type, FTE %, Monthly Fully-Loaded Cost, Currency, Country/Location`

**`04_Revenue_Bookings`** — one row per deal/contract line
`Entity Code, Period/Transaction Date, Customer ID/Name, Deal/Contract ID, Revenue Type, Product/Segment, Region, Amount, Currency, Contract Start Date, Contract End Date, Recurring`
- `Revenue Type`: New Business / Expansion / Renewal / Contraction / Churn — this is what feeds the Phase 4 revenue-driver bridge directly, so a tab missing this column meaningfully weakens that analysis (flag it, don't silently substitute a blended growth number).

**`05_Budget_Forecast`** — same account taxonomy as the trial balance, framed as plan
`Entity Code, Fiscal Year, Period, Account Code, Account Name, Account Type, Budgeted Amount, Currency, Version`
- `Version`: Original Budget / Reforecast / Latest Forecast — keep all versions distinguishable rather than overwriting the original budget with the latest reforecast, since "actual vs. original budget" and "actual vs. latest forecast" are both meaningful and different questions.

**`06_FX_Rates`** — only needed when consolidating more than one currency
`Period, From Currency, To Currency, Rate Type, Rate, Source`
- `Rate Type`: Period-End Spot / Period-Average, matching the current-rate consolidation method in `multi-currency-multi-entity-consolidation.md`.

**`07_Intercompany_Detail`** — only needed when consolidating more than one entity
`Period, Entity (From), Entity (To), Transaction Type, Amount, Currency, Description`
- If this tab isn't provided for a multi-entity consolidation, Phase 3's intercompany elimination step cannot fully run — this must produce an explicit disclosure in the deliverables, not a silently under-eliminated consolidated number.

## Optional tabs (provide when they apply — never flagged as missing when absent)

**`08_Debt_Covenants`** — one row per facility × covenant × period
`Entity Code, Period, Facility Name, Facility Type, Limit, Drawn, Currency, Maturity Date, Interest Rate, Covenant Name, Covenant Definition, Threshold, Comparison (>= / <=), Reported Value`
- `Reported Value` is the company's own covenant calculation, if it has one. Phase 4 also recomputes it
  from the GL per the definition, and flags any difference.

**`09_Operational_KPIs`** — one row per entity × period × metric (long format)
`Entity Code, Period, Segment, Metric, Value, Unit, Notes`
- Non-financial inputs for the industry KPI packs (`industry-kpi-packs.md`): units shipped, stores,
  sq ft, billable / available hours, GMV, active customers, ARR, NPL, and so on. Use metric names as the
  business uses them, and map them to pack KPIs in Phase 0 once.

**`10_Adjustments_Events`** — free-form register of known events
`Entity Code, Period, Event Type, Description, Amount, Currency, One-Time (Y/N)`
- `Event Type`: Acquisition / Disposal / Ownership Change / Restructuring / Impairment / Litigation /
  Restatement / Refinancing / Other. This is the most direct way for a user to tell the review about a
  complex event, rather than leaving it to be inferred from the ledger.

## Handling deviations from this contract

- **Extra tabs:** keep and log them; they may feed future taxonomy extensions even if unused this cycle.
- **Missing optional columns** (e.g., `Department` in the trial balance tab): proceed without that cut of the analysis and say so, rather than blocking the cycle.
- **Missing optional tabs** (e.g., no `06_FX_Rates` tab in a single-entity, single-currency file): that's expected and not a gap — only flag it as a gap when the file's own content implies it's needed (multiple currencies or entities appear elsewhere in the workbook) but the tab is absent.
- **Different tab names, column order, or naming:** classify by header meaning and tab content (using the multilingual/synonym recognition in `file-triage-guide.md`), not by tab position or exact name — these are the target schema, not a rigid layout the user must conform to exactly. A tab named "P&L" or "Trial Balance" with the right column shapes should be recognized as `01_GL_TrialBalance` just as confidently as one already named that.
- **A completely different export layout within the one workbook** (e.g., a native QuickBooks or Xero export pasted onto a tab as-is): fall back fully to `file-triage-guide.md`'s content-based classification; treat successful mapping to this schema as the goal, not a precondition for accepting the tab.
- **More than 8 tabs, or category data split across several tabs** (e.g., separate GL tabs per entity or per quarter): each tab is still triaged independently and rows are pooled into the matching category — one tab per category is the common case, not a hard limit.

## Non-Excel single files

The user may drop a `.docx`, `.pdf`, or `.pptx` instead of an Excel workbook — a scanned or exported
financial statement, a board memo with embedded tables, a PDF trial balance, and so on. Treat this the
same as an Excel file that needs triage, with two differences:

1. **Extraction comes first.** Read the document in full (with the relevant `pdf`/`docx`/`pptx` skill)
   and pull out every table, figure, and labeled line item before attempting classification — a table
   embedded in a Word doc or a PDF export is still structured data, just not in worksheet form.
2. **Map extracted content into the same tab categories above**, by meaning, exactly as you would map a
   mis-named Excel tab. A PDF trial balance's rows still map to `01_GL_TrialBalance`'s columns even
   though there was never a literal `01_GL_TrialBalance` tab; an org chart or headcount table in a PPTX
   maps to `03_Headcount`.

Confidence tiers apply the same way: a PDF table with clear column headers ("Account," "Debit," "Credit")
is HIGH confidence; a narrative paragraph that states a figure in prose ("cash ended the quarter at
$2.1M") is at best MEDIUM (the number is stated but its full context — currency, entity, exact period
cutoff — may need inference); a scanned image with unclear or unreadable numbers is LOW confidence and
belongs in the batched question round rather than a guessed transcription. Because a single document
of this kind will rarely cover every category the way a full multi-tab workbook can, it's normal and
expected for several categories to come back "not present this cycle" — state that plainly in the File
Log rather than treating it as an error, per the doctrine in `file-triage-guide.md`.
