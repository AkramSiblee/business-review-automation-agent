# Locale & accounting standards configuration

Confirm this once per business (or once per entity, if consolidating several) in Phase 0. Getting this wrong silently is the single biggest rigor risk for a tool used by people in different countries \u2014 a misread decimal separator or an unstated accounting standard can be off by orders of magnitude or misstate margins without ever throwing an error.

## 1. Reporting currency & FX

- Confirm the **reporting currency** (the currency all consolidated outputs will be shown in) explicitly \u2014 never infer it from the first file's apparent currency.
- If any entity's functional currency differs from the reporting currency, confirm the **FX methodology**: current-rate method is the default (balance sheet translated at period-end spot rate, income statement at period-average rate, translation adjustment to equity/OCI) unless the business specifies otherwise (e.g., temporal method for an entity in a hyperinflationary economy).
- Confirm the **FX rate source** (e.g., a specific rate the business already uses for statutory reporting) so the automation doesn't introduce a rate discrepancy against the business's own books.
- State the exact rates used, by entity and period, in the workbook \u2014 never bury this inside a single blended number.

## 2. Fiscal calendar

- Confirm whether the business runs a **calendar fiscal year** (Jan-Dec) or a different one (common alternatives: Apr-Mar, Jul-Jun, Oct-Sep, or a 4-4-5 retail calendar).
- Confirm whether "Q1," "Q2," etc. refer to fiscal or calendar quarters \u2014 these are easy to silently conflate and will misalign every period-over-period comparison if assumed wrong.
- If the business uses a 4-4-5 or 4-4-4 calendar, confirm which weeks fall in which "month" for period-cutoff purposes.

## 3. Accounting standard

- Confirm the standard in use: US GAAP, IFRS, or a local GAAP variant (e.g., Ind AS, UK GAAP, local statutory basis).
- Where entities on different standards are being rolled up, flag the areas most likely to cause a distorted combined view without adjustment:
  - **Leases:** IFRS 16 capitalizes nearly all leases on-balance-sheet; ASC 842 (US GAAP) still distinguishes finance vs. operating lease expense presentation. A mixed rollup can look inconsistent in EBITDA and asset/liability figures for no operational reason.
  - **Revenue recognition timing:** generally converged (ASC 606 / IFRS 15) but industry-specific application can still differ \u2014 note if any entity uses a materially different recognition pattern.
  - **R&D / development costs:** IFRS permits capitalizing qualifying development costs; US GAAP generally expenses R&D as incurred. This alone can swing opex and margin comparisons between entities on different standards.
  - **Inventory valuation:** LIFO is permitted under US GAAP but not IFRS \u2014 note the method per entity if inventory is material.
- This list is not exhaustive \u2014 the point is to flag known divergence areas explicitly rather than present a blended multi-standard number as if it were prepared on one consistent basis.

## 4. Number format detection (decimal & thousands separators)

Do not assume a convention \u2014 detect it from the file, and fall back to configuration only when the file itself is ambiguous.

- **Detection heuristic:** scan a sample of numeric strings in the file. A pattern like `1.234.567,89` (period as thousands separator, comma as decimal) indicates the common European/Latin American convention; `1,234,567.89` indicates the common US/UK/most-of-Asia convention. Internal consistency within a single file is a strong signal \u2014 a file won't normally mix conventions.
- **Ambiguous cases** (e.g., all values are small enough that no thousands separator appears, like `1234,56` vs `1234.56`): fall back to the Phase 0 reporting-currency convention and log it as a MEDIUM-confidence assumption, not a silent default.
- **Unit scale:** check headers and notes for an explicit scale indicator ("$000s," "in millions," "\u20b9 in lakhs," "\u20b9 in crores"). If absent and the numbers look implausibly small or large for the line item (e.g., "Revenue: 45" with no other context), flag rather than guess which scale was intended.

## 5. Date format detection

- **Detection heuristic:** if any date's first or second numeric component exceeds 12, that resolves day-vs-month ordering for the whole file (a file is internally consistent). E.g., `15/03/2026` must be DD/MM because there's no 15th month.
- **Genuinely ambiguous dates** (both components \u2264 12, e.g., `03/04/2026`): cross-reference against the known reporting period being processed \u2014 if this is the March cycle, a date reading as March is the more defensible MEDIUM-confidence assumption. If the file's own date format can't be resolved this way at all (no unambiguous date anywhere in it), this becomes a LOW-confidence item for the batched question round \u2014 ask once for that file's convention and reuse the answer for every date in it.
- **Always normalize dates to an unambiguous format internally and in every output** (e.g., "31 Mar 2026" or ISO 8601 `2026-03-31`) \u2014 never pass through a bare `03/04/2026`-style date into a deliverable, since it re-introduces the same ambiguity for the reader that was just resolved for the parser.

## 6. Output formatting defaults

Unless the business specifies otherwise, default to: reporting currency shown as an ISO code alongside the symbol (e.g., "USD 1,234,567"), dates as "DD Mon YYYY," and numbers formatted per the reporting currency's common convention (comma thousands / period decimal for USD, GBP, most of Asia; period thousands / comma decimal for most of continental Europe and Latin America). State the convention used once on each deliverable's cover or summary section so a reader anywhere can orient immediately.
