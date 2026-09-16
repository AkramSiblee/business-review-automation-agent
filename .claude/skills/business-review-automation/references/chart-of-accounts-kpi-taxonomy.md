# Default chart of accounts & KPI taxonomy (edit to fit the business)

This is a starting point, not a fixed standard. Walk through it once with the user in Phase 0, edit line items to match their actual business and accounting standard, then treat the agreed version as fixed across cycles \u2014 comparability period-to-period depends on not silently changing this.

## Income statement (P&L)

- **Revenue**
  - Recurring / subscription revenue (if applicable)
  - Product revenue
  - Services revenue
  - Other revenue
- **Cost of revenue / COGS**
- **Gross profit** (calculated)
- **Operating expenses**
  - Sales & Marketing
  - Research & Development / Product / Engineering \u2014 *note: IFRS permits capitalizing qualifying development costs; US GAAP generally expenses R&D as incurred. State which basis applies before comparing this line across entities on different standards.*
  - General & Administrative
  - Other operating expense
- **Operating income (EBIT)** (calculated)
- **Non-operating items** (interest income/expense, other income/expense, one-time items \u2014 tag one-time items explicitly)
- **Pretax income** (calculated)
- **Tax provision**
- **Net income** (calculated)

## Balance sheet

- **Assets:** cash & equivalents; accounts receivable; inventory (if applicable \u2014 *note valuation method: LIFO is permitted under US GAAP but not IFRS*); prepaid & other current assets; PP&E, net; right-of-use lease assets (material under IFRS 16 for most operating leases; may be smaller under ASC 842 depending on lease classification); intangibles & goodwill; other long-term assets
- **Liabilities:** accounts payable; accrued liabilities; deferred revenue; lease liabilities (see note above \u2014 confirm which standard drives this entity's presentation); short-term debt; long-term debt; other liabilities
- **Equity:** common stock / APIC; retained earnings; cumulative translation adjustment (if consolidating foreign entities); other equity (AOCI, treasury, etc.)

## Cash flow

- **Operating:** net income; D&A; stock-based comp; deferred tax; working-capital changes; other operating
- **Investing:** capex; acquisitions; other investing
- **Financing:** debt issuance/repayment; equity issuance/buybacks; dividends; other financing

## Revenue drivers (for the bridge in Phase 4)

- New business / new logos
- Expansion / upsell (existing customers)
- Churn / contraction (existing customers)
- Price vs. volume, where the underlying detail supports separating them
- By segment, product line, or region, if the business tracks revenue that way

## Headcount

- By department: Sales, Marketing, Engineering/Product, G&A, Customer Success/Operations, Other
- New hires, departures (voluntary/involuntary), net change, by department
- Fully-loaded cost per head, by department (or at minimum blended) \u2014 note the currency and whether cost includes statutory employer contributions, which vary significantly by country and can distort cross-entity cost-per-head comparisons if not normalized

## Standard KPI dashboard (company-wide rollup)

- Revenue growth: MoM, QoQ, YoY, and vs. budget
- Gross margin % and operating margin %, trended
- Opex as % of revenue, by category (S&M, R&D, G&A)
- Headcount: total, by department, and cost per head
- Cash balance, monthly burn rate, runway in months
- Working capital: DSO / DPO (if the business and data support it)
- Budget variance %: revenue, opex, net income
- Customer/revenue concentration (if data available)
- Rule-of-40-style growth+margin composite (optional \u2014 only if the business model is subscription/SaaS-like)

## Notes for whoever edits this

- Delete anything that doesn't apply rather than leaving zeros everywhere.
- Add anything industry- or geography-specific that's missing (e.g., bookings/ARR for SaaS, same-store sales for retail, statutory reserve requirements for a regulated entity).
- **State the accounting standard(s) in use next to this taxonomy once it's confirmed** \u2014 if consolidating entities on different standards, note it per entity so the divergence points flagged above can be tracked consistently.
- Once edited, this becomes the fixed target every uploaded file's line items get mapped into (Phase 1) \u2014 changing it mid-year breaks comparability, so batch edits at year-end/annual review time where possible.
