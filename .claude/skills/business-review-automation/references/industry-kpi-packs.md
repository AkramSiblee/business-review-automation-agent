# Industry KPI packs

Selected once in Phase 0, per entity or per segment. A group can run more than one pack, for example a
SaaS segment and a hardware segment. Record the selection on the workbook's Config & Setup tab. The packs
*add* to the standard KPI dashboard in `chart-of-accounts-kpi-taxonomy.md`; they don't replace it.

**How to pick a pack:** look at the revenue account names and the bookings `Revenue Type` / `Recurring`
fields.
- Recurring > 60% of revenue → SaaS / subscription.
- Inventory and COGS present, with COGS > 50% of revenue → manufacturing / distribution.
- Store, location or SKU detail present → retail.
- Timesheet, utilization or billable data, with people cost > 50% of revenue → professional services.
- Contract assets / POC / milestones present → project-based.
- GMV and take rate present → marketplace.
- Loans, deposits and interest income as the main revenue → lending / financial services.

The selection is MEDIUM until the user confirms it once. After that, reuse it.

**Rule for every pack:** compute a KPI only if its input data was actually provided. A KPI whose data is
absent is shown as "not available – needs `<input>`". Never estimate it from a proxy without labelling it.
The `09_Operational_KPIs` tab (see `standard-input-file-templates.md`) is where non-financial inputs such
as units, stores and billable hours arrive.

---

## SaaS / subscription
- **KPIs:** ARR / MRR, net revenue retention (NRR), gross revenue retention (GRR), logo churn, CAC
  payback (months), magic number, Rule of 40, deferred revenue growth, % of revenue recurring.
- **Red flags:**
  - NRR < 100%
  - GRR < 85%
  - CAC payback > 24 months
  - deferred revenue falling while revenue is flat (leading churn indicator)
  - a single customer > 10% of ARR
- **Pitfalls:**
  - ARR from bookings is contract value, not GL revenue (playbook §1).
  - Multi-year prepaid deals inflate billings but not ARR.

## Manufacturing / distribution
- **KPIs:** gross margin by product line, inventory turns / DIO, cash conversion cycle (DSO + DIO − DPO),
  capacity utilization, scrap / yield, backlog, price vs volume vs mix bridge, freight and input-cost
  inflation pass-through.
- **Red flags:**
  - DIO up > 20% vs the trailing 3 months
  - inventory growing faster than COGS for 2+ periods (obsolescence)
  - GM down > 2 pts with input costs up (no pass-through)
  - cash conversion cycle lengthening > 10 days
- **Pitfalls:**
  - Standard-cost variances may sit outside COGS.
  - LIFO vs FIFO comparability (playbook §8).

## Retail / e-commerce
- **KPIs:** like-for-like / same-store sales growth, sales per sq ft or per store, average order value,
  conversion, traffic, gross margin after markdowns, inventory weeks of cover, shrink %, return rate,
  fulfilment cost per order.
- **Red flags:**
  - like-for-like growth negative while total growth is positive (growth bought with new stores)
  - weeks of cover up > 25%
  - return rate up > 3 pts
  - markdown-driven margin erosion
- **Pitfalls:**
  - Use the 4-4-5 / 53-week calendar where the business does (locale config).
  - Compare like-for-like weeks, not calendar months.

## Professional services / agencies
- **KPIs:** utilization (billable ÷ available hours), realization (billed ÷ standard rate), revenue per
  FTE, billable headcount mix, WIP days and unbilled receivables, contractor %, backlog / signed
  pipeline coverage.
- **Red flags:**
  - utilization < 70% or down > 5 pts
  - WIP days up > 15%
  - revenue per FTE falling while headcount grows
  - top client > 20% of revenue
- **Pitfalls:**
  - People cost is the COGS. Make sure roster-to-GL tie-outs (playbook §1) are done before computing
    margin per head.

## Project-based / construction / engineering
- **KPIs:** backlog and book-to-bill, contract margin by project, POC revenue vs billing, contract assets
  vs liabilities, cost-to-complete changes, retention receivables.
- **Red flags:**
  - margin fade on a project (estimated margin falling period over period)
  - contract assets growing faster than revenue
  - a loss-making contract without a loss provision
  - book-to-bill < 1 for 2+ periods
- **Pitfalls:**
  - Revenue depends on estimates. Flag every estimate change on a top-5 project.

## Marketplace / platform
- **KPIs:** GMV, take rate (net revenue ÷ GMV), active buyers and sellers, order frequency, contribution
  margin per order, incentives as % of GMV.
- **Red flags:**
  - take rate falling > 1 pt
  - incentives as % of GMV rising while GMV growth slows
  - a principal vs agent presentation change (playbook §7)

## Lending / financial services
- **KPIs:** net interest margin, cost of funds, loan book growth, NPL ratio, cost of risk, provision
  coverage, capital adequacy / leverage ratio, liquidity coverage, cost-to-income.
- **Red flags:**
  - NPL ratio up > 1 pt
  - provision coverage falling while NPL rises
  - capital ratio within 2 pts of the regulatory minimum
- **Pitfalls:**
  - Revenue, cost and liquidity definitions differ from a commercial business.
  - Don't run the generic gross-margin or runway KPIs. Mark them not applicable.
  - Regulatory ratios must come from the file, never be recomputed from a guess at regulatory definitions.

## Non-profit / public sector (if encountered)
- **KPIs:** restricted vs unrestricted funds, program expense ratio, months of unrestricted reserves,
  grant pipeline, cost per beneficiary.
- **Pitfalls:**
  - "Revenue vs budget" may be grant income timing. Treat restricted-fund movements as their own line.
