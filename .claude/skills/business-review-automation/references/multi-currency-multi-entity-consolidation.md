# Multi-currency & multi-entity consolidation

Applies whenever the review rolls up more than one entity or currency. Confirm the methodology once in Phase 0; apply it identically every cycle.

## Translation methodology

**Current-rate method (default):**
- Balance sheet: translate at the period-end spot rate.
- Income statement: translate at the period-average rate for the period (or a weighted rate if the business already computes one for statutory purposes \u2014 match their existing practice rather than introducing a new methodology).
- The resulting translation gain/loss goes to a cumulative translation adjustment within equity/OCI, not through the income statement. Expect a nonzero CTA movement most periods, not just at year-end — spot and average rates rarely match closing-to-closing, so a period with a CTA swing but no unusual underlying activity is normal, not a data error to chase.

**Temporal method (use only if specified, e.g., for an entity in a hyperinflationary economy or one whose functional currency differs from its local currency):**
- Monetary assets/liabilities at the closing rate; non-monetary items at historical rates; translation differences flow through the income statement.
- This is a meaningfully different result from the current-rate method \u2014 do not switch between the two methods across cycles without an explicit, logged reason.

## What to disclose every cycle

- The exact rate(s) used per entity and period (spot and average, as applicable), sourced consistently with the business's own statutory reporting where possible.
- Which entities were translated under which method, if more than one method is in use.
- The cumulative translation adjustment balance and its period change, in the balance sheet / equity section \u2014 don't let it disappear into a generic "other" line.

## Intercompany eliminations

- Identify intercompany revenue, payables/receivables, and any intercompany profit-in-inventory where the source files make this possible.
- Eliminate what can be identified; **explicitly disclose what could not be eliminated** because the underlying intercompany detail wasn't provided this cycle \u2014 a consolidated total that silently includes unreconciled intercompany balances is a materially misleading number, and the honest response is disclosure, not a false sense of precision.

## Consolidation method per entity (from `00_Company_Config`)

| `Consolidation Method` / `Ownership %` | Treatment | What group KPIs include |
|---|---|---|
| Full, 100% | Line-by-line | Everything |
| Full, < 100% (subsidiary with NCI) | Line-by-line at 100%. NCI share of net income shown below net income ("attributable to: owners of the parent / non-controlling interests"). NCI share of net assets as its own equity line. | 100% of revenue, opex, headcount and cash. Per-share or "owner" metrics use the parent-attributable figures. |
| Equity method (associate / JV, typically 20–50%) | One P&L line, "share of profit of associates" = ownership % × investee net income (after IC profit elimination). One balance-sheet line, the investment, rolled forward by share of profit less dividends received. | None of its revenue, opex, headcount or cash. Show it as its own memo line. |
| Proportionate (joint operation, or a local-GAAP election) | Ownership % of each line | That % of each line |
| Not consolidated / fair value (< 20% investment) | Investment at the value given; dividends as other income | None |

If `Ownership %` is missing for an entity, assume 100% / full only when nothing in the data suggests
otherwise (MEDIUM). If ownership is 20–50% and the method is blank, it's LOW: the choice changes revenue
materially.

An ownership change during the period (a step acquisition, or selling down a stake) is a structural
change. Log it with its date, and apply the old method before the date and the new method after.

## Acquisitions: purchase-price allocation & goodwill

- Consolidate from the acquisition (control) date only. Never back-fill pre-acquisition months into
  group actuals. Pre-acquisition months may appear in a separate **pro forma** view, labelled as such.
- Goodwill = consideration (plus NCI, plus the fair value of any previously held interest) − fair value
  of the identifiable net assets.
- With no PPA in the file, show the excess as **provisional goodwill** and disclose it (MEDIUM).
- Acquired intangibles amortize; goodwill doesn't (IFRS / US GAAP public companies). Don't estimate
  amortization of intangibles that haven't been identified.
- Deal costs are expensed and treated as one-time. Earn-outs / contingent consideration are a liability;
  fair-value changes in them go through P&L and are also one-time.

## Disposals & discontinued operations

- The gain or loss on disposal = proceeds − carrying amount of net assets disposed (including allocated
  goodwill) ± recycled CTA for a foreign operation.
- If the unit is a major line of business or a geographic area (IFRS 5 / ASC 205-20): present its results
  as one line, net of tax, after continuing net income, and restate the comparatives. Compute all KPIs and
  thresholds on continuing operations.

## Entities with different fiscal year-ends

Align them to the parent's fiscal calendar using monthly data where available. If an entity reports
only on its own quarter-ends, use the nearest quarter (lag of 3 months or less, IFRS 10 / ASC 810
practice) and disclose the lag. Adjust for material events in the gap if they are known.

## Hyperinflationary economies

- **IFRS (IAS 29):** restate the entity's non-monetary items and P&L with a general price index *before*
  translating. Then translate *everything*, P&L included, at the closing rate.
- **US GAAP (ASC 830):** remeasure into the reporting currency as if it were the functional currency
  (temporal method), with gains and losses in P&L.
- The index and the classification come from the user, or from the latest IPTF / central-bank data the
  user supplies. Never assume them. Without them, present the entity un-restated with an explicit
  caveat, and treat it as a LOW item.

## Constant currency

For every multi-currency review, also translate the current period at the comparison period's average
rates (prior year, or the budget rate when the budget was set at a fixed rate). Show:

`reported change = constant-currency change + FX translation impact`

The FX translation impact is a presentation effect, not operating performance. Transaction FX gains and
losses (in P&L) are separate again.

## Entity changes mid-period

- **Acquisition:** flag the entity and the date it entered the consolidation; show both an as-reported (with the new entity) and an organic/like-for-like (prior-entity-set only) view of growth and margin trends for at least the cycle in which it occurred, so period-over-period comparisons aren't silently distorted by the addition.
- **Divestiture:** mirror the same treatment in reverse \u2014 show continuing-operations-only trends alongside the as-reported total for the transition period.
- **New subsidiary formed internally (not an acquisition):** note it exists but generally does not need an organic-vs-reported split unless it's large enough to be individually material.

## When consolidation isn't reliable enough to present with confidence

If entity-level data quality, missing intercompany detail, or unclear translation source rates mean the consolidated figure has real uncertainty, say so directly in the deliverable rather than presenting a precise-looking number \u2014 a stated range or an explicit caveat is more useful to a global board than false precision.
