# Red-flag thresholds (materiality-scaled, not fixed-dollar)

Confirm once in Phase 0; reuse every cycle. All thresholds here are expressed as percentages or standard deviations rather than fixed currency amounts \u2014 a fixed dollar/euro/rupee threshold breaks the moment this is used for a business of a different size, which defeats the purpose for a tool meant to serve users of any scale, anywhere. Tune sensitivity if the business consistently trips a threshold on normal, explainable seasonality.

## Fixed materiality-scaled defaults

| Metric | Default trigger | Notes |
|---|---|---|
| Revenue vs. budget | \|variance\| > 10% of budgeted revenue | Tighten to 5% for a mature, predictable business; loosen for an early-stage one |
| Opex line vs. budget | \|variance\| > 15% of that category's budget, for any single category | Category-level, not just total \u2014 a big miss in one line can hide inside an on-budget total |
| Gross margin | Move of > 2 percentage points MoM or > 3 points QoQ | Points, not percent-of-percent, and independent of absolute revenue scale |
| Headcount | Net change > 5% in a single month, or any department > 10% QoQ | Also flag any month with more departures than hires in a department expected to be growing |
| Cash runway | Falls below 12 months at current burn | Escalate further below 6 months |
| Monthly burn | Increases > 20% MoM without a corresponding, explained revenue or headcount driver | Expressed as a rate of change, not an absolute currency amount |
| DSO (days sales outstanding) | Increases > 15% vs. trailing 3-month average | Currency-agnostic by construction |
| Customer/revenue concentration | Top 1 customer > 15% of revenue, or top 5 > 40% | Only where customer-level data is available |
| Prior-cycle flag, unresolved | Any item flagged last cycle with no improvement | Escalate rather than re-flag identically |
| Covenant headroom | < 15% headroom on any covenant, or any breach | Per the facility's own definitions. A breach is always a headline item. |
| Going-concern indicators | Any one present (playbook §5.4) | 2+ present → slide 1 / memo executive summary. Never conclude, only list. |
| Debt maturity | Any facility maturing within 12 months without refinancing evidence | Also check that maturities fall inside the runway window |
| AR vs revenue growth | AR growth exceeds revenue growth by > 15 pts (same period, YoY or QoQ) | Cutoff / collection risk signal |
| Inventory vs COGS growth | Inventory growth exceeds COGS growth for 2+ consecutive periods | Obsolescence signal |
| DPO stretch | DPO up > 15% vs the trailing 3-month average | Treat the cash benefit as non-recurring |
| Deferred revenue | Falls > 10% QoQ while recognized revenue is flat or up | Leading churn indicator (subscription businesses) |
| Effective tax rate | Moves > 10 pts vs the prior period or the statutory rate | Needs a stated driver |
| "One-time" items recurring | Same add-back category in 3+ of the last 6 periods | Show Adjusted EBITDA with and without it |
| Capitalization shift | Capitalized cost up > 25% while the related opex line falls | Margin improvement may be accounting, not operations |
| FX step change | Any functional currency's spot rate moves > 10% in one month | Devaluation / hyperinflation watch. Split the variance at constant currency. |
| Organic vs reported gap | Reported growth exceeds organic growth by > 5 pts | Growth is acquisition-driven. Lead with the organic figure. |
| Industry pack thresholds | Per `industry-kpi-packs.md` | Only for the pack(s) selected in Phase 0 |

## Statistical (trend-based) detection \u2014 the complement to fixed thresholds

Fixed percentage thresholds have a blind spot: they under-flag on metrics that are normally very stable (where even a 6% move is unusual and meaningful) and over-flag on metrics that are normally volatile (where a 12% move is unremarkable). Once roughly 6+ periods of history exist for a metric, compute the trailing mean and standard deviation and flag any value more than **~2 standard deviations** from trend, in addition to the fixed thresholds above. Report both triggers when they disagree \u2014 a metric that trips the statistical check but not the fixed one (or vice versa) is itself informative about how the business's normal variability is changing.

Before ~6 periods of history exist (e.g., the first two quarters of using this system), rely on the fixed thresholds alone and say so explicitly \u2014 don't compute a standard deviation off two or three data points and present it as meaningful.

## Scaling movers by relative size, not absolute size

When ranking the "top movers" for the insight-generation phase, rank by the variance's size **relative to its own base** (e.g., % of revenue, % of the relevant budget line) rather than by raw currency magnitude \u2014 otherwise a large business's routine rounding noise crowds out a smaller business's genuinely significant swings, and the reverse.

## How thresholds are used

- **Phase 2 (validation):** anything tripping a threshold or the statistical check gets a "confirm before proceeding" flag if it's also large enough to trigger the Autonomy Doctrine's materiality override; otherwise it proceeds and is logged for the red-flag scan.
- **Phase 5 (insight generation):** anything tripping either check is guaranteed a place in the top-movers/red-flags list regardless of how it would otherwise rank by raw size.
- **Phase 8 (QA):** confirm every tripped threshold this cycle appears in both the memo's "risks & watch items" section and the deck's red-flags slide.
