# Complex scenarios playbook

Read this every cycle in Phase 1 (sections 1–2), Phase 2 (sections 3–9), and Phase 5 (section 10).
The simple case (one or a few wholly-owned entities, clean monthly activity data, a healthy balance
sheet) needs none of it. Real files rarely stay simple, and each scenario below is something that
otherwise produces a confident-looking, wrong review.

Every entry follows the same shape, so the autonomy doctrine applies mechanically:

- **Detect:** the signal in the data. `profile_inputs.py` already checks for the ones marked ⚙, so
  read its Section 4 first and don't re-derive them.
- **Default:** what to do when nobody answers. The aim is a defensible default, so that very few
  items have to stop as LOW.
- **Tier:** the starting confidence. The materiality override can raise it.
- **Disclose:** what has to appear in the deliverables, and where.

A scenario that is detected but not handled is a defect. Log it in the Assumptions & Judgment Log even
when the default was applied silently.

---

## 1. Source-of-truth hierarchy (when two sources disagree)

Most "LOW" items in practice are two inputs giving different numbers for the same thing. Resolve them
with this fixed precedence, and stop only when the gap is material *and* the hierarchy doesn't settle it.

| Question | Authoritative source | Supporting / driver evidence only | Notes |
|---|---|---|---|
| P&L amounts (revenue, cost, opex) | GL / trial balance | Bookings, roster, sub-ledgers | The GL is what is audited. Others explain *why*, never *how much*. |
| Cash balance | Bank position | GL cash account | If the GL differs from the bank, report the bank balance and flag the reconciling item. |
| Restricted cash | Bank position `Restricted Cash` / notes | GL | Always excluded from liquidity and runway. |
| Headcount (heads, hires, leavers) | HR roster | GL people cost | Cost per head = GL people cost ÷ roster FTE, with the definition stated. |
| People cost | GL (salary + benefits + payroll tax + people-related COGS) | Roster `Monthly Fully-Loaded Cost` | Roster cost often includes employer on-costs, bonus accruals or a planning rate. Disclose a gap > 10%, don't fix it. |
| Customer movements (new / expansion / churn) | Bookings / CRM | GL revenue lines | Bookings are contract value, the GL is recognized revenue. They tie only over the contract life, never in-month. |
| FX rates | `06_FX_Rates` (the company's treasury source) | Market rates you know of | Never substitute a rate from memory. A missing rate is a gap, not a lookup. |
| Intercompany | Both sides' GL + `07_Intercompany_Detail` | — | When the sides disagree, eliminate the lower amount and disclose the difference as unreconciled. |
| Budget | `Original Budget` version | Reforecast / Latest Forecast | Show both comparisons when both exist. Never overwrite the original. |
| Narrative notes in the file | The figure in the data | The note's explanation | A note that contradicts its own figure goes to the batched round. |

Apply the hierarchy as HIGH when the gap is immaterial, and as MEDIUM with disclosure when it's
material. Escalate to LOW only when the losing source would change a red flag or a headline KPI *and*
there's a credible reason to think the authoritative source is the one that's wrong.

---

## 2. Data-shape traps (Phase 1, before any figure is used)

**2.1 YTD-cumulative figures labelled as monthly** ⚙
- Detect: P&L account balances rise in absolute terms month after month for most accounts, and the
  retained-earnings rollforward fails until you difference the months.
- Default: convert to period activity (month *n* minus month *n−1*, restarting at fiscal year start) and
  re-run the RE rollforward. If the rollforward then passes, it's HIGH.
- Disclose: in the File Log.

**2.2 Subtotal or total rows inside a detail export** ⚙
- Detect: account names containing "Total", "Subtotal", "Net", "Gross Profit"; rows with no account
  code but with amounts; debits ≠ credits by exactly 2× one section.
- Default: drop the subtotal rows and re-check that debits = credits. HIGH if it then balances, MEDIUM
  otherwise.

**2.3 Months laid out as columns (wide / pivot layout)** ⚙
- Detect: three or more headers that parse as periods.
- Default: unpivot to long format (one row per entity × account × period) before classifying. MEDIUM.

**2.4 Single signed Amount column instead of Debit/Credit** ⚙
- Detect: an `Amount` / `Balance` / `Value` column with no Debit/Credit pair.
- Default: infer the sign convention from revenue. If revenue is mostly negative, the file is
  credit-negative (the usual ERP convention). Normalize, then confirm that assets = liabilities + equity.
  HIGH if the balance check passes.

**2.5 Unit scale stated in headers or titles** ⚙
- Detect: `(000)`, `'000`, `in thousands`, `USD m`, `lakhs`, `crores`, `₹ Cr`.
- Default: scale to actual units. If different tabs use different scales, normalize every tab and log
  each one.
- Materiality: getting a 1,000× error wrong flips every flag, so a scale inferred only from magnitude
  (no label) is LOW.

**2.6 Excel error values in the input** ⚙
- Detect: `#REF!`, `#DIV/0!`, `#N/A`, `#VALUE!`, `#NAME?`.
- Default: treat the cell as missing, never as zero. Material errors are LOW, because the source
  workbook is broken upstream.

**2.7 Missing months inside the range** ⚙
- Detect: the GL period sequence has holes.
- Default: never interpolate. Compute the comparisons that are still valid and mark the rest "n/a – period
  missing". MEDIUM.

**2.8 Mixed currencies within one entity's rows** ⚙
- Detect: one entity has more than one `Currency` value in a tab.
- Default: translate each row from its own currency to the functional currency at the same-period rate.
  Expected for multiple bank accounts; for GL rows it's MEDIUM.

**2.9 Chart-of-accounts change mid-year** ⚙
- Detect: an account set that changes between periods for the same entity (accounts appear or disappear).
- Default: when amounts move from an old code to a new one in the same period, map the new code to the
  same taxonomy line and log the remap. An account that disappears with no replacement goes to the
  continuity check (Phase 2).

**2.10 Late, adjusting or post-close entries**
- Detect: a period labelled 13 or "ADJ", a note saying "audit adjustment", a prior period that changed
  versus the history workbook.
- Default: fold period 13 / ADJ into the last fiscal month and log it. A changed prior period is handled
  as a restatement (SKILL Phase 2): keep both versions.

**2.11 Hidden tabs, or tabs that look like scratch work** ⚙
- Detect: a hidden sheet, or tab names like "Sheet1", "old", "copy of", "draft", "backup".
- Default: profile the tab but exclude it unless it is the only source for a category. Log its fate either
  way. A hidden tab that duplicates a visible one is the classic double-count.

**2.12 Future-dated periods in actuals** ⚙
- Detect: actuals periods later than the latest closed month (or later than today).
- Default: treat them as forecast leakage. Exclude them from actuals and log. MEDIUM.

---

## 3. Group structure & consolidation scope

Read `multi-currency-multi-entity-consolidation.md` for the mechanics. This section covers detection
and defaults.

**3.1 Partial ownership (non-controlling interest)** ⚙
- Detect: `Ownership %` < 100 with `Consolidation Method` = Full.
- Default: consolidate 100% of the lines and attribute NCI share of net income and equity. Show "net
  income attributable to parent" and the NCI on the face of the statements. HIGH when ownership is given.

**3.2 Equity-method investee / associate / JV** ⚙
- Detect: method = Equity / Associate / JV, or ownership of roughly 20–50%.
- Default: no line-by-line consolidation. Record one line, share of profit (ownership % × investee NI),
  and one investment balance. Exclude its revenue, opex and headcount from group KPIs. Exclude its cash
  from group liquidity.
- Note: if its full GL was provided, it's still not consolidated. Use it only for the one-line
  calculation.

**3.3 Proportionate consolidation** (joint operations, or a configured local-GAAP choice)
- Default: include ownership % of each line. State the method on the FX & Consolidation tab.

**3.4 Acquisition during the period**
- Detect: an entity whose first GL period is after the fiscal year start, an acquisition note, a goodwill or
  intangibles step-up, or an investing outflow labelled acquisition.
- Default: consolidate from the acquisition date only, and show both as-reported and organic
  (like-for-like) growth.
- Purchase-price allocation: if one is given, use it. If not, show the excess of consideration over net
  assets as provisional goodwill. Mark it provisional and MEDIUM (it's normal for a PPA to stay provisional
  for up to 12 months).
- Deal costs are expensed and treated as one-time.

**3.5 Disposal / discontinued operations**
- Detect: an entity or segment whose GL stops mid-year, a sale note, a gain/loss on disposal.
- Default: if the unit is a major line of business or a geographic area, present it as discontinued: one
  line, net of tax, below continuing net income. Restate the comparatives the same way. All KPIs,
  thresholds and trends are computed on continuing operations. MEDIUM. The "major line" test is
  judgment, so log it.

**3.6 Entities with different fiscal year-ends** ⚙
- Detect: `Fiscal Year Start Month` differs between entities in config.
- Default: align them to the parent's calendar using monthly data. If only quarterly data exists for the
  odd entity, use its nearest quarter (lag of 3 months or less) and disclose the lag. MEDIUM.

**3.7 Dormant, holding or shell entities**
- Detect: an entity with balance-sheet rows only and no P&L.
- Default: consolidate the balance sheet. These are expected to have no budget, bookings or headcount,
  so don't raise "entity missing" gaps for them. Log once.

**3.8 Common-control reorganization (internal transfer of a subsidiary)**
- Default: no goodwill and no gain. Use carrying values (predecessor basis). Note it in the File Log.

---

## 4. Currency stress

**4.1 Hyperinflationary economy**
- Detect: three-year cumulative inflation approaching 100%, a functional currency on the IPTF/SEC
  watch-list as of the period (verify the current list; it changes), or a monthly spot move > 10% ⚙.
- Default: under IFRS, IAS 29 restatement applies. Under US GAAP (ASC 830), remeasure as if the
  reporting currency were functional.
- Don't build an IAS 29 index restatement from assumptions. Without a general price index in the file
  it's LOW. Ask for the index, or present that entity's figures with an explicit "not restated for
  hyperinflation" caveat.

**4.2 Devaluation or step change in a rate** ⚙
- Detect: spot moves > 10% in one month.
- Default: use the rates given. Split the variance into FX vs operational with constant currency (4.4).
  Flag the translation hit to equity (CTA) separately from the transaction losses in P&L.

**4.3 Missing or partial FX rates** ⚙
- Default: never carry forward or fill rates from memory. A missing average rate for one month is
  MEDIUM: use that month's spot as a proxy and disclose. A missing spot rate for the closing balance sheet
  is LOW.

**4.4 Constant-currency variance** (always do this when multi-currency)
- Recompute current-period results at the comparison period's rates (prior year, or the budget rate if the
  budget was set at a fixed rate).
- Report "reported growth = constant-currency growth + FX impact". A budget variance driven mostly by FX
  is not an operating miss, so say so.

**4.5 Capital controls / trapped cash**
- Detect: notes on repatriation limits, or entities in countries with exchange controls.
- Default: show that cash as restricted for group-liquidity purposes even when the GL treats it as
  unrestricted. MEDIUM.

---

## 5. Liquidity, solvency & going concern

**5.1 Covenant tests**
- Use `08_Debt_Covenants` if present; otherwise look for covenant notes.
- Default: compute each covenant *as the facility defines it*. "EBITDA" in a covenant usually allows
  add-backs, so use the definition given. If no definition is given, use reported EBITDA and say so.
- Show headroom as a %. Headroom < 15% is a red flag; a breach is a headline item.
- A covenant defined but not computable from the data is LOW.

**5.2 Runway that ignores obligations**
- Runway = (unrestricted cash + undrawn committed facility) ÷ normalized monthly net burn.
- Normalize the burn: exclude one-off inflows (debt draws, equity raises, asset sales) and one-off
  outflows (restructuring cash).
- Also show the debt maturities falling inside the runway window.

**5.3 A financing inflow masking operating burn**
- Detect: cash up while operating cash flow is negative; a debt or equity inflow in the period.
- Default: lead the cash commentary with *operating* cash flow, not the change in cash. This was the
  Meridian US01 revolver case.

**5.4 Going-concern indicators** (IAS 1.25 / ASC 205-40 style)
- Any of: negative equity; current liabilities > current assets combined with negative operating cash;
  normalized runway < 12 months; a covenant breach or < 15% headroom; debt maturing within 12 months
  with no refinancing evidence; recurring losses with declining cash.
- Default: never conclude on going concern; that's management's and the auditor's call. List the
  indicators present, factually, in a dedicated memo section and a deck callout. If two or more are
  present, it's a headline item on slide 1.

**5.5 Working-capital management masking burn**
- Detect: DPO rising > 15% vs the trailing 3 months while DSO is flat or rising; payables growing
  faster than cost; factoring or supplier-finance notes.
- Default: report the cash benefit as non-recurring and recompute burn without it.

---

## 6. Earnings quality & normalization

**6.1 One-time items**
- Use the `One-Time Item` flag. Without one, apply this policy: restructuring, impairment, write-downs,
  M&A deal costs, litigation settlements, gains/losses on disposal, and FX on a one-off transaction.
- Show an **Adjusted EBITDA bridge** that reconciles reported operating income → EBITDA → Adjusted
  EBITDA, one line per add-back, every add-back sourced. Present the reported figure first, always.

**6.2 "Non-recurring" items that recur**
- Detect: the same add-back category in 3+ of the last 6 periods.
- Default: keep the label the business uses, but flag in Red Flags that it is recurring and show Adjusted
  EBITDA both with and without it.

**6.3 Revenue cutoff / channel stuffing signals**
- Detect: revenue concentrated in the last month of a quarter well above the pattern; AR growth >
  revenue growth by > 15 pts; DSO up > 15% at period end; large credit notes in the following period;
  bill-and-hold notes.
- Default: never adjust revenue. Flag the pattern with the numbers. MEDIUM.

**6.4 Capitalization shifts**
- Detect: capitalized software / development / interest rising while the matching opex falls; capex
  growing faster than D&A for 3+ periods.
- Default: show opex both as reported and with the capitalized amount added back, so margin improvement
  from accounting choices isn't read as operating improvement.

**6.5 Reserve and provision releases**
- Detect: provision, allowance or accrual balances that decrease while revenue or AR grows; negative
  expense lines.
- Default: identify the P&L benefit. If it's material to the margin, show margin with and without it.

**6.6 Related-party transactions**
- Detect: notes, counterparty names matching entities or owners, transfer-pricing recharges.
- Default: disclose amounts. Eliminate within the group (section 3). For outside parties, list them in
  the memo appendix.

---

## 7. Revenue complexity

- **Deferred revenue rollforward:** opening + billings − revenue recognized = closing. Falling deferred
  revenue alongside stable recognized revenue is a leading indicator of churn, so flag it.
- **Bookings vs revenue:** never tie these month to month (see section 1). Use bookings for the bridge's
  drivers, and scale the bridge to GL revenue with an explicit "timing / recognition" plug line.
- **Gross vs net (principal vs agent):** a marketplace or reseller line changing presentation changes
  revenue growth, not economics. Compare gross profit, and flag a presentation change.
- **Percentage of completion / contract assets:** watch the contract assets (unbilled) to revenue ratio.
  Growing unbilled balances is a cash-conversion and estimate risk.
- **Backlog / RPO:** report it if provided. Never derive it from bookings without contract end dates.
- **Concentration:** top 1 / top 5 customers from bookings or AR aging. Single-customer exposure is a
  red flag per `red-flag-thresholds.md` even when revenue is growing.

---

## 8. Tax & cost-side complexity

- **Effective tax rate (ETR) swing:** a change of more than 10 pts vs the prior period or the statutory
  rate needs a stated driver: losses without tax benefit (valuation allowance), country mix, one-offs.
  Never back-solve a rate.
- **Loss-making entities with tax expense, profitable ones with none:** note it. It usually means a
  valuation allowance or a tax-loss carryforward.
- **Transfer pricing:** recharges between entities set the entity-level margins. Comment on
  consolidated margin; present entity margins only with a "post-recharge" label.
- **Inventory:** inventory growing faster than COGS for 2+ periods is an obsolescence risk; a write-down
  is a one-time item (6.1). LIFO entities can't be compared to FIFO/IFRS entities without a LIFO reserve.
- **Impairment triggers:** a sustained revenue decline, a segment exit, or market cap below book value
  (if given). Flag the trigger; never estimate an impairment.
- **Leases:** ASC 842 operating leases put rent in opex; under IFRS 16 it moves to depreciation and
  interest. So EBITDA is not comparable across standards. Show the lease-adjusted view when entities are
  on mixed standards.

---

## 9. Plan & forecast complexity

- **Multiple versions:** show actual vs the original budget *and* actual vs the latest forecast. A
  mid-year reforecast that "fixes" a miss is itself a finding.
- **Budget at a fixed FX rate:** compute the variance at the budget rate (constant currency, 4.4) and show
  FX separately.
- **Budget phasing that doesn't match actuals seasonality:** if a full year is on plan but months swing
  ±20% against plan in mirror-image patterns, comment on phasing, not performance.
- **Budget missing for an entity or account:** compare like-for-like on budgeted scope only, and show the
  unbudgeted scope as its own line (the Meridian IN01 case).
- **Budget built on a different taxonomy:** map it to the agreed taxonomy and log the mapping. Never
  compare at a level of detail the budget doesn't have.

---

## 10. When many things go wrong at once (Phase 5 prioritization)

A complex cycle can easily trip 30+ flags. A board deck that lists all of them communicates nothing.

1. **Score each flag** on: magnitude relative to base (% of revenue, EBITDA or cash), direction (adverse
   vs favourable), persistence (new / repeat / worsening), and whether it is **structural** (strategy,
   liquidity, solvency, data integrity) or **operational** (timing, phasing, one-off).
2. **Collapse related flags into one storyline.** Example: a hardware exit, an inventory write-down, an
   RIF, a restructuring charge and a gross-margin jump are *one* story ("exiting hardware"), with
   several numbers.
3. **Look at the interactions,** because they are where complex cases hide the real finding:
   - financing inflow × operating burn (5.3)
   - FX translation × budget miss (4.4)
   - capitalization × margin improvement (6.4)
   - an acquisition × organic growth (3.4)
   - a provision release × margin (6.5)
   - churn in bookings × stable GL revenue (7)
4. **Board deck and memo executive summary:** at most 5 storylines, ranked. Everything else goes to the
   workbook's Red Flags tab and a memo appendix list, with nothing dropped.
5. **Liquidity, solvency, covenant and data-integrity issues always outrank performance issues,**
   whatever their size.
