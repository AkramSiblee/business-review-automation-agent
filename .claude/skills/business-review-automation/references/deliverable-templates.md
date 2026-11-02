# Deliverable templates

Follow these structures in Phase 6, scaled per Phase 7. All three deliverables are generated from the same consolidated workbook \u2014 never re-type a figure for the memo or deck that the workbook already computed.

## Global formatting rules (apply to every deliverable, every cadence)

- State the **reporting currency (ISO code)**, **fiscal period**, and **accounting basis** explicitly on the cover or opening summary of every deliverable.
- Use unambiguous date formats everywhere in outputs (e.g., "31 Mar 2026"), never a bare numeric date.
- Format numbers per the Phase 0 locale configuration, not per whatever convention a source file happened to use.
- If multiple currencies/entities are consolidated, state the FX methodology and rates used in the workbook and reference it (don't re-derive it) in the memo/deck.

## Workbook (.xlsx) \u2014 same tab set every cadence, always fully updated

0. **Config & Setup** \u2014 the confirmed Phase 0 decisions in one compact block: reporting currency, fiscal year convention, accounting standard, materiality thresholds, and (if multi-entity) consolidation methodology, each with the date first confirmed. Append a dated line each time a later cycle logs a structural change. This tab is what lets a future cycle skip re-reading the Phase 0 reference files \u2014 read this tab instead when it's already present and no structural change is flagged (see `SKILL.md` Phase 0).
1. **File Log & Lineage** \u2014 every file received this cycle, its triage classification and confidence tier, and any flags raised (unclassified files, missing categories, continuity gaps)
2. **Assumptions & Judgment Log** \u2014 every MEDIUM- and LOW-confidence call made this cycle (locale/date/currency assumptions, taxonomy placements, consolidation judgment calls), each with its confidence tier and resolution status. This is the audit trail that lets a human reviewer trust an autonomously-run cycle without redoing the work.
3. **Taxonomy Map** \u2014 the agreed chart of accounts / KPI taxonomy, with each source file's line items mapped to it, and the accounting standard noted per entity if more than one is in play
4. **Income Statement** \u2014 full history + current period + budget, with variance columns, in the reporting currency
5. **Balance Sheet** \u2014 full history + current period, with an articulation check row (Assets = Liabilities + Equity, must read TRUE)
6. **Cash Flow** \u2014 full history + current period, ending cash linked to the Balance Sheet cash line
7. **FX & Consolidation** (only if multi-entity/currency) \u2014 rates used, translation method, intercompany eliminations applied and any that couldn't be applied, entity changes mid-period
8. **Revenue Drivers** \u2014 the new/expansion/churn/price-volume bridge
9. **Headcount & Opex Detail** \u2014 by department, with hiring/attrition detail and cost-per-head
10. **KPI Dashboard** \u2014 every metric from the taxonomy's KPI list in one view, trended
11. **Red Flags & Variance Log** \u2014 every threshold trip (fixed and statistical) this cycle, with the driver-based explanation, and whether it's new or carried over. Add a storyline ID and a priority score (playbook \u00a710), so the memo and deck pull the top 5 from here.

**Conditional tabs.** Add each one when its trigger applies, and keep it in every later cycle once added,
for comparability:

12. **Adjustments & Non-GAAP Bridge.** Trigger: any one-time item or add-back. Reported operating
    income \u2192 EBITDA \u2192 Adjusted EBITDA, one row per add-back, with its source (GL account / note),
    category, and how many of the last 6 periods the category has appeared in.
13. **Liquidity, Debt & Covenants.** Trigger: any borrowing, or runway < 24 months.
    - Facility register: limit, drawn, undrawn, maturity, rate.
    - Covenant tests: definition, actual, threshold, headroom %.
    - Normalized burn and runway.
    - Maturity ladder.
    - Going-concern indicator checklist, with each item present / absent.
14. **Ownership & Scope.** Trigger: NCI, an equity-method entity, an acquisition, a disposal, or
    discontinued operations.
    - Entity, method, ownership % and date of change.
    - Parent vs NCI split.
    - Equity-method rollforward.
    - Goodwill (provisional or final).
    - Organic vs reported bridge.
    - Continuing vs discontinued P&L.
15. **Constant Currency.** Trigger: multi-currency. Current period at comparison rates, and the FX impact
    on revenue, EBITDA and net income, per entity and in total.
16. **Industry KPIs.** Trigger: a pack was selected in Phase 0. Show the pack's KPIs, trended, and mark
    any whose input data was missing as "not available".

## Executive memo (.docx) \u2014 depth scales with cadence

**Monthly (flash):** Executive Summary + KPI table + Top Movers + Red Flags. Two to three pages.

**Quarterly (full):**
1. Executive Summary (the narrative arc, one page, with currency/period/basis stated)
2. Financial Performance (P&L walk, margin bridge)
3. Balance Sheet & Cash Position
4. Revenue & Growth Drivers
5. Headcount & Organization
6. Risks & Watch Items (every tripped threshold, explained)
7. Outlook / Next-Period Focus
8. Appendix reference to the workbook (state the filename, not a re-derivation)

**Annual:** the Quarterly structure plus a Multi-Year Trend section and a Strategic Narrative / Next-Year Planning Input section.

**Conditional memo sections (any cadence, including the monthly flash).** Insert each one when its
trigger applies:

- **Liquidity & Going-Concern Indicators.** Trigger: any indicator present, or covenant headroom < 15%.
  Place it directly after the Executive Summary. List the indicators factually, with no conclusion.
- **Earnings Quality.** Trigger: the Adjustments tab exists. A reported → adjusted reconciliation, plus
  recurring "one-time" items and capitalization or provision effects.
- **Scope & Comparability.** Trigger: an acquisition, a disposal, NCI, an equity-method entity, or a
  restatement. State the organic vs reported and continuing vs discontinued figures, and what a reader
  should compare against what.
- **FX & Constant Currency.** Trigger: multi-currency with an FX impact > 2% of revenue or EBITDA.

The Executive Summary leads with at most 5 ranked storylines (playbook §10).

## Board deck (.pptx) \u2014 slide count scales with cadence

**Monthly (flash), ~5-7 slides:** Title/Executive Summary \u2192 KPI Dashboard \u2192 Top Movers \u2192 Cash & Runway \u2192 Red Flags.

**Quarterly (full), ~10-12 slides:**
1. Title / Executive Summary (currency, period, and accounting basis stated on this slide)
2. KPI Dashboard (scorecard view)
3. P&L Performance (chart + commentary)
4. Revenue Drivers (bridge/waterfall chart)
5. Opex & Headcount
6. Balance Sheet & Cash / Runway
7. Budget vs. Actual \u2014 variance highlights
8. Risks & Red Flags
9. Outlook / Next-Quarter Focus
10. Appendix (methodology, data lineage, FX/consolidation note if applicable)

**Annual:** the Quarterly structure plus a multi-year trend slide and a strategic-narrative slide, typically 13-15 slides total.

**Conditional slides (any cadence):**
- **Liquidity & Covenants:** headroom gauge, runway, maturity ladder. Placed as slide 2 when 2+
  going-concern indicators are present or any covenant is breached.
- **Organic vs Reported / Scope Changes:** when acquisitions, disposals or discontinued operations exist.
- **Reported → Adjusted EBITDA bridge:** a waterfall, when add-backs exceed 5% of reported EBITDA.

Keep the main body within the cadence's slide count: move detail to the appendix rather than adding
body slides.

Design standard for every deck regardless of cadence: native charts (not screenshots), one clear "so what" per chart, no filler slides, and the same visual QA pass every cycle.

## Reconciliation & Sign-off (.xlsx) - Phase 9, every cycle

Built by `scripts/build_signoff.py` from `<outputs>/_work/<Period>/signoff_inputs.json`. The script owns the
layout (tabs: Start Here, 1 Control Totals, 2 Built-in Checks, 3 Judgment Calls, 4 Open Items, 5 Sign-off;
yellow cells for the client, PASS/CHECK formulas, dropdowns). The cycle only supplies the content below.

**Spec keys:** `period`, `period_text`, `currency`, `basis`, `company`, `workbook`, then four lists:

| List | Each entry | Notes |
|---|---|---|
| `figures` | `section`, `entity`, `figure`, `unit`, `value`, `source` (model sheet!cell), `where_to_find`, `tolerance`, optional `client_value` | `value` is read from the finished model, never retyped. |
| `checks` | `check`, `result` (PASS / WARN / FAIL), `detail` | Structural tie-outs plus the Phase 8 script results. |
| `judgments` | `ref`, `tier`, `item`, `treatment`, `impact` | Every MEDIUM and LOW call. Omit the key to read them from the model's Assumptions & Judgment Log tab. |
| `open_items` | `item`, `why_open`, `needed` | LOW items carried open, data gaps, and unresolved items from the prior period's sign-off. |

**Minimum control figures.** Choose figures the client can read directly off a system they already trust,
not figures only the review can produce:

1. **Per entity, in its own currency** (from that entity's accounting system): revenue, net income, closing
   cash, total assets, and period-end headcount (from HR/payroll) where headcount was supplied. Use tolerance
   `"1"`: these should match the source exactly, so any difference points to an incomplete export or a parsing
   error (YTD, unit scale, sign convention).
2. **Group, in the reporting currency:** revenue, net income (attributable to the parent when NCI exists),
   closing cash, and total equity. Use tolerance `"0.5%"`, because the client's own consolidation may use
   slightly different rates. Set `where_to_find` to "Your consolidated management accounts, if prepared".
3. **Each tested covenant ratio** (tolerance `"0.01"`), against the compliance certificate sent to the
   lender. Where the review's recomputed figure differs from the company's reported figure by definition,
   say so in `where_to_find`.
4. **Anything the client handed over as a control total** in conversation, with `client_value` filled in.

Keep it to roughly 5 per entity plus 5 for the group. The aim is a 15-30 minute review, not a second audit.
