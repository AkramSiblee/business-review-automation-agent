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
11. **Red Flags & Variance Log** \u2014 every threshold trip (fixed and statistical) this cycle, with the driver-based explanation, and whether it's new or carried over

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

Design standard for every deck regardless of cadence: native charts (not screenshots), one clear "so what" per chart, no filler slides, and the same visual QA pass every cycle.
