---
name: business-review-automation
description: Use whenever the user wants a recurring monthly, quarterly, or annual business review (MBR/QBR/ABR) built end-to-end from a single uploaded file rather than live connectors — e.g. "run this month's business review," "build the QBR deck from this workbook," "here's this quarter's financials and headcount file, put together the board package," or any request to turn one periodic company file (typically a multi-tab Excel workbook, but a docx/pdf/pptx also works) into a consolidated financial/operational review. Built for autonomous, hands-off operation across users in different countries, currencies, accounting standards, and file conventions — use it whether this is the first setup cycle or the fiftieth routine one. Not for a one-off historical case study or external-company valuation — use three-statement-case-to-deck for that.
---

# Business Review Automation (MBR / QBR / ABR)

Turns a single uploaded file per cycle — the expected and best-supported case is one Excel workbook with multiple tabs (one tab per data category: GL/P&L, balance sheet, cash, headcount, revenue/bookings, budget, FX, intercompany); a docx, pdf, or pptx is also accepted and extracted best-effort into the same categories — into three linked deliverables every cycle: a detailed workbook, a CFO/board-ready written memo, and a presentation-grade deck. Whatever format or currency, language, or accounting standard it arrives in, and however many of the categories it actually covers this cycle. Designed to run end-to-end with minimal human input, for users anywhere in the world. See `references/standard-input-file-templates.md` for the exact tab/column contract, including how a non-Excel file maps onto it.

**The standard throughout:** the three deliverables are generated *from one consolidated workbook*, never typed separately. A CFO-grade review explains *why* a number moved, not just that it moved. And because this runs for people in different countries with different currencies, calendars, and accounting standards, every output states its currency, fiscal period, and accounting basis explicitly — never assume the reader shares the preparer's conventions.

## Autonomy doctrine — read this before every cycle

The default mode is **run the full cycle to a complete draft in one pass**, not stop-and-check at every step. Stalling repeatedly on low-stakes questions defeats the purpose of automation; silently guessing on high-stakes ones defeats the purpose of rigor. Resolve this with confidence tiers, applied at every judgment point in every phase below:

- **HIGH confidence** (the evidence is unambiguous — e.g., a column literally headed "Revenue," a date format confirmed by a day-value >12): proceed silently, log the decision in the workbook's Assumptions & Judgment Log (Phase 6), move on.
- **MEDIUM confidence** (a reasonable, defensible call, but a genuine alternative reading exists): proceed with the most defensible interpretation, log it prominently, and surface it in the final summary as something to glance at — but do not stop the pipeline for it.
- **LOW confidence** (the file, figure, or period genuinely could not be resolved with reasonable defaults — e.g., an unclassifiable file with no matching category, a number that would flip a red flag depending on which reading is right): this is the one tier that stops and asks. Batch every LOW-confidence item from the whole cycle into a single round of questions at the point they're discovered, rather than interrupting repeatedly.
- **Materiality override:** even a HIGH-confidence item gets treated as if MEDIUM if it's large enough to flip a red-flag threshold (Phase 5) or change a headline KPI's direction — a confidently-read wrong number is worse than an uncertain one, precisely because nothing will flag it later.
- **Never silently fabricate.** A gap gets a documented, reasonable default and a flag, or an explicit "not available this cycle" — never a plausible-looking invented figure.
- **One deliberate exception to "run in one pass":** Phase 2.5 below is a hard checkpoint. The pipeline always stops there and waits for the user before consolidating anything into the history workbook — this is not a LOW-confidence judgment call, it applies every cycle regardless of confidence tier.

## Phase 0 — Standing setup (once per business, revisited only when something structural changes)

0. **Fast-path check, every cycle, before reading anything else.** Look for an existing history workbook in the resolved outputs folder. If one exists and this cycle carries no flagged structural change, read its **Config & Setup** tab directly (`references/deliverable-templates.md`) for the reporting currency, fiscal year convention, accounting standard, thresholds, and consolidation methodology, and skip straight to Phase 1 — do not re-read the reference files in steps 1-4 below. Fall through to the full setup below only on a genuine first cycle for this business, or when a structural change is being declared this cycle.
1. **Locale & standards configuration.** Read `references/locale-and-standards-config.md`. Confirm once, per business (or per entity if consolidating multiple): reporting currency, fiscal year convention, accounting standard (US GAAP / IFRS / local GAAP), and the number/date format conventions to use in output. This is the configuration most likely to silently corrupt a global user's numbers if skipped — do not infer it from the first file alone without checking it explicitly at least once.
2. **Taxonomy.** Read `references/chart-of-accounts-kpi-taxonomy.md` — a default P&L / balance sheet / cash flow / revenue-driver / headcount taxonomy and KPI list, with GAAP/IFRS variant notes. Confirm or edit once; treat the agreed version as fixed across cycles so comparability isn't broken by silent redesign.
3. **Thresholds.** Read `references/red-flag-thresholds.md` — defaults are materiality-scaled (percentage- and standard-deviation-based) rather than fixed dollar amounts, so the same configuration works for a business of any size or currency. Confirm or adjust once.
4. **Multi-entity/currency setup, if applicable.** If more than one entity or currency is involved, read `references/multi-currency-multi-entity-consolidation.md` and confirm the consolidation methodology once.
5. **History workbook.** Create the master workbook with the taxonomy and configuration built in (including the Config & Setup tab) but no period data yet.
6. **Log structural changes going forward.** Any later change to the taxonomy, thresholds, or locale configuration gets dated and logged with a reason, in both the workbook's Config & Setup tab and the Assumptions & Judgment Log — silent, undated changes are what make multi-year trend data untrustworthy, and an undated change is also what would force every future cycle back onto the full setup path unnecessarily.

## Phases 1-2.5 run inside a fresh `general-purpose` agent, then pause for the user

Delegate Phases 1, 2, and 2.5 below to a fresh agent via the Agent tool with `subagent_type: "general-purpose"`,
rather than running the raw file parsing in the main thread — this keeps the main orchestrator thread's
context lean across many cycles, since raw spreadsheet contents never need to pass through it. This agent
stops after Phase 2.5 and does **not** touch Phase 3 (consolidation) yet — see the checkpoint below.

**History note (do not reintroduce):** this used to delegate to a dedicated custom subagent,
`business-review-data-processor` (`.claude/agents/business-review-data-processor.md`). That file has
been removed. On this project's Windows/Claude Code setup, a custom subagent whose frontmatter declares
`tools: Bash` silently receives **no shell/code-execution tool at all** — the allowlist resolver doesn't
map the literal string `Bash` to the platform's actual shell tool (`PowerShell` on Windows), and there is
no error, just a subagent that reports itself blocked on the first binary file it tries to open.
Confirmed by direct test: a fresh `general-purpose` agent (which is granted *all* tools rather than a
named allowlist) has no such problem and gets working shell access immediately. Do not recreate a custom
`.claude/agents/*.md` subagent for this skill's Phase 1-3 work — use `general-purpose` with the
self-contained prompt below instead.

**What to put in the spawn prompt**, every cycle (this agent starts cold with no memory of this
conversation, so the prompt must be fully self-contained):

- Instruct it to read `.claude/skills/business-review-automation/SKILL.md` in full first, for the
  autonomy doctrine (HIGH/MEDIUM/LOW confidence tiers) and the exact rules for Phases 1, 2, 2.5, and 3 —
  treat that as the authoritative process spec.
- Tell it which reference files to read and when: `references/standard-input-file-templates.md` and
  `references/file-triage-guide.md` always (Phase 1 schema matching and fallback classification);
  `references/chart-of-accounts-kpi-taxonomy.md` and `references/red-flag-thresholds.md` always (it
  needs the concrete structure/values, not just a summary); `references/multilingual-term-tables.md`
  only if a file's language isn't otherwise resolved; `references/multi-currency-multi-entity-consolidation.md`
  only if multi-entity/currency.
- Give it: the resolved inputs folder path, the resolved outputs folder path, the Phase 0 configuration
  (freshly confirmed this cycle, or read from the workbook's Config & Setup tab via the Phase 0 fast
  path), and whether a structural change was declared this cycle.
- Tell it to run Phase 1 (file triage & locale-aware parsing) over the one file in the inputs folder —
  its tabs if it's an Excel workbook, or its extracted tables/sections if it's a docx/pdf/pptx (if more
  than one file is somehow present this cycle, triage all of them the same way rather than erroring) —
  then Phase 2 (validation & anomaly detection) against the parsed data and the existing history workbook
  if one exists, and Phase 2.5 (produce the standalone Data Validation Report — see below) — **then stop.**
  Tell it explicitly not to start Phase 3 (consolidation) or touch the history workbook's period data yet;
  it should hold the parsed data in its own context so it can resume Phase 3 later without re-parsing.
- Tell it explicitly **not** to return raw spreadsheet rows, full file dumps, or a step-by-step restating
  of its parsing work. It should return only: a File Log summary (one line per input file — classified
  as, confidence tier, fate: used/flagged/unclassified); the Judgment Log entries it drafted so far this
  cycle; any genuine LOW-confidence items worded ready for the batched question round; and the Data
  Validation Report's file path plus a short summary of what it found.

**Checkpoint — do not skip.** Present the Data Validation Report's summary (and any batched LOW-confidence
questions) to the user in the main thread and wait for their go-ahead or corrections — a missing file
supplied, a triage call overridden, a flagged gap acknowledged as expected. Do not proceed to Phase 3 on
your own judgment, even at HIGH confidence across the board; this checkpoint is unconditional, per the
autonomy doctrine's one deliberate exception above.

Once the user responds, resume the *same* Phase 1-2.5 agent via SendMessage (not a fresh agent — it
already holds the parsed data) and tell it to proceed with Phase 3 (consolidation — append the validated
period, applying the confirmed Phase 0 configuration exactly as given, never re-deriving it, and folding
in any corrections the user just gave), then write/update the workbook's **Config & Setup**, **File Log &
Lineage**, **Assumptions & Judgment Log**, and **Taxonomy Map** tabs plus the financial-statement tabs,
per `references/deliverable-templates.md`. Create the workbook fresh if this is the first cycle; otherwise
update the existing one in place. If a structural change was declared, log it with a reason and today's
date in both the Config & Setup tab and the Judgment Log. Have it return the same kind of summary as
before (File Log, Judgment Log entries, workbook file path, confirmation Phase 3 completed) — still no raw
spreadsheet rows.

Continue Phases 4-8 in the main thread using that report and the updated workbook — never pull the raw
spreadsheet contents back into the main thread. If delegation still somehow reports itself blocked on a
tool-access issue, fall back to running Phases 1-3 directly in the main orchestrator thread rather than
retrying the same delegation (the checkpoint still applies: pause after Phase 2.5 and wait for the user
before consolidating).

## Phase 1 — File triage & locale-aware parsing (every cycle, run independently to completion)

0. **Confirm what arrived.** The expected input is one file: a multi-tab Excel workbook is the primary case, with a docx, pdf, or pptx also accepted (see "Non-Excel single files" in `references/standard-input-file-templates.md`). If it's a workbook, each tab is triaged independently as its own candidate category below; if it's a docx/pdf/pptx, extract its tables/figures first, then triage each extracted block the same way.
1. Check whether the workbook's tabs (or the non-Excel file's extracted sections) match the concrete schemas in `references/standard-input-file-templates.md` — a match gives HIGH confidence by construction. If they don't match exactly, fall back to the content-based heuristics in `references/file-triage-guide.md`. Assign a confidence tier per the doctrine above either way.
2. **Parse numbers and dates according to the file's own internal evidence, not assumption.** Detect decimal/thousands-separator convention and date-component order from patterns within the file itself (see `references/locale-and-standards-config.md` for the detection heuristics); fall back to the Phase 0 locale configuration only when the file itself is ambiguous, and log that fallback as a MEDIUM-confidence assumption.
3. Detect the file's currency and unit scale (actual / thousands / millions / regional conventions like lakhs or crores) from headers or note fields where present; if genuinely absent, assume the Phase 0 reporting currency and flag as an assumption.
4. Map classified line items into the fixed taxonomy. New or unrecognized line items get a MEDIUM-confidence best-fit placement and a log entry, not a silent drop and not a pipeline stop, unless the item is large enough to trigger the materiality override.
5. If a tab or extracted section doesn't match any category, or an expected category has no data this cycle, note it plainly in the File Log and proceed with the rest of the cycle — a gap in one category should not block the rest of the review. This is routine (not an error) when the input is a non-Excel single document that was never going to cover all eight categories.
6. If file content is in a language the classification heuristics don't cover with reasonable confidence, consult `references/multilingual-term-tables.md` where it reaches HIGH/MEDIUM confidence; anything left is a LOW-confidence item for the batched question round, and once resolved for one cycle, reuse that entity's mapping key in every future cycle rather than re-asking.

## Phase 2 — Validation & anomaly detection

- **Structural checks:** P&L flows into the retained-earnings rollforward; balance sheet balances; cash flow reconciles to the reported cash position; headcount file total ties to the payroll/opex line it should drive. An interim or pre-close trial balance will often fail the balance-sheet-balances check until current-period earnings are rolled into retained earnings — treat that specific imbalance as an expected unclosed-earnings adjustment to make, not a red flag, and confirm the balance check passes only after making it.
- **Continuity check:** any account, department, entity, or segment present last cycle and absent this cycle is flagged — the model never assumes whether that's a real discontinuation, a rename, or a data gap.
- **Restatement handling:** if a file revises a prior period's previously-reported figures, keep both the "as previously reported" and "as restated" values visible in the history workbook rather than overwriting history silently — a global multi-user tool has to assume restatements will happen and be auditable when they do.
- **Statistical anomaly detection, not just fixed thresholds:** once enough history exists (roughly 6+ periods), flag values that deviate materially from the trailing trend (mean and spread), in addition to the fixed materiality-scaled thresholds from Phase 0 — this catches issues a flat percentage threshold misses on unusually stable metrics, and avoids over-flagging on metrics that are normally volatile.
- **Duplicate and double-count checks:** watch for the same transaction or balance appearing in two source files (common when both a GL export and a sub-ledger detail file are provided).

## Phase 2.5 — Data Validation Report & checkpoint (every cycle — the pipeline always pauses here)

Before any data is consolidated into the history workbook, produce a standalone **Data Validation
Report** (`.xlsx`, one sheet, `openpyxl`) in the resolved outputs folder, named
`<Period>_Data_Validation_Report.xlsx` (e.g. `2026-Q1_Data_Validation_Report.xlsx`). This is a distinct
file, not a tab buried inside the history workbook — the point is that it can be opened and skimmed in
seconds, before the full workbook/memo/deck exist, so the user can catch a data problem while it's still
cheap to fix. It covers:

- **Expected vs received:** every standard input category (`references/standard-input-file-templates.md`)
  against what actually arrived this cycle — Present / Missing / Unrecognized, with the matched tab name
  (or extracted section, for a non-Excel file) and Phase 1 triage confidence tier for each present category.
- **Coverage per file:** entities, periods, and row counts found in each received file.
- **Cross-file entity/period gaps:** any entity, period, account, or department present in one file but
  absent in another where it would normally be expected — e.g., an entity with a full year of GL data but
  zero budget rows. Name it explicitly; don't fold it silently into an average or a lower confidence tier
  on some other line.
- **Structural check results from Phase 2:** balance sheet balances, cash flow reconciles, headcount ties,
  continuity check, restatement flags, duplicate/double-count flags — pass/fail per entity/period.
- **Every LOW-confidence item** discovered so far, batched per the autonomy doctrine, plus any
  MEDIUM-confidence call material enough to be worth a second look before it propagates further.

**This is the one unconditional checkpoint in the cycle** (see the autonomy doctrine above). Stop here
regardless of confidence tier — even an all-HIGH-confidence cycle still pauses, because the point of this
report is letting the user catch something the automation itself would have no reason to flag (e.g., a
file they meant to include but forgot to export). Present the report's summary to the user and wait for
their go-ahead or corrections before continuing to Phase 3.

## Phase 3 — Consolidation into the history workbook (multi-entity/currency aware)

- Append the validated period to the running master workbook.
- If multiple entities/currencies are involved, apply the confirmed methodology from `references/multi-currency-multi-entity-consolidation.md`: translate at the appropriate rate basis (spot for balance sheet, average for P&L, under the current-rate method unless configured otherwise), eliminate identifiable intercompany transactions, and disclose plainly if intercompany data wasn't available to fully eliminate.
- Flag entity changes mid-period (acquisition, divestiture, new subsidiary) explicitly, and show an organic/like-for-like view alongside the as-reported consolidated view where the data supports it.
- Recompute MoM/QoQ/YoY and actual-vs-budget views automatically — this should never require being re-specified each cycle.

## Phase 4 — Driver-based analysis

- Revenue: new business, expansion/upsell, churn/contraction, and price vs. volume where the data supports it.
- Opex: headcount-driven vs. non-headcount cost, by department, with hiring/attrition detail.
- Margin: a bridge from prior to current period isolating price, cost, mix, and one-time effects.
- Cash: operating cash generation vs. reported net income, burn rate, runway.
- Balance sheet: working-capital trend, leverage, covenant-relevant ratios.
- **Standard-awareness:** if entities on different accounting standards are being rolled up (e.g., one subsidiary under IFRS, another under US GAAP), flag known divergence points that commonly distort a combined view — lease capitalization, revenue-recognition timing, R&D/development cost capitalization, inventory valuation method — rather than presenting a blended number as if it were prepared on one consistent basis.

## Phase 5 — Insight generation & red-flag scan

- Rank the largest movers by dollar impact relative to their own base (e.g., % of revenue or % of the relevant line), not by raw dollar size alone — this is what makes the ranking meaningful across companies of very different scale.
- Run every metric against both the fixed materiality-scaled thresholds and the statistical trend check from Phase 2; anything tripping either gets a place on the list regardless of how it would otherwise rank.
- Draft the one-line "why" for each flagged item, tied to the Phase 4 driver analysis, and reuse this narrative in both the memo and the deck rather than redrafting it twice.
- Carry forward any prior-cycle flag that hasn't improved and say so explicitly — repetition without progress is itself the finding.

## Phase 6 — Generate the three deliverables

Before building any output file, view the relevant skill: `/mnt/skills/public/xlsx/SKILL.md`, `/mnt/skills/public/docx/SKILL.md`, `/mnt/skills/public/pptx/SKILL.md`. These are Anthropic's hosted-sandbox skill paths and won't exist in a local Claude Code checkout (e.g. this project run via the CLI on Windows) — if they're not found, generate the files directly with `openpyxl` (.xlsx), `python-docx` (.docx), and `python-pptx` (.pptx), installing them first if needed. Structure templates, including the workbook's required **Assumptions & Judgment Log** tab, are in `references/deliverable-templates.md`.

- Every financial figure in every deliverable states its currency (code, not just symbol, given a global audience — "1,234 USD" not just "$1,234") and the accounting basis it was prepared under.
- Dates use an unambiguous format in every output (e.g., "31 Mar 2026," never a bare numeric date that could be read two ways).
- Number formatting in outputs follows the Phase 0 configuration, not the format any particular source file happened to arrive in.
- The workbook's Assumptions & Judgment Log lists every MEDIUM- and LOW-confidence call made this cycle, so a human reviewer can audit the automation's judgment without redoing the work.
- Never re-derive in the memo or deck a number the workbook already computed — pull it forward.

## Phase 7 — Scale depth to cadence

- **Monthly:** flash review — KPI dashboard, top movers, red flags, lighter deck (5-7 slides). The workbook is still fully updated; this is where the compounding trend value gets built even though the monthly output itself is light.
- **Quarterly:** the full three-deliverable package at full depth, plus quarter-over-quarter and quarter-to-date-vs-plan.
- **Annual:** full package plus a multi-year trend section, a strategic narrative, and explicit flags for where this year's actuals should change next year's budget assumptions.
- **First cycle covering a full prior year in one batch** (e.g. onboarding a business mid-year with all months-to-date supplied at once): populate every period into the history workbook as usual, but generate only the deliverables for the *most recent completed period* at its natural cadence — don't generate a separate report per historical month. Note plainly in the memo/deck that trend lines cover a backfilled history, not periods reviewed as they occurred.

## Phase 8 — QA (every cycle, before anything is delivered)

- [ ] Workbook recalculates with zero formula errors; every articulation check reads TRUE
- [ ] Deck passes file validation, has no placeholder text, and passes a visual QA render pass
- [ ] Memo has been checked against the workbook for numerical consistency
- [ ] **Cross-document check:** headline figures in the deck and memo match the workbook exactly
- [ ] Currency, fiscal period, and accounting basis are stated explicitly on every deliverable's cover/summary
- [ ] Every MEDIUM/LOW-confidence judgment call this cycle appears in the Assumptions & Judgment Log
- [ ] FX rates and consolidation methodology used (if applicable) are documented and consistent with Phase 0 configuration
- [ ] Every file received this cycle has a stated fate in the File Log: used, flagged, or explicitly unclassified — nothing silently dropped
- [ ] The Data Validation Report (Phase 2.5) was generated, presented to the user, and the pipeline actually paused for their response before Phase 3 ran — not skipped even on an all-HIGH-confidence cycle
- [ ] Every LOW-confidence item from Phase 1-5 was either resolved via the batched question round or explicitly carried as an open flag into the deliverables — never silently dropped because it was inconvenient to resolve

## Output format

Four files per cycle: a `.xlsx` Data Validation Report (Phase 2.5, generated first, before the checkpoint),
then — once the user has responded to that checkpoint — the three main deliverables: `.xlsx` (full model),
`.docx` (executive memo), `.pptx` (board deck). All named consistently with period and currency (e.g.,
`2026-Q3_Business_Review_USD.xlsx`, `2026-Q3_Data_Validation_Report.xlsx`) so a global user's archive stays
navigable across currencies and periods.
