# Project: Business Review Automation Agent

This project turns a single recurring input file each cycle — typically one multi-tab Excel workbook
(GL/trial balance, cash, headcount, revenue/bookings, budget, FX, intercompany all as tabs inside that
one file), though a docx, pdf, or pptx is also accepted — into three linked deliverables: a
formula-driven Excel workbook, a written memo, and a board deck. The actual method is fully specified
in the skill below -- read it in full before doing any review work, don't try to reconstruct the
process from this file alone.

## The skill (read this first, every time)

The complete process lives at:

`.claude/skills/business-review-automation/SKILL.md`

with supporting reference files in `.claude/skills/business-review-automation/references/`.

If this project's Claude Code version auto-discovers project-level skills from `.claude/skills/`,
it will already be loaded. If you're not certain that happened, or the skill isn't showing up in
your available context, explicitly read `SKILL.md` and everything in `references/` before starting
any review cycle -- don't proceed on partial knowledge of the method.

## Where the data lives

Input and output files are **not** stored inside this repository -- they live in a
Google-Drive-for-Desktop-synced folder on the user's machine, so a file dropped into Drive from a
phone or browser shows up locally without any custom sync code.

0. Run `python scripts/verify_setup.py` first, every cycle. It resolves the paths below the same way
   every other script does (`scripts/brlib.py`), and it stops you early if Drive isn't mounted or an
   input is still syncing.
1. Read `config/data-paths.json` first, every cycle.
2. Resolve the real input folder as `<google_drive_root>/<inputs_subfolder>` and the real output
   folder as `<google_drive_root>/<outputs_subfolder>` from that file.
3. If `google_drive_root` is still unset/`"REPLACE_ME/..."` (i.e. the user hasn't configured it yet),
   fall back to the local `data/inputs/` and `data/outputs/` folders in this project, and tell the user
   plainly that Drive isn't configured yet so they know why files aren't syncing to their phone/browser.
4. **Read the one file present in the resolved inputs folder** -- every tab if it's an Excel workbook,
   or everything extractable (tables, figures, labeled line items) if it's a docx/pdf/pptx. The
   resolved inputs folder should normally contain exactly one file per cycle; if more than one is
   somehow present, triage all of them per Phase 1 rather than erroring, but a single file is the
   expected and best-supported shape -- see `references/standard-input-file-templates.md`. Write every
   deliverable (workbook, memo, deck, and the Assumptions & Judgment Log) to the resolved outputs
   folder. Never write review deliverables anywhere else.

## First run / testing this pipeline

`sample-data/sample-test-data/Business_Review_Sample_Input_FY2025.xlsx` is a full synthetic test
dataset (Meridian Technologies Group, FY2025, 3 entities, ~2,000 rows across all 8 standard tabs in
one workbook) built specifically to exercise this skill: customer churn and concentration risk, a
restructuring charge and RIF, an inventory write-down, a debt draw that masks a cash-vs-earnings
divergence, multi-currency/multi-standard consolidation (US GAAP / IFRS / Ind AS), transfer pricing
between entities, a deliberately missing budget tab, ambiguous DD/MM/YYYY dates, and a disclosed
restatement. Copy that single file into the resolved inputs folder (or point `config/data-paths.json`
at `sample-data/sample-test-data` directly) and run a full cycle before connecting real data, to
confirm the pipeline behaves end-to-end and to see how the skill's autonomy doctrine handles each
embedded issue.

## Running a cycle

Ask in plain language, e.g. "Run this month's business review" or "Run a cycle using the sample test
data." Claude Code should then follow the skill's phases (0 through 8) exactly as written, batching any
LOW-confidence questions into a single round rather than asking repeatedly, and writing everything to
the resolved outputs folder at the end.

**Every cycle produces a Data Validation Report first and pauses.** Before consolidating anything into
the history workbook, the skill's Phase 2.5 writes a standalone `<Period>_Data_Validation_Report.xlsx` to
the resolved outputs folder listing what input categories (tabs, or extracted sections for a non-Excel
file) were expected vs. received, what's missing, and any cross-category gaps (e.g. an entity with GL
data but no matching budget data). The pipeline then stops and waits for confirmation or corrections
before building the workbook, memo, and deck — this checkpoint is unconditional and applies even when
every other judgment call this cycle was HIGH confidence.

**Every cycle ends with a Reconciliation & Sign-off file.** Real client data has no answer key, so Phase 9
writes `<Period>_Reconciliation_and_Signoff.xlsx` (via `scripts/build_signoff.py`): control totals for the
client to compare with their own accounting system, bank and payroll, the cycle's built-in check results,
every MEDIUM/LOW judgment call for them to agree or correct, open items, and a sign-off block. It doesn't
block delivery. When the client returns it, run `build_signoff.py --check`, act on every CHECK figure and
correction, and rebuild. At the start of the next cycle, check the previous period's file and carry its
unresolved items forward.

## Conventions

- Never invent figures to fill a gap; follow the skill's autonomy doctrine (HIGH/MEDIUM/LOW confidence
  tiers) instead.
- Every output states its currency, fiscal period, and accounting basis explicitly.
- Structural changes to the taxonomy, thresholds, or locale configuration get logged with a reason,
  per Phase 0 of the skill -- don't silently redesign the taxonomy between cycles.
