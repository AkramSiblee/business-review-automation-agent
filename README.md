# Business Review Automation Agent (Claude Code project)

## What this actually is

This is **not** a standalone program that runs by itself. It's a project folder for **Claude Code**
(Anthropic's coding agent) -- Claude Code is the agent; the files here are the instructions and test
data that make it good at this one specific job (turning a periodic company file into a business
review workbook, memo, and deck). Opening this folder in VS Code with the Claude Code extension and
asking it to "run this month's business review" is what actually does the work.

## Input format

Each cycle, you provide **one file**. The best-supported and expected case is a **single Excel
workbook with multiple tabs** -- one tab per data category (company config, GL/trial balance, cash,
headcount, revenue/bookings, budget, FX rates, intercompany detail). A `.docx`, `.pdf`, or `.pptx`
is also accepted -- Claude will extract whatever tables and figures it can find and map them into the
same categories, best-effort -- but a multi-tab Excel workbook will always give the most complete,
highest-confidence result since every category can be represented cleanly. See
`.claude/skills/business-review-automation/references/standard-input-file-templates.md` for the exact
tab/column contract, and `sample-data/` (below) for two ready-made examples of that layout.

## Prerequisites

1. **VS Code**, with the Claude Code extension (or the Claude Code CLI, which also integrates with
   VS Code's terminal).
2. **Claude Code installed** -- follow Anthropic's current install instructions rather than anything
   below, since these details change: https://docs.claude.com/en/docs/claude-code/overview
3. **Google Drive for Desktop** installed and signed in, *if* you want the phone/browser-drop-a-file
   workflow described below. Not required if you're happy working with a local folder instead.

## Setup

1. Unzip this project anywhere on your machine and open the folder in VS Code.
2. Open the Claude Code panel (or terminal) inside VS Code.
3. **Verify the skill loaded.** Ask Claude Code: *"What does the business-review-automation skill
   cover?"* If it can answer in detail (autonomy doctrine, 8 phases, confidence tiers), it's loaded.
   If not, tell it explicitly: *"Read .claude/skills/business-review-automation/SKILL.md and
   everything in its references/ folder before we start."* Project-level skill auto-discovery is a
   newer Claude Code feature and the exact convention may have moved since this was built -- if it
   doesn't pick the skill up automatically, the explicit-read approach always works regardless.
4. **Point it at your data folder.** All client data lives in ONE folder, with one subfolder per
   client, and each client folder has an `Input` and an `Output` subfolder:
   ```
   Business-Review-Automation/
   ├── CLNT-1/
   │   ├── Input/     <- this cycle's one input file
   │   └── Output/    <- every deliverable lands here
   └── CLNT-2/
       ├── Input/
       └── Output/
   ```
   Create it in Google Drive (web or desktop) and find where Google Drive for Desktop syncs it locally
   (typically `~/Google Drive/My Drive/...` on Mac, `G:\My Drive\...` or a path under your user folder
   on Windows). In `config/data-paths.json`, set `google_drive_root` to that local path and
   `active_client` to the client folder you're working on. Each cycle, that client's `Input` folder
   should hold exactly one input file -- nothing else. To switch clients, change `active_client`, or
   simply tell Claude which client to run.
5. **Run the setup check** (optional but recommended -- catches path typos before your first real
   cycle):
   ```
   python scripts/verify_setup.py
   ```
   (Use `python3` on Mac/Linux.) Claude Code also runs this at the start of every cycle. It catches an
   unmounted Drive, a file that's still syncing, or an input left open in Excel.

## First run: test the pipeline before trusting it with real data

`sample-data/sample-test-data/Business_Review_Sample_Input_FY2025.xlsx` is a single workbook (8 tabs,
one per standard category) with a full synthetic FY2025 dataset (Meridian Technologies Group, 3
entities, ~2,000 rows) purpose-built to exercise every part of this skill -- customer churn, a
restructuring charge, an RIF, multi-currency/multi-standard consolidation, transfer pricing, a
deliberately missing tab, ambiguous dates, and a disclosed restatement.

Copy that one file into your Drive `inputs` folder (or just point `data-paths.json` at
`sample-data/sample-test-data` directly for a first test), then in VS Code ask Claude Code:

> Run a business review cycle using the sample test data.

Check the outputs folder for the workbook, memo, and deck, and read the Assumptions & Judgment Log
to see how it handled each deliberately tricky element.

**Then stress-test it on a complex group.** `sample-data/complex-test-data/Business_Review_Complex_Input_FY2025.xlsx`
is a second synthetic workbook: Halcyon Industrial Holdings, 8 entities in 7 currencies. It plants every
scenario in the complex-scenarios playbook:
- **Group structure:** a 70%-owned subsidiary, a 40% equity-method JV with an April year, a mid-year acquisition
  and a disposal.
- **Currency:** a hyperinflationary entity.
- **Financing and earnings quality:** a covenant near breach and a loan maturing within 12 months, an
  undisclosed provision release and recurring "one-time" costs.
- **Data shape:** year-to-date, signed and thousands-scaled ledgers in three languages.

To check the result, open `sample-data/answer-keys/Complex_Test_Answer_Key_FY2025.xlsx`. It explains in plain
English what the review produces and how, lists every planted issue with what a correct review should do, gives
the expected figures, and has a scorecard you fill in. Copy only the input workbook into the inputs folder, never
the answer key. The technical files that build the dataset live in `developer/complex-test-data/`.

## Ongoing use

1. Each cycle, export your real data (GL/trial balance, cash, headcount, revenue/bookings, budget, FX
   rates if multi-currency, intercompany detail if multi-entity) into **one file**. A single Excel
   workbook with one tab per category, using the naming/column conventions in
   `.claude/skills/business-review-automation/references/standard-input-file-templates.md`, gives the
   most complete result. If a multi-tab workbook isn't practical this cycle, a `.docx`, `.pdf`, or
   `.pptx` with the same information works too -- Claude will extract what it can and flag anything it
   can't confidently classify.
2. Drop that one file into the Drive `inputs` folder from wherever's convenient -- phone, browser,
   desktop. It'll sync locally within moments.
3. In VS Code, ask Claude Code to run the cycle.
4. Deliverables land in the Drive `outputs` folder and sync back up automatically -- no extra step.
5. Open `<Period>_Reconciliation_and_Signoff.xlsx` from the outputs folder, fill in the yellow cells, save it
   in place, and tell Claude "check the sign-off for <Period>" (see below).

## How do I know the results are right? (real data has no answer key)

The test datasets come with answer keys because their correct answers were planted. Your real data doesn't,
so every cycle checks itself against figures you already trust:

- **Reconciliation & Sign-off file (every cycle).** It lists the review's key figures per entity (revenue, net
  income, cash, total assets, headcount) and for the group. Type the same figures from your own accounting
  system, bank statements and payroll into the yellow column: each one shows PASS or CHECK. Then agree with or
  correct each judgment call the review made, give owners to open items, and sign. Claude reads your completed
  file back, investigates every CHECK, applies your corrections, and rebuilds the deliverables (your entries are
  kept, and you re-approve the changed numbers).
- **Built-in checks.** Every cycle verifies that the balance sheet balances, cash ties to the bank data,
  payroll ties to headcount, intercompany trading cancels out, the workbook has no formula errors, and the memo
  and deck quote the same numbers. Results are on the sign-off file's "Built-in Checks" tab. They prove the
  review is consistent with itself; your control totals prove it matches reality.

**Before relying on it for a new client:**
1. **Back-test a closed period.** Run a cycle on a past quarter or year that already has an approved board pack
   or audited accounts, and compare. That past pack is your answer key for real data.
2. **Run in parallel for 2-3 cycles.** Keep the existing manual process going alongside, and compare the
   outputs before switching over.
3. **Then sign off every cycle.** The next cycle starts by reading the previous sign-off, so anything left
   unresolved carries forward instead of being forgotten.

## If you'd rather not install Google Drive for Desktop

You can skip Drive entirely and use the project's own `data/` folder, which has the same layout
(`data/<CLIENT>/Input` and `data/<CLIENT>/Output`; see `data/README.md`). Drag your one file into the
client's `Input` folder and pull results out of its `Output` folder manually. `config/data-paths.json`'s
fallback covers this automatically (leave `google_drive_root` unset).

A fully API-based alternative (Claude Code pulling/pushing to Drive directly via the Drive API,
without any local sync client) is also possible but needs its own Google Cloud OAuth credentials
that only you can create -- ask if you want that path built out instead; it's a heavier setup and
not included here by default.

## Project layout

```
business-review-automation-agent/
├── CLAUDE.md                              <- Claude Code reads this automatically every session
├── README.md                              <- this file
├── config/
│   └── data-paths.json                    <- edit this: your real Drive-synced folder paths
├── .claude/skills/business-review-automation/
│   ├── SKILL.md                           <- the full method (8 phases, autonomy doctrine)
│   └── references/                        <- taxonomy, thresholds, templates, triage guide, etc.
│       ├── complex-scenarios-playbook.md  <- source-of-truth hierarchy, data-shape traps, NCI/equity method,
│       │                                     M&A, hyperinflation, covenants, going concern, earnings quality
│       └── industry-kpi-packs.md          <- SaaS, manufacturing, retail, services, projects, marketplace, lending
├── data/
│   ├── inputs/                            <- drop your ONE file here each cycle (fallback if Drive isn't configured)
│   └── outputs/                           <- deliverables land here (fallback if Drive isn't configured)
├── sample-data/
│   ├── sample-test-data/                  <- Meridian FY2025 test file, for first-run validation
│   ├── complex-test-data/                 <- Halcyon FY2025 stress-test input (the only file to copy)
│   └── answer-keys/                       <- plain-English answer key + scorecard (Excel); never copy to inputs
├── developer/
│   └── complex-test-data/                 <- technical files: dataset generator, EXPECTED.json, answer-key builder
└── scripts/
    ├── brlib.py                           <- shared: path resolution, tab loading/classification, safe writes
    ├── verify_setup.py                    <- pre-flight check, start of every cycle
    ├── profile_inputs.py                  <- Phase 1/2.5: triage + Data Validation Report
    ├── verify_workbook.py                 <- Phase 8: Excel recalculation + formula-error scan
    ├── render_deck.py                     <- Phase 8: slide PNGs for visual QA
    ├── final_qa.py                        <- Phase 8: placeholders, currency/period/basis, cross-doc figures
    ├── build_signoff.py                   <- Phase 9: Reconciliation & Sign-off file; --check reads it back
    └── reference/meridian_fy2025/         <- worked example from the sample run (structure only, not reusable as-is)
```

Per-cycle build scripts are written to `<outputs>/_work/<Period>/`, next to the data, so client
figures never end up in this repo.
