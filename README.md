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
4. **Point it at your Google Drive folder.** In Google Drive (web or desktop), create a folder --
   e.g. `Business-Review-Automation` with `inputs` and `outputs` subfolders inside it. Find where
   Google Drive for Desktop syncs that to locally (typically `~/Google Drive/My Drive/...` on Mac,
   `G:\My Drive\...` or a path under your user folder on Windows). Open `config/data-paths.json` in
   this project and replace `google_drive_root` with that local path. Each cycle, the `inputs`
   subfolder should hold exactly your one input file -- nothing else.
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

## If you'd rather not install Google Drive for Desktop

You can skip Drive entirely and just use the local `data/inputs/` and `data/outputs/` folders in
this project directly -- drag your one file into `data/inputs/` and pull results out of
`data/outputs/` manually. `config/data-paths.json`'s fallback covers this automatically (leave
`google_drive_root` unset).

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
│   └── sample-test-data/                  <- Meridian FY2025 test file, for first-run validation
└── scripts/
    ├── brlib.py                           <- shared: path resolution, tab loading/classification, safe writes
    ├── verify_setup.py                    <- pre-flight check, start of every cycle
    ├── profile_inputs.py                  <- Phase 1/2.5: triage + Data Validation Report
    ├── verify_workbook.py                 <- Phase 8: Excel recalculation + formula-error scan
    ├── render_deck.py                     <- Phase 8: slide PNGs for visual QA
    ├── final_qa.py                        <- Phase 8: placeholders, currency/period/basis, cross-doc figures
    └── reference/meridian_fy2025/         <- worked example from the sample run (structure only, not reusable as-is)
```

Per-cycle build scripts are written to `<outputs>/_work/<Period>/`, next to the data, so client
figures never end up in this repo.
