# System Requirements

What a brand-new machine needs before it can run a full business review cycle with this project
and get back a workbook, memo, and deck. This project is **built and tested on Windows**; notes
below flag what's Windows-specific vs. portable.

## Required

1. **Windows 10/11.**
   This project's Phase 8 automated QA (`scripts/verify_workbook.py`, `scripts/render_deck.py`)
   drives Microsoft Excel and PowerPoint directly via COM automation (`win32com`), which only
   exists on Windows. Mac/Linux can run the core generation scripts (openpyxl/python-docx/python-pptx
   produce the same .xlsx/.docx/.pptx files fine), but the automated "recalculate and scan for
   formula errors" / "visual QA render" step has no equivalent there and hasn't been tested on
   those platforms — you'd be doing that QA step manually instead.

2. **Claude Code** (Anthropic's coding agent) — this is what actually reads the skill and runs
   the cycle; nothing in this project executes on its own. Install per Anthropic's current
   instructions (these change, so don't rely on a static version here):
   https://docs.claude.com/en/docs/claude-code/overview
   - Requires an Anthropic account signed in through Claude Code, with either a Claude subscription
     (Pro/Max, which includes Claude Code usage) or API billing enabled.
   - Requires a live internet connection every cycle — Claude Code calls Anthropic's hosted Claude
     models; there is no local/offline model involved anywhere in this project.

3. **VS Code**, with the Claude Code extension — the workflow this project's README assumes
   (open the folder, use the Claude Code panel/terminal). The Claude Code CLI on its own also
   works in any terminal if you'd rather skip VS Code; the extension is a convenience, not a
   hard dependency of the skill itself.

4. **Python 3.9+**, with these packages installed (`pip install <name>`):
   - `openpyxl` — builds/reads the .xlsx workbook and Data Validation Report
   - `python-docx` — builds the .docx memo and methodology document
   - `python-pptx` — builds the .pptx board deck
   - `pywin32` — **Windows only**; required for the Phase 8 QA scripts (`verify_workbook.py`,
     `render_deck.py`) to drive Excel and PowerPoint via COM

5. **Microsoft Excel and Microsoft PowerPoint**, desktop versions, licensed and installed locally
   (not Office web/mobile — COM automation needs the desktop app). Used only for Phase 8 QA:
   forcing the workbook to recalculate and scanning for formula errors, and rendering the deck for
   a visual QA pass. Without these installed, that automated QA step can't run as designed.

## Optional

- **Google Drive for Desktop** — only needed for the phone/browser "drop a file, it syncs, Claude
  picks it up" workflow described in the README. Skip it and use the local `data/<CLIENT>/Input` /
  `data/<CLIENT>/Output` fallback folders instead; the pipeline works identically either way.
- **Git** — the project folder is now a git repo (initialized locally) so changes are diffable
  going forward. Not required by the skill itself, but recommended for tracking edits to the
  skill, config, or scripts over time.

## Not required

- No GPU, no local/downloaded LLM weights — all reasoning happens via Claude Code's hosted models.
- No database or backend service.
- No cloud storage account beyond the optional Google Drive integration above.
- Hardware needs are modest: any laptop from roughly the last 6 years with a few hundred MB of
  free disk space for Python packages and per-cycle output files is sufficient.
