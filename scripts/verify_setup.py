#!/usr/bin/env python3
"""
Pre-flight check -- run at the start of every cycle (and once after first setup):

  python scripts/verify_setup.py

Checks config/data-paths.json, that the Drive folders are reachable and writable, what's waiting in
the inputs folder, Python packages, Office automation for Phase 8 QA, and the skill files.
Exit code 0 = ready to run a cycle, 1 = something must be fixed first.
"""
import importlib
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import brlib  # noqa: E402

PACKAGES = [("openpyxl", "openpyxl", True), ("docx", "python-docx", True), ("pptx", "python-pptx", True),
            ("win32com.client", "pywin32", False)]


def main():
    problems = []

    def ok(msg):
        print(f"[OK]   {msg}")

    def fail(msg):
        print(f"[FAIL] {msg}")
        problems.append(msg)

    def info(msg):
        print(f"[INFO] {msg}")

    try:
        cfg = brlib.load_config()
        ok("config/data-paths.json parses")
    except FileNotFoundError:
        fail("config/data-paths.json not found")
        sys.exit(1)
    except json.JSONDecodeError as e:
        fail(f"config/data-paths.json is not valid JSON: {e}")
        sys.exit(1)

    root = (cfg.get("google_drive_root") or "").strip()
    if not root or root.startswith("REPLACE_ME"):
        info("google_drive_root not configured -- using local data/inputs and data/outputs. Files won't "
             "sync to your phone/browser until you set it.")
    elif not os.path.isdir(root):
        fail(f"google_drive_root '{root}' is not reachable. Start Google Drive for Desktop (and check the "
             f"drive letter/path), or the cycle will silently fall back to local folders.")

    inputs_dir, outputs_dir, using_drive = brlib.resolve_paths()
    for label, path in (("inputs", inputs_dir), ("outputs", outputs_dir)):
        if not os.path.isdir(path):
            fail(f"{label} folder does not exist: {path}")
            continue
        try:
            probe = os.path.join(path, ".write_test")
            with open(probe, "w") as f:
                f.write("ok")
            os.remove(probe)
            ok(f"{label} folder exists and is writable: {path}")
        except OSError as e:
            fail(f"{label} folder is not writable: {e}")

    files = brlib.list_input_files(inputs_dir)
    if os.path.isdir(inputs_dir):
        locks = [n for n in os.listdir(inputs_dir) if n.startswith("~$")]
        if locks:
            info(f"Input file(s) open in Office right now ({', '.join(locks)}) -- save and close before running.")
        partial = [n for n in os.listdir(inputs_dir) if n.lower().endswith((".tmp", ".crdownload", ".partial"))]
        if partial:
            fail(f"Partially synced/downloaded file(s) in inputs: {', '.join(partial)} -- wait for sync to finish.")
    if not files:
        info("Inputs folder is empty -- drop this cycle's file in before running (or use the sample data).")
    elif len(files) == 1:
        ok(f"One input file waiting: {os.path.basename(files[0])}")
    else:
        info(f"{len(files)} input files waiting (one multi-tab workbook is the standard case; they'll be "
             f"pooled per category): {', '.join(os.path.basename(f) for f in files)}")
    for f in files:
        if os.path.getsize(f) == 0:
            fail(f"{os.path.basename(f)} is 0 bytes -- probably still syncing from Drive (or 'online-only').")

    for module, pip_name, required in PACKAGES:
        try:
            importlib.import_module(module)
            ok(f"Python package {pip_name}")
        except ImportError:
            if required:
                fail(f"Python package {pip_name} missing -- pip install {pip_name}")
            else:
                info(f"{pip_name} missing -- Phase 8 Excel/PowerPoint QA will be manual (pip install {pip_name})")

    sample = os.path.join(brlib.ROOT, "sample-data", "sample-test-data")
    if os.path.isdir(sample) and any(n.endswith(".xlsx") for n in os.listdir(sample)):
        ok("Sample test workbook present")
    else:
        info("Sample test workbook not found (only needed for test runs)")

    skill = os.path.join(brlib.ROOT, ".claude", "skills", "business-review-automation")
    refs = os.path.join(skill, "references")
    if os.path.exists(os.path.join(skill, "SKILL.md")) and os.path.isdir(refs):
        ok(f"Skill present ({len(os.listdir(refs))} reference files)")
    else:
        fail("Skill files missing under .claude/skills/business-review-automation/")

    print()
    print("Ready to run a cycle." if not problems else f"Fix the {len(problems)} [FAIL] item(s) above first.")
    sys.exit(0 if not problems else 1)


if __name__ == "__main__":
    main()
