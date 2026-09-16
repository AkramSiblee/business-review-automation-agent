#!/usr/bin/env python3
"""
Sanity-checks config/data-paths.json before you run a real business review cycle.
Run from the project root: python3 scripts/verify_setup.py
"""
import json
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(ROOT, "config", "data-paths.json")


def fail(msg):
    print(f"[FAIL] {msg}")
    return False


def ok(msg):
    print(f"[OK]   {msg}")
    return True


def main():
    print(f"Checking {CONFIG_PATH} ...\n")
    if not os.path.exists(CONFIG_PATH):
        fail("config/data-paths.json not found.")
        sys.exit(1)

    with open(CONFIG_PATH) as f:
        cfg = json.load(f)

    root = cfg.get("google_drive_root", "")
    inputs_sub = cfg.get("inputs_subfolder", "inputs")
    outputs_sub = cfg.get("outputs_subfolder", "outputs")

    all_good = True

    if not root or root.startswith("REPLACE_ME"):
        print("[INFO] google_drive_root is not configured -- falling back to local data/inputs "
              "and data/outputs. This is fine for a first test run with the sample data, but real "
              "files won't sync to your phone/browser until you set this.")
        inputs_path = os.path.join(ROOT, "data", "inputs")
        outputs_path = os.path.join(ROOT, "data", "outputs")
    else:
        inputs_path = os.path.join(root, inputs_sub)
        outputs_path = os.path.join(root, outputs_sub)

    for label, path in [("inputs", inputs_path), ("outputs", outputs_path)]:
        if os.path.isdir(path):
            all_good = ok(f"{label} folder exists: {path}")
        else:
            all_good = fail(f"{label} folder does not exist yet: {path} "
                             f"(create it, or check for a typo in data-paths.json)")

        if os.path.isdir(path):
            test_file = os.path.join(path, ".write_test")
            try:
                with open(test_file, "w") as f:
                    f.write("ok")
                os.remove(test_file)
                ok(f"{label} folder is writable")
            except Exception as e:
                all_good = fail(f"{label} folder is not writable: {e}")

    sample_dir = os.path.join(ROOT, "sample-data", "sample-test-data")
    sample_files = [f for f in os.listdir(sample_dir) if f.endswith(".xlsx")] if os.path.isdir(sample_dir) else []
    if len(sample_files) == 1:
        ok(f"Sample test workbook present: {sample_files[0]}")
    else:
        fail(f"Expected exactly 1 sample test workbook, found {len(sample_files)} in {sample_dir}")
        all_good = False

    skill_path = os.path.join(ROOT, ".claude", "skills", "business-review-automation", "SKILL.md")
    if os.path.exists(skill_path):
        ok("Skill file present: .claude/skills/business-review-automation/SKILL.md")
    else:
        fail("Skill file missing -- the project may not have unzipped correctly.")
        all_good = False

    print()
    if all_good:
        print("Setup looks good. In VS Code, ask Claude Code to run a business review cycle.")
    else:
        print("Fix the items marked [FAIL] above, then re-run this script.")
    sys.exit(0 if all_good else 1)


if __name__ == "__main__":
    main()
