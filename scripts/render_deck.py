#!/usr/bin/env python3
"""
Phase 8 visual QA -- export every slide of a deck to PNG so Claude can look at each one
(Read the PNGs) for overflow, overlaps, unreadable charts, and leftover placeholder text.

  python scripts/render_deck.py <path-to-deck.pptx> [--out <folder>]

Default output folder: _deck_preview/<deck name>/ next to the deck (underscore-prefixed so the
input scanner and QA globbing ignore it). Exit code 3 = PowerPoint automation unavailable.
"""
import argparse
import os
import sys


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("path")
    ap.add_argument("--out")
    ap.add_argument("--width", type=int, default=1600)
    args = ap.parse_args()
    path = os.path.abspath(args.path)
    if not os.path.exists(path):
        print(f"[FAIL] Not found: {path}")
        sys.exit(1)
    stem = os.path.splitext(os.path.basename(path))[0]
    out = os.path.abspath(args.out or os.path.join(os.path.dirname(path), "_deck_preview", stem))
    os.makedirs(out, exist_ok=True)
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        print("[SKIP] pywin32 not available -- render the deck manually for visual QA.")
        sys.exit(3)

    pythoncom.CoInitialize()
    ppt = pres = None
    try:
        try:
            ppt = win32.Dispatch("PowerPoint.Application")
        except Exception as e:
            print(f"[SKIP] Desktop PowerPoint not available via COM ({e}).")
            sys.exit(3)
        pres = ppt.Presentations.Open(path, ReadOnly=True, Untitled=False, WithWindow=False)
        w, h = pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight
        height = int(args.width * h / w)
        n = pres.Slides.Count
        for i in range(1, n + 1):
            pres.Slides(i).Export(os.path.join(out, f"slide_{i:02d}.png"), "PNG", args.width, height)
    finally:
        if pres is not None:
            pres.Close()
        if ppt is not None and ppt.Presentations.Count == 0:
            ppt.Quit()  # leave PowerPoint running if the user had their own decks open
        pythoncom.CoUninitialize()
    print(f"Exported {n} slide image(s) to {out}")


if __name__ == "__main__":
    main()
