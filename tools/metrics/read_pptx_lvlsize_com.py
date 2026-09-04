# -*- coding: utf-8 -*-
"""Ask PowerPoint what size each run of the lvlsize repro resolved to.

    python tools/metrics/read_pptx_lvlsize_com.py
"""
from __future__ import annotations

import argparse
import json
import os
import sys

ap = argparse.ArgumentParser()
ap.add_argument("--plan", default="lvlsize", help="basename: lvlsize / lvlsize2")
ARGS = ap.parse_args()
SRC = os.path.abspath(os.path.join("tools", "metrics", ARGS.plan + ".pptx"))
PLAN = os.path.join("tools", "metrics", ARGS.plan + ".json")

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def main() -> None:
    import win32com.client

    plan = json.load(open(PLAN, encoding="utf-8"))
    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Open(SRC, WithWindow=False)
    try:
        for arm in plan:
            slide = pres.Slides(arm["slide"])
            # The placeholder that holds the arm's three runs.
            hit = None
            for i in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(i)
                try:
                    if "alpha" in shape.TextFrame.TextRange.Text:
                        hit = shape
                        break
                except Exception:
                    continue
            if hit is None:
                print("%-22s (no shape found)" % arm["label"])
                continue
            para = hit.TextFrame.TextRange.Paragraphs(1)
            runs = para.Runs()
            got = [float(para.Runs(j).Font.Size) for j in range(1, runs.Count + 1)]
            print("%-22s declared %-18s -> resolved %s"
                  % (arm["label"], arm["sizes"], got))
    finally:
        pres.Close()
        app.Quit()


if __name__ == "__main__":
    main()
