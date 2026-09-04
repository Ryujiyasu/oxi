# -*- coding: utf-8 -*-
"""Ask PowerPoint how many lines each width holds, and report the boundary.

For each case the narrowest width whose `Lines.Count` is 1 is PowerPoint's own
measurement of that line, to within the master unit it measures in. Beside it,
the engine's number for the same text, so the two can be compared as numbers.

    python tools/metrics/read_pptx_breakwidth_com.py
"""
from __future__ import annotations

import json
import os
import sys

SRC = os.path.abspath(os.path.join("tools", "metrics", "breakwidth.pptx"))
PLAN = os.path.join("tools", "metrics", "breakwidth.json")

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def main() -> None:
    import win32com.client

    plan = json.load(open(PLAN, encoding="utf-8"))
    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Open(SRC, WithWindow=False)
    counts = {}
    try:
        for si in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(si)
            for k in range(slide.Shapes.Count):
                shape = slide.Shapes(k + 1)
                try:
                    tr = shape.TextFrame.TextRange
                    n = tr.Paragraphs(1).Lines().Count
                    w = shape.Width
                except Exception:
                    continue
                counts[(si, k)] = (n, round(float(w), 3))
    finally:
        pres.Close()
        app.Quit()

    by_label = {}
    for arm in plan:
        key = (arm["slide"], arm["box"])
        if key not in counts:
            continue
        n, w = counts[key]
        by_label.setdefault(arm["label"], []).append((arm["width"], n, arm))

    for label, rows in by_label.items():
        rows.sort()
        one = [w for w, n, _ in rows if n == 1]
        two = [w for w, n, _ in rows if n > 1]
        arm = rows[0][2]
        print("%-11s %s %.1fpt  %r" % (label, arm["face"], arm["size"], arm["text"][:46]))
        if one and two:
            print("             widest that WRAPS %.3f pt, narrowest that FITS %.3f pt"
                  % (max(two), min(one)))
            print("             => PowerPoint measures the line at %.3f pt "
                  "(%.1f master units)" % (min(one), min(one) / 0.125))
        else:
            print("             every arm gave %s line(s) -- widen the sweep"
                  % ({n for _, n, _ in rows}))


if __name__ == "__main__":
    main()
