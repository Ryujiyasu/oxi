# -*- coding: utf-8 -*-
"""Read the FIT-EDGE probe: which arm is the first to break the line?

No export needed -- `Paragraphs(1).Lines.Count` is PowerPoint's own answer, and
the arms differ only in box width.

Usage: python tools/metrics/read_pptx_fitedge_com.py
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import win32com.client

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "pptx_probes" / "fitedge"


def main() -> None:
    arms = json.loads((OUT / "arms.json").read_text(encoding="utf-8"))
    app = win32com.client.Dispatch("PowerPoint.Application")
    rows = []
    try:
        pres = app.Presentations.Open(str((OUT / "probe_fitedge.pptx").resolve()),
                                      WithWindow=False)
        try:
            print(f"{'arm':>4}{'box pt':>11}{'box-text':>11}{'in mu':>9}  lines")
            for a in arms:
                tr = pres.Slides(a["slide"]).Shapes(1).TextFrame.TextRange
                n = tr.Paragraphs(1).Lines().Count
                slack = a["box_pt"] - a["text_pt"]
                rows.append((a["k16"], slack, n))
                print(f"{a['slide']:>4}{a['box_pt']:11.4f}{slack:+11.4f}"
                      f"{slack * 8:+9.3f}  {n}")
        finally:
            pres.Saved = True
            pres.Close()
    finally:
        app.Quit()
    whole = [s for _, s, n in rows if n == 1]
    broke = [s for _, s, n in rows if n > 1]
    if whole and broke:
        print(f"\n  breaks up to slack {max(broke):+.4f}pt "
              f"({max(broke) * 8:+.3f} master units)")
        print(f"  stays whole from  {min(whole):+.4f}pt "
              f"({min(whole) * 8:+.3f} master units)")
        print("  ★an inclusive test would stay whole at slack 0.000; a strict one "
              "would need slack > 0")


if __name__ == "__main__":
    main()
