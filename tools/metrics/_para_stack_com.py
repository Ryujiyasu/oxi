# -*- coding: utf-8 -*-
"""Per-paragraph Word advance over a range, with the mark font that governs it.

Info6 in a Latin document is quantised to the 96dpi pixel (0.75pt) -- see
_pb_pxrun_gen.py, where a 20-paragraph run proves the CURSOR is exact and only
the reported position is rounded.  So a single advance carries +-0.75 of
reporting noise; a RUN of them still pins the mean.  This dump pairs each
advance with the paragraph's font/size/spacing so an empty-paragraph height
question can be answered per paragraph instead of per document.

  python _para_stack_com.py <docx> <from_index> <to_index>
"""
import os
import sys

import win32com.client as w

sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def main() -> None:
    path = os.path.abspath(sys.argv[1])
    lo, hi = int(sys.argv[2]), int(sys.argv[3])
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    d = app.Documents.Open(path, ReadOnly=True)
    try:
        d.Repaginate()
        hi = min(hi, d.Paragraphs.Count)
        prev = None
        for i in range(lo, hi + 1):
            p = d.Paragraphs(i)
            rng = p.Range
            c = d.Range(rng.Start, rng.Start)
            pg, y = c.Information(3), round(c.Information(6), 2)
            txt = rng.Text.replace("\r", "").replace("\x07", "")
            adv = ""
            if prev and prev[0] == pg:
                adv = "%+7.2f" % (y - prev[1])
            prev = (pg, y)
            print(f"i={i:5d} pg={pg:3d} y={y:8.2f} adv={adv:>8s} "
                  f"sz={rng.Font.Size} font={str(rng.Font.Name)[:18]:18s} "
                  f"rule={p.LineSpacingRule} ls={round(p.LineSpacing, 2)} "
                  f"sb={p.SpaceBefore} sa={p.SpaceAfter} {txt[:38]!r}")
    finally:
        d.Close(False)
        app.Quit()


if __name__ == "__main__":
    main()
