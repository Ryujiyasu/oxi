# -*- coding: utf-8 -*-
"""Word COM truth around an orphan/widow page boundary.

Prints page + vertical position for a window of paragraphs around a text
prefix, so a KEEP/PUSH decision can be compared against Oxi's cursor rather
than assumed.  Collapsed start ranges throughout (the R30 fix: a range's own
Information() reports the ACTIVE END, which lands on the next page for a
paragraph whose trailing marker overflows).

  python _orphan_boundary_com.py <docx> "<text prefix>" [window]
"""
import os
import sys

import win32com.client as w

sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def main() -> None:
    path = os.path.abspath(sys.argv[1])
    needle = sys.argv[2]
    win = int(sys.argv[3]) if len(sys.argv) > 3 else 6

    app = w.DispatchEx("Word.Application")
    app.Visible = False
    d = app.Documents.Open(path, ReadOnly=True)
    try:
        d.Repaginate()
        n = d.Paragraphs.Count
        texts = []
        for i in range(1, n + 1):
            texts.append(d.Paragraphs(i).Range.Text.replace("\r", "").replace("\x07", ""))
        hits = [i for i, t in enumerate(texts, 1) if t.startswith(needle)]
        if not hits:
            print("no match")
            return
        print(f"{len(hits)} hit(s): {hits}")
        for h in hits:
            print(f"--- hit at paragraph {h} ---")
            lo, hi = max(1, h - win), min(n, h + win)
            for i in range(lo, hi + 1):
                rng = d.Paragraphs(i).Range
                c = d.Range(rng.Start, rng.Start)
                e = d.Range(rng.End - 1, rng.End - 1)
                mark = ">>" if i == h else "  "
                print(f"{mark} i={i:5d} page={c.Information(3):3d} y={c.Information(6):8.2f} "
                      f"endpage={e.Information(3):3d} endy={e.Information(6):8.2f} "
                      f"{texts[i - 1][:44]!r}")
    finally:
        d.Close(False)
        app.Quit()


if __name__ == "__main__":
    main()
