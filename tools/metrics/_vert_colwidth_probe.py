# -*- coding: utf-8 -*-
"""Per-paragraph COLUMN WIDTH in a vertical page, paired with what might decide it.

`educational__015355870669f8d3` p7 is the residual after S1290: Word lays it out
in 18pt units and gives each paragraph 1, 2 or 3 of them, while Oxi allocates a
different mix and the page drifts. S1185 already derived the width as
`ceil(natural_line_height / grid_pitch)` cells; this reads, per paragraph, the
width Word actually used next to the inputs that rule takes -- font size, the
run's font, and the paragraph's line-spacing rule -- so the discriminator is
read off rather than guessed.

The width is the x STEP to the next paragraph, which is only the paragraph's own
width when it occupies ONE column; where it wraps, the step is width x columns.
So the per-character column walk (`_vert_band_columns.py`) is what says how many
columns a paragraph used, and this adds the properties beside it.

    python _vert_colwidth_probe.py <docx> <page>
"""
import os
import sys

import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

wdActiveEndPageNumber = 3
wdHorizontalPositionRelativeToPage = 5
wdVerticalPositionRelativeToPage = 6


def main():
    path, page = os.path.abspath(sys.argv[1]), int(sys.argv[2])
    app = win32.gencache.EnsureDispatch("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    doc = app.Documents.Open(path, ReadOnly=True, AddToRecentFiles=False)
    try:
        rows = []
        for i, p in enumerate(doc.Paragraphs, 1):
            rng = p.Range
            st = doc.Range(rng.Start, rng.Start)
            if int(st.Information(wdActiveEndPageNumber)) != page:
                continue
            f = rng.Font
            rows.append({
                "i": i,
                "x": round(float(st.Information(wdHorizontalPositionRelativeToPage)), 2),
                "y": round(float(st.Information(wdVerticalPositionRelativeToPage)), 2),
                "size": float(f.Size) if f.Size != 9999999 else None,
                "name_fe": f.NameFarEast,
                "name_ascii": f.NameAscii,
                "rule": int(p.LineSpacingRule),
                "spacing": round(float(p.LineSpacing), 2),
                "style": str(p.Style),
                "n_chars": len(rng.Text.rstrip("\r\x07")),
                "text": rng.Text.rstrip("\r\x07")[:26],
            })
        print(f"  {'i':>4} {'x':>8} {'step':>6} {'chars':>5} {'size':>6} "
              f"{'rule':>4} {'spacing':>8}  font / text")
        for k, r in enumerate(rows):
            step = "" if k + 1 >= len(rows) else f"{r['x'] - rows[k+1]['x']:.0f}"
            print(f"  {r['i']:>4} {r['x']:>8.2f} {step:>6} {r['n_chars']:>5} "
                  f"{str(r['size']):>6} {r['rule']:>4} {r['spacing']:>8} "
                  f" {r['name_fe']}/{r['name_ascii']}  {r['text']!r}")
    finally:
        doc.Close(False)
        app.Quit()


if __name__ == "__main__":
    main()
