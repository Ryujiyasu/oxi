# -*- coding: utf-8 -*-
"""Per-paragraph COLUMN WIDTH in a vertical page, paired with what might decide it.

`educational__015355870669f8d3` p7 is the residual after S1290: Word lays it out
in 18pt units and gives each paragraph 1, 2 or 3 of them, while Oxi allocates a
different mix and the page drifts. S1185 already derived the width as
`ceil(natural_line_height / grid_pitch)` cells; this reads, per paragraph, the
width Word actually used next to the inputs that rule takes -- font size, the
run's font, and the paragraph's line-spacing rule -- so the discriminator is
read off rather than guessed.

★The x STEP to the next paragraph is NOT the width. It is `width x columns`, and
for a paragraph that wraps, the column count is unknown from the start position
alone -- which is how the first reading of this page came out contradicting
S1185 in BOTH directions at once. So this walks the paragraph's own characters
and records the distinct column x's it occupies: the spacing between a
paragraph's OWN columns is its width, measured without knowing how many it took,
and `step / width` then says how many it used. An empty paragraph has no
characters and so no self-measurement -- its step is all there is, and it is
only a width if it took one column.

    python _vert_colwidth_probe.py <docx> <page>
"""
import os
import sys

import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

wdActiveEndSectionNumber = 2
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
            # The paragraph's own columns: distinct x over its characters, in
            # layout order (right to left). Their spacing is the column width.
            cols = []
            for pos in range(rng.Start, rng.End):
                if (doc.Range(pos, pos + 1).Text or "") in ("\r", "\x07", ""):
                    continue
                x = round(float(doc.Range(pos, pos).Information(
                    wdHorizontalPositionRelativeToPage)), 2)
                if x not in cols:
                    cols.append(x)
            rows.append({
                "i": i,
                "x": round(float(st.Information(wdHorizontalPositionRelativeToPage)), 2),
                "y": round(float(st.Information(wdVerticalPositionRelativeToPage)), 2),
                "sec": int(st.Information(wdActiveEndSectionNumber)),
                "size": float(f.Size) if f.Size != 9999999 else None,
                "name_fe": f.NameFarEast,
                "name_ascii": f.NameAscii,
                "rule": int(p.LineSpacingRule),
                "spacing": round(float(p.LineSpacing), 2),
                "before": round(float(p.SpaceBefore), 2),
                "after": round(float(p.SpaceAfter), 2),
                "cols": cols,
                "n_chars": len(rng.Text.rstrip("\r\x07")),
                "text": rng.Text.rstrip("\r\x07")[:20],
            })
        print(f"  {'i':>4} {'sec':>3} {'x':>8} {'y':>7} {'step':>5} {'ncol':>4} "
              f"{'width':>6} {'chars':>5} {'size':>5} {'rule':>4} {'sp':>6} "
              f"{'bef/aft':>9}  font / text")
        for k, r in enumerate(rows):
            step = None if k + 1 >= len(rows) else r["x"] - rows[k + 1]["x"]
            c = r["cols"]
            # Width from the paragraph's own columns; a 1-column paragraph
            # cannot say, and falls back to the step (which is then the width).
            width = min(c[j] - c[j + 1] for j in range(len(c) - 1)) if len(c) > 1 else None
            ncol = len(c) if c else 0
            if width is None and step is not None and ncol <= 1:
                width, ncol = step, 1
            print(f"  {r['i']:>4} {r['sec']:>3} {r['x']:>8.2f} {r['y']:>7.2f} "
                  f"{('%.0f' % step) if step is not None else '':>5} {ncol:>4} "
                  f"{('%.2f' % width) if width is not None else '':>6} "
                  f"{r['n_chars']:>5} {str(r['size']):>5} {r['rule']:>4} "
                  f"{r['spacing']:>6} {r['before']:>4}/{r['after']:<4} "
                  f" {r['name_fe']}/{r['name_ascii']}  {r['text']!r}")
    finally:
        doc.Close(False)
        app.Quit()


if __name__ == "__main__":
    main()
