# -*- coding: utf-8 -*-
"""How much aki does each class of 約物 lend a cell line?

The count sweep said the pool caps at half an em however many closing brackets a
line carries; re-reading the long-line sweep by what is being squeezed in said a
comma lends about a quarter em EACH. Both cannot be a single per-line number, and
the body-side `s475_max_compress_pt` has carried per-class caps all along (comma
0.283em, period 0.5em, closing-solo 0.07em) from its own measurements. So sweep the
class here too.

Each arm is the same 30-character line with one class of 約物 every third character,
swept over 801 cell widths. For every width we ask whether Word squeezed in the
character that did not fit naturally, and split the answer by how many 約物 the line
was carrying — the limit that appears is the class's capacity times the count.

    python _cw_yaku_class.py            # generate, export through Word, measure
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import _cw_law as L  # noqa: E402
import _cw_yaku_probe as P  # noqa: E402

EM = L.SZ / 2.0
CLASSES = {"C1": "、", "C2": "。", "C3": "）", "C4": "（", "C5": "・"}
NCH = 30


def text_for(mark):
    body = []
    while len(body) < NCH - 1:
        body += ["亜", "亜", mark]
    return "甲" + "".join(body[:NCH - 1])


def main():
    os.makedirs(L.OUT, exist_ok=True)
    L.WINDOWS = [((40.0 + 10.8) * 20, (80.0 + 10.8) * 20, 1.0)]
    ws = L.widths()
    print("%d widths, content area 40..80pt, em %.1f" % (len(ws), EM))
    for name, mark in CLASSES.items():
        text = text_for(mark)
        docx = os.path.join(L.OUT, "cwyc_%s.docx" % name)
        if not os.path.exists(docx):
            P._build_exact(docx, text)
        cells = P.measure(L.export(docx), text)
        if len(cells) != len(ws):
            print("  %s %r: %d cells vs %d widths -- skipped"
                  % (name, mark, len(cells), len(ws)))
            continue
        # limit per count: the largest shortfall at which Word still took a
        # non-約物 character, given how many 約物 the line already held
        by = {}
        always = [0, 0]
        for w, lines in zip(ws, cells):
            inner = w / 20.0 - 10.8
            n = len(lines[0])
            fit = int((w - 216) // (L.SZ * 10))
            if fit >= len(text):
                continue
            demand = ((fit + 1) * EM - inner) / EM
            took = n > fit
            if text[fit] == mark:                 # the 約物 itself is line-final
                always[0] += 1 if took else 0
                always[1] += 1
                continue
            cnt = text[:fit].count(mark)
            t, c, lim = by.get(cnt, (0, 0, 0.0))
            by[cnt] = (t + (1 if took else 0), c + 1, max(lim, demand if took else 0.0))
        print("\n=== %s  %r" % (name, mark))
        print("    line-final %r taken: %d/%d" % (mark, always[0], always[1]))
        print("    %8s %8s %7s %14s" % ("on line", "taken", "of", "max demand/em"))
        for cnt in sorted(by):
            t, c, lim = by[cnt]
            if c < 10:
                continue
            print("    %8d %8d %7d %14.2f  (per mark %.3f)"
                  % (cnt, t, c, lim, lim / cnt if cnt else 0))


if __name__ == "__main__":
    main()
