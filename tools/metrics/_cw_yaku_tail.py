# -*- coding: utf-8 -*-
"""Does Word spend a cell line's half-em when the break is CHEAP?

Every arm so far broke at the end of a five- or twelve-character string, so the
alternative to compressing was always a one-character last line -- and Word always
compressed. tokyoshugyo says otherwise: Word breaks 「…確認するこ」/「と。」, leaving a
two-character line, where Oxi spends the pool and fits it. So the pool's SIZE is not
the open question any more; when Word chooses to spend it is.

This arm makes the break cheap. The text is long enough that a break leaves a dozen
characters behind, and the sweep asks only how many characters Word puts on the
FIRST line. If that never exceeds what fits naturally, compression is not a general
packing device at all -- it is something Word reaches for only near the end.

    python _cw_yaku_tail.py
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
ARMS = {
    "T1": "甲亜、亜亜、亜亜、亜亜、亜亜、亜亜、亜亜、亜亜、亜亜、亜亜、",
    "T2": "甲亜亜亜亜亜亜亜亜亜亜亜亜亜亜亜亜亜亜亜亜亜亜亜亜亜亜亜亜亜",
}
NOTE = {
    "T1": "読点 every third character -- the pool is always available",
    "T2": "no 約物 anywhere -- the control",
}


def main():
    os.makedirs(L.OUT, exist_ok=True)
    # content area 40..80pt: the first line takes 3 to 7 characters, and a break
    # always leaves 20+ behind.
    L.WINDOWS = [((40.0 + 10.8) * 20, (80.0 + 10.8) * 20, 1.0)]
    ws = L.widths()
    print(f"sweep {len(ws)} widths, content area 40..80pt, em {EM}")
    for name, text in ARMS.items():
        docx = os.path.join(L.OUT, f"cwyt_{name}.docx")
        if not os.path.exists(docx):
            P._build_exact(docx, text)
        cells = P.measure(L.export(docx), text)
        if len(cells) != len(ws):
            print(f"{name}: {len(cells)} cells vs {len(ws)} widths -- skipped")
            continue
        over = {}
        for w, lines in zip(ws, cells):
            inner = w / 20.0 - 10.8
            n = len(lines[0])
            natural_fit = int((w - 216) // (L.SZ * 10))     # in twips, as Word does it
            over[n - natural_fit] = over.get(n - natural_fit, 0) + 1
        print(f"\n=== {name} -- {NOTE[name]}")
        print(f"    first-line characters minus what fits naturally: "
              f"{dict(sorted(over.items()))}")


if __name__ == "__main__":
    main()
