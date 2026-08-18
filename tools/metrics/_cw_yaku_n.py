# -*- coding: utf-8 -*-
"""Does a cell line's compression pool keep growing with every 約物 on it?

The first sweep put at most two 約物 on a line and found the pool additive: one
gave half an em of headroom, two gave a whole em. Real text is not like that --
tokyoshugyo's 条文 lines carry nine or more （「」。、） each, so an additive pool
would be four and a half ems of slack, and Oxi packing to it runs 37 paragraphs
ahead of Word.

So sweep the COUNT. Each arm is the same length of text with N of its characters
replaced by a closing bracket, and the cell width walks from "everything fits" to
"short by more than N half-ems". The demand at which the line finally breaks is
the pool, and plotting it against N says whether it is N x 0.5em or saturates.

    python _cw_yaku_n.py            # generate, export through Word, measure
    python _cw_yaku_n.py --keep
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
NCH = 12                     # characters per line, fixed across arms
STEP_TW = 2.0                # 0.1pt -- fine enough to place the break, cheap enough


def text_for(n):
    """甲 + (NCH-1) characters, n of them closing brackets, spread out."""
    body = ["亜"] * (NCH - 1)
    for k in range(n):
        body[(k * (NCH - 1)) // max(n, 1)] = "）"
    return "甲" + "".join(body)


def main():
    os.makedirs(L.OUT, exist_ok=True)
    keep = "--keep" in sys.argv
    natural = NCH * EM
    print(f"{NCH} characters = {natural:.1f}pt natural, sweep step {STEP_TW / 20:.2f}pt")
    print(f"  {'約物':>4}{'pool if additive':>18}{'breaks at demand':>18}{'= em':>8}"
          f"{'per 約物':>10}")
    for n in (0, 1, 2, 3, 4, 6, 8):
        text = text_for(n)
        pool = n * 0.5 * EM
        # from one em of slack to one em past an additive pool
        L.WINDOWS = [((natural - pool - EM + 10.8) * 20,
                      (natural + EM + 10.8) * 20, STEP_TW)]
        docx = os.path.join(L.OUT, f"cwyn_{n}.docx")
        if not (keep and os.path.exists(docx)):
            P._build_exact(docx, text)
        cells = P.measure(L.export(docx), text)
        ws = L.widths()
        if len(cells) != len(ws):
            print(f"  {n:>4}  {len(cells)} cells vs {len(ws)} widths -- skipped")
            continue
        held = [w / 20.0 - 10.8 for w, lines in zip(ws, cells)
                if len(lines[0]) >= len(text)]
        if not held:
            print(f"  {n:>4}{pool:>18.2f}{'never holds':>18}")
            continue
        inner = min(held)                       # narrowest cell that still holds
        demand = natural - inner
        print(f"  {n:>4}{pool:>18.2f}{demand:>18.2f}{demand / EM:>8.3f}"
              f"{(demand / n / EM if n else 0):>10.3f}")


if __name__ == "__main__":
    main()
