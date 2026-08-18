# -*- coding: utf-8 -*-
"""How much space does Word put at a CJK/Latin boundary, and where does the
"API 18.76 vs 18.00" of c7b923e5 actually go?

`_cw_latin_adv.py` says the per-character part of the error is small (MS
PGothic, com_tw vs Word: +0.02..0.09pt per glyph), far short of the +0.76 the
archive measured across three letters. So the rest is the boundary -- the
autoSpaceDE gap Word opens between a CJK glyph and the Latin run next to it.

It is measured the same way, from the pen: the advance Word gives the CJK glyph
that PRECEDES a Latin one, minus that glyph's own em, is the gap it opened.
`_cw_spacing.py` put the gap near a quarter em from a self-authored repro only;
this asks the real documents.

    python _cw_latin_gap.py c7b923e5 d77a58 04b88e 34140b
"""
import collections
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import _cb_budget as B                       # noqa: E402
from _cw_latin_adv import (is_latin, is_cjk, family_of, load_tables,  # noqa: E402
                           nominal_size, RAGGED_MARGIN)


def main():
    import fitz
    metrics, _com = load_tables()
    for prefix in (sys.argv[1:] or ["c7b923e5"]):
        docx = B.docx_for(prefix)
        rt = docx[:-5] + "_rt.pdf"
        if not os.path.exists(rt):
            continue
        # (font, size, direction) -> [gap in em]
        gaps = collections.defaultdict(list)
        for pg in fitz.open(rt):
            lines = []
            for b in pg.get_text("rawdict")["blocks"]:
                for ln in b.get("lines", []):
                    if "".join(c["c"] for s in ln["spans"]
                               for c in s["chars"]).strip():
                        lines.append(ln)
            if not lines:
                continue
            edge = max(ln["bbox"][2] for ln in lines)
            for ln in lines:
                if ln["bbox"][2] > edge - RAGGED_MARGIN:
                    continue                  # justified: the gap is stretched
                for s in ln["spans"]:
                    sc = s["chars"]
                    for i in range(len(sc) - 1):
                        a, b2 = sc[i]["c"], sc[i + 1]["c"]
                        d = (sc[i + 1]["origin"][0] - sc[i]["origin"][0]) / s["size"]
                        if is_cjk(a) and is_latin(b2):
                            # a is fullwidth: whatever exceeds 1em is the gap
                            gaps[(s["font"], round(s["size"], 2), "CJK>lat")].append(d - 1.0)
                        elif is_latin(a) and is_cjk(b2):
                            fam = family_of(s["font"])
                            upm, w = metrics.get(fam, (None, {}))
                            em = w.get(str(ord(a)))
                            if em is None or not upm:
                                continue
                            gaps[(s["font"], round(s["size"], 2), "lat>CJK")].append(
                                d - em / upm)
        print("\n== %s ==" % os.path.basename(docx)[:40])
        print("%-14s %-7s %-9s %-4s %-8s %-8s %s"
              % ("font", "size", "dir", "n", "med_em", "p25", "p75"))
        for (fnt, fs, dr), v in sorted(gaps.items(), key=lambda kv: -len(kv[1])):
            if len(v) < 4:
                continue
            v = sorted(v)
            print("%-14s %-7.1f %-9s %-4d %-8.4f %-8.4f %.4f"
                  % (fnt[:14], nominal_size(fs), dr, len(v),
                     v[len(v) // 2], v[len(v) // 4], v[3 * len(v) // 4]))


if __name__ == "__main__":
    main()
