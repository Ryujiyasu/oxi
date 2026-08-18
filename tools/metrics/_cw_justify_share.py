# -*- coding: utf-8 -*-
"""When a CJK line is justified, which gaps take the slack?

c7b923e5 p2 puts the whole of its remaining error in one place: Word gives the
の before "API" an advance of 14.23 where Oxi gives 13.20, while the three Latin
letters themselves agree to 0.03pt across all three. So the residual is not the
Latin run's width at all -- it is that Oxi spreads a justified line's slack
evenly over every character, and Word appears not to.

For every justified line this measures each character's advance against its
natural one (design em, plus a quarter em where a CJK glyph is followed by a
Latin one, which `_cw_latin_gap.py` measured on ragged lines) and reports the
surplus by class. If Word loaded the slack onto the CJK/Latin boundaries, the
boundary class carries it and the CJK class sits near zero.

    python _cw_justify_share.py c7b923e5 d77a58 tokyoshugyo
"""
import collections
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import _cb_budget as B                                            # noqa: E402
from _cw_latin_adv import (is_latin, is_cjk, family_of, load_tables,  # noqa: E402
                           nominal_size, RAGGED_MARGIN)

YAKU = "、。，．・：；！？）」』】〉》”’（「『【〈《“‘"


def klass(ch, nxt):
    if is_cjk(ch) and is_latin(nxt):
        return "CJK>lat"
    if is_latin(ch) and is_cjk(nxt):
        return "lat>CJK"
    if ch in YAKU:
        return "yakumono"
    if is_cjk(ch):
        return "CJK"
    if is_latin(ch):
        return "latin"
    return "other"


def main():
    import fitz
    metrics, _ = load_tables()
    for prefix in (sys.argv[1:] or ["c7b923e5"]):
        docx = B.docx_for(prefix)
        rt = docx[:-5] + "_rt.pdf"
        if not os.path.exists(rt):
            continue
        share = collections.defaultdict(list)
        nlines = 0
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
                if ln["bbox"][2] < edge - 1.0:
                    continue                     # ragged: no slack to share
                nlines += 1
                for s in ln["spans"]:
                    fam = family_of(s["font"])
                    upm, w = metrics.get(fam, (None, {}))
                    if not upm:
                        continue
                    nom = nominal_size(round(s["size"], 2))
                    k = nom / s["size"]
                    sc = s["chars"]
                    for i in range(len(sc) - 1):
                        ch, nxt = sc[i]["c"], sc[i + 1]["c"]
                        if ch.isspace():
                            continue
                        em = w.get(str(ord(ch)))
                        if em is None:
                            continue
                        nat = em / upm * nom
                        if is_cjk(ch) and is_latin(nxt):
                            nat += 0.25 * nom          # the measured boundary
                        got = (sc[i + 1]["origin"][0] - sc[i]["origin"][0]) * k
                        if abs(got - nat) > 0.6 * nom:
                            continue                   # a tab or a span seam
                        share[klass(ch, nxt)].append(got - nat)
        print("\n== %s ==  %d justified lines" % (os.path.basename(docx)[:34], nlines))
        print("%-9s %-6s %-9s %-9s %s" % ("class", "n", "med_extra", "mean", "p90"))
        for cls, v in sorted(share.items(), key=lambda kv: -len(kv[1])):
            if len(v) < 5:
                continue
            v = sorted(v)
            print("%-9s %-6d %-9.3f %-9.3f %.3f"
                  % (cls, len(v), v[len(v) // 2], sum(v) / len(v),
                     v[int(0.9 * (len(v) - 1))]))


if __name__ == "__main__":
    main()
