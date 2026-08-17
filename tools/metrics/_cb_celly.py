# -*- coding: utf-8 -*-
"""When the line breaks match Word but the pixels get worse, where does it go?

`OXI_CELLPAD=0.5` takes 34140b to 291/291 lines -- every line the same characters
as Word -- and its SSIM still drops 0.0067. So the remaining error is not what
each line CONTAINS but where each line SITS. This joins the two engines' lines by
text and reports the vertical offset, per page, for whichever arm is asked.

★Oxi's `element.y` is the LINE BOX top and the glyph is drawn at `y +
text_y_off`, while Word's PDF line bbox is the glyph box, so the join adds the
offset before comparing -- otherwise it invents a difference that grows with the
line height ([[feedback_text_y_vs_info6]]).

    python _cb_celly.py 34140b                       # default arm
    python _cb_celly.py 34140b --arm=OXI_CELLPAD=0.5
"""
import collections
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
sys.stdout.reconfigure(encoding="utf-8")

import _cb_budget as B  # noqa: E402


def oxi_glyph_lines(layout):
    out = []
    for pi, page in enumerate(layout["pages"], 1):
        rows = collections.OrderedDict()
        for e in page["elements"]:
            if e["type"] != "text" or not (e.get("text") or "").strip():
                continue
            y = e["y"] + (e.get("text_y_off") or 0.0)
            key = (round(y, 1), e.get("cell_row_idx"), e.get("cell_col_idx"))
            rows.setdefault(key, []).append(e)
        for (y, _r, _c), els in rows.items():
            els.sort(key=lambda e: e["x"])
            out.append({"page": pi, "y": y,
                        "x0": min(e["x"] for e in els),
                        "text": "".join(e["text"] for e in els)})
    return out


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    envs = ""
    for a in sys.argv[1:]:
        if a.startswith("--arm="):
            envs = a[len("--arm="):]
    docx = B.docx_for(args[0] if args else "34140b")
    _recs, layout = B.run_oxi(docx, tag="celly", envs=envs)
    oxi = oxi_glyph_lines(layout)
    word = B.word_lines(docx)

    wby = collections.defaultdict(list)
    for w in word:
        wby[(w["page"], B.norm(w["text"]))].append(w)
    rows = []
    for o in oxi:
        k = (o["page"], B.norm(o["text"]))
        if wby.get(k):
            w = wby[k].pop(0)
            rows.append((o["page"], o["y"], w["y"], o["x0"], w["x0"], B.norm(o["text"])[:14]))
    print("== %s ==  arm=%s   joined %d lines of %d"
          % (os.path.basename(docx)[:30], envs or "default", len(rows), len(oxi)))
    per = collections.defaultdict(list)
    for pg, oy, wy, ox, wx, _t in rows:
        per[pg].append((oy - wy, ox - wx))
    print("%-4s %-6s %-9s %-9s %-9s %s" % ("pg", "n", "dy med", "dy min", "dy max", "dx med"))
    for pg in sorted(per):
        v = per[pg]
        dys = sorted(d for d, _ in v)
        dxs = sorted(x for _, x in v)
        print("%-4d %-6d %-9.2f %-9.2f %-9.2f %.2f"
              % (pg, len(v), dys[len(dys) // 2], dys[0], dys[-1], dxs[len(dxs) // 2]))
    worst = sorted(rows, key=lambda r: -abs(r[1] - r[2]))[:6]
    print("\nlargest vertical offsets:")
    for pg, oy, wy, _ox, _wx, t in worst:
        print("   p%-2d oxi %.1f  word %.1f  dy %+.2f  %s" % (pg, oy, wy, oy - wy, t))


if __name__ == "__main__":
    main()
