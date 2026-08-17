# -*- coding: utf-8 -*-
"""What advance does Word give a PROPORTIONAL Japanese glyph, before justification?

The 10tw (0.5pt) rounding in `char_width_pt_with_gdi_map` is COM-confirmed on
Cambria, TNR, Arial, Calibri, Century, MS Mincho and MS Gothic -- every one of
which is either Latin or fullwidth, so the rounding is invisible (a full em is
always a multiple of 10tw). No proportional CJK face was ever checked, and
c7b923e5 loses a character per line by 0.20pt.

Reading the advances off a justified line does not answer it: justification moves
every glyph, and the same character comes back as 0.855em on one line and 0.841em
on another. But the LAST line of a justified paragraph is not justified -- Word
lays it out ragged, at natural advances. So this collects advances from
paragraph-final lines only (a line whose right edge falls well short of the
column) and puts Word's number next to the two Oxi could use: the design em from
the metrics table, and that em rounded to 10tw.

    python _cb_pgothic_adv.py                      # c7b923e5, MS PGothic
    python _cb_pgothic_adv.py kojin                # another document
"""
import collections
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ.setdefault("PYTHONIOENCODING", "utf-8")
sys.stdout.reconfigure(encoding="utf-8")

import _cb_budget as B  # noqa: E402

METRICS = os.path.join(B.REPO, "crates", "oxidocs-core", "src", "font", "data",
                       "font_metrics_compact.json")
# A line ending this far short of the page's widest line is ragged, i.e. the
# final line of its paragraph, i.e. laid out at natural advances.
RAGGED_MARGIN = 12.0


def is_cjk(ch):
    cp = ord(ch)
    return (0x3000 <= cp <= 0x30FF or 0x4E00 <= cp <= 0x9FFF
            or 0xFF00 <= cp <= 0xFFEF or 0x3400 <= cp <= 0x4DBF)


def design_table(family):
    for f in json.load(open(METRICS, encoding="utf-8")):
        if f["family"] == family:
            return f["units_per_em"], f.get("widths", {})
    return None, {}


def main():
    prefix = sys.argv[1] if len(sys.argv) > 1 else "c7b923e5"
    docx = B.docx_for(prefix)
    import fitz
    rt = docx[:-5] + "_rt.pdf"
    if not os.path.exists(rt):
        sys.exit("no Word PDF next to %s" % os.path.basename(docx))
    pdf = fitz.open(rt)

    adv = collections.defaultdict(list)
    faces = collections.Counter()
    for pg in pdf:
        lines = []
        for b in pg.get_text("rawdict")["blocks"]:
            for ln in b.get("lines", []):
                chars = [c for s in ln["spans"] for c in s["chars"]]
                if not "".join(c["c"] for c in chars).strip():
                    continue
                lines.append((ln, chars))
        if not lines:
            continue
        edge = max(ln["bbox"][2] for ln, _ in lines)
        for ln, chars in lines:
            if ln["bbox"][2] > edge - RAGGED_MARGIN:
                continue                      # justified: the advances are stretched
            # ★Pairs must stay INSIDE one span: a heading run in a fullwidth face
            # sits in the same line list as the body, and mixing the two reported
            # の at a flat 1.000em -- the fullwidth face's number, not this one's.
            for s in ln["spans"]:
                sc = s["chars"]
                faces[s["font"]] += len(sc)
                for i, c in enumerate(sc[:-1]):
                    nxt = sc[i + 1]["c"]
                    # ★Both sides must be CJK. A CJK glyph followed by Latin
                    # carries autoSpaceDE's quarter-em gap, which read as の at
                    # 1.25-1.45em and dragged its median to a flat 1.000.
                    if not (is_cjk(c["c"]) and is_cjk(nxt)):
                        continue
                    a = (sc[i + 1]["bbox"][0] - c["bbox"][0]) / s["size"]
                    if not 0.2 < a < 1.05:      # a gap this wide is not an advance
                        continue
                    adv[(s["font"], c["c"])].append(a)

    face = faces.most_common(1)[0][0] if faces else "?"
    family = {"MS-PGothic": "MS PGothic", "MS-PMincho": "MS PMincho",
              "MS-Mincho": "MS Mincho", "MS-Gothic": "MS Gothic"}.get(face, face)
    upm, widths = design_table(family)
    print("== %s ==  ragged (paragraph-final) lines only; face=%s upm=%s"
          % (os.path.basename(docx)[:34], face, upm))
    print("%-4s %-5s %-9s %-9s %-9s %-9s %s"
          % ("ch", "n", "word_em", "design_em", "d-w", "rounded@10.5", "r-w@10.5"))
    rows = []
    for (fnt, ch), v in adv.items():
        if fnt != face or len(v) < 3 or ch.isspace():
            continue
        v = sorted(v)
        we = v[len(v) // 2]
        de = widths.get(str(ord(ch)))
        de = None if de is None or not upm else de / upm
        # what Oxi's 10tw rounding produces at 10.5pt, and what Word's own
        # advance would be at the same size
        rnd = None if de is None else (int(de * 10.5 * 20.0 / 10.0 + 0.5) * 10.0) / 20.0
        rows.append((abs((de - we) if de else 0), ch, len(v), we, de, rnd, we * 10.5))
    rows.sort(reverse=True)
    for _k, ch, n, we, de, rnd, wpt in rows[:28]:
        print("%-4s %-5d %-9.4f %-9s %-9s %-9s %s"
              % (ch, n, we,
                 "%.4f" % de if de else "-",
                 "%+.4f" % (de - we) if de else "-",
                 "%.2f" % rnd if rnd else "-",
                 "%+.2f" % (rnd - wpt) if rnd else "-"))
    if rows:
        dd = [de - we for _k, _c, _n, we, de, _r, _w in rows if de]
        rr = [rnd - wpt for _k, _c, _n, _we, de, rnd, wpt in rows if de]
        print("\nn=%d chars | design-vs-Word mean %+.4f em (max |%.4f|)"
              % (len(rows), sum(dd) / len(dd), max(abs(x) for x in dd)))
        print("rounded-vs-Word mean %+.3f pt at 10.5 (max |%.3f|) -- "
              "the per-character cost of the 10tw rule on this face"
              % (sum(rr) / len(rr), max(abs(x) for x in rr)))


if __name__ == "__main__":
    main()
