# -*- coding: utf-8 -*-
"""Does Word compress 約物 inside a table cell, and by how much?

34140b p2 says it does: 「（　年度）」 at 10.5pt in a cell whose content area is
49.47pt comes out as 10.56 / 10.59 / 10.44 / 10.44 / **7.44** = 49.47 exactly, and
the same string in a 53.07pt cell keeps its closing bracket at 10.44. The bracket
gives up exactly what the line is short by.

An earlier pass concluded the opposite -- "no demand-driven compression in cells" --
from a median over a whole document. Most cells have no demand, so the median is the
answer for the cells that were never under pressure. So condition on demand and
never average across it.

Per line inside a cell this measures:

    demand   = natural width - the cell's content width
    supplied = natural width - the width Word actually drew

and reports supplied against demand, split by which 約物 the line carries.

    python _cw_yaku.py                # every cached Word export
    python _cw_yaku.py 34140b 04b88e  # only these
    python _cw_yaku.py --lines        # also dump the compressed lines
"""
import glob
import os
import re
import sys
import zipfile
from collections import Counter, defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(os.path.dirname(HERE))
DOCS = os.path.join(REPO, "tools", "golden-test", "documents", "docx")
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

CLOSING = "）」』〕】》〉｝］、。，．"
OPENING = "（「『〔【《〈｛［"
MIDDLE = "・：；"
IDEO_SPACE = "　"


def klass(c):
    if c in CLOSING:
        return "closing"
    if c in OPENING:
        return "opening"
    if c in MIDDLE:
        return "middle"
    if c == IDEO_SPACE:
        return "ideospace"
    return "text"


def fullwidth(c):
    o = ord(c)
    return (0x3000 <= o <= 0x303F or 0x3040 <= o <= 0x30FF or 0x4E00 <= o <= 0x9FFF
            or 0x3400 <= o <= 0x4DBF or 0xFF01 <= o <= 0xFF60 or 0xFFE0 <= o <= 0xFFE6)


def cellmar(docx):
    """Left/right cell margin in points. Word's default is 108 twips a side."""
    try:
        x = zipfile.ZipFile(docx).read("word/document.xml").decode("utf-8")
    except Exception:
        return 5.4, 5.4, False
    m = re.search(r"<w:tblCellMar>(.*?)</w:tblCellMar>", x, re.S)
    if not m:
        return 5.4, 5.4, False
    body = m.group(1)
    def side(tag, d):
        mm = re.search(r'<w:%s w:w="(-?\d+)"' % tag, body)
        return int(mm.group(1)) / 20.0 if mm else d
    return side("left", 5.4), side("right", 5.4), True


def rules_of(pg):
    out = []
    for dr in pg.get_drawings():
        for it in dr["items"]:
            if it[0] == "l":
                p, q = it[1], it[2]
                if abs(p.x - q.x) < 0.4 and abs(p.y - q.y) > 1.0:
                    out.append((p.x, min(p.y, q.y), max(p.y, q.y)))
            elif it[0] == "re":
                r = it[1]
                if r.width < 3.0 and r.height > 1.0:
                    out.append((r.x0 + r.width / 2, r.y0, r.y1))
    return out


def lines_of(pdf, marl, marr):
    """Every line that sits between two vertical rules, with its content area."""
    import fitz
    got = []
    for pg in fitz.open(pdf):
        rules = rules_of(pg)
        for b in pg.get_text("rawdict")["blocks"]:
            for ln in b.get("lines", []):
                cs = [(c["c"], c["bbox"][0], sp["size"], sp["font"])
                      for sp in ln["spans"] for c in sp["chars"]]
                if len(cs) < 2:
                    continue
                y0, y1 = ln["bbox"][1], ln["bbox"][3]
                near = sorted({x for x, a, b_ in rules if a <= y0 + 1 and b_ >= y1 - 1})
                left = [x for x in near if x < cs[0][1]]
                right = [x for x in near if x > cs[-1][1]]
                if not left or not right:
                    continue                      # not inside a cell
                # A rule BETWEEN two of the line's characters means PDF extraction
                # welded two cells of the same row into one "line". Their combined
                # text against one cell's width is meaningless.
                if any(cs[0][1] < x < cs[-1][1] for x in near):
                    continue
                got.append((right[0] - left[-1] - marl - marr, cs))
    return got


def natural_table(rows):
    """The advance a (face, size, character) takes when nothing squeezes it.

    Assuming one em is wrong the moment a PROPORTIONAL Japanese face appears -- its
    brackets are half-width by design, and calling that compression put 34pt of
    phantom squeeze on lines with room to spare. So read the natural advance off the
    corpus instead: per (face, size, char) the most common advance wins, because most
    occurrences of any character are under no pressure at all."""
    seen = defaultdict(Counter)
    for adv, key in rows:
        seen[key][round(adv, 2)] += 1
    return {k: c.most_common(1)[0][0] for k, c in seen.items() if sum(c.values()) >= 3}


def main():
    import fitz  # noqa: F401
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    show = "--lines" in sys.argv
    pdfs = sorted(glob.glob(os.path.join(DOCS, "*_rt.pdf")))
    if args:
        pdfs = [p for p in pdfs if any(a in os.path.basename(p) for a in args)]
    rows = []
    skipped = Counter()
    for pdf in pdfs:
        docx = pdf[:-7] + ".docx"
        marl, marr, explicit = cellmar(docx)
        doc = os.path.basename(pdf)[:12]
        for inner, cs in lines_of(pdf, marl, marr):
            body = cs
            while body and not body[-1][0].strip():
                body = body[:-1]
            if len(body) < 2:
                continue
            if not all(fullwidth(c) for c, _, _, _ in body):
                skipped["has non-fullwidth"] += 1
                continue
            span = body if len(body) < len(cs) else body[:-1]
            if not span:
                continue
            end = cs[len(body)][1] if len(body) < len(cs) else body[-1][1]
            per = [(body[i + 1][1] - body[i][1] if i + 1 < len(body) else end - body[i][1],
                    (body[i][3], round(body[i][2], 2), body[i][0]))
                   for i in range(len(span))]
            rows.append(dict(doc=doc, inner=inner, measured=end - body[0][1],
                             whole=len(body) < len(cs), per=per,
                             text="".join(c for c, _, _, _ in body),
                             chars=body, em=body[0][2]))
    nat = natural_table([(a, k) for r in rows for a, k in r["per"]])
    keep = []
    for r in rows:
        if not all(k in nat for _, k in r["per"]):
            skipped["face/size seen < 3 times"] += 1
            continue
        r["natural"] = sum(nat[k] for _, k in r["per"])
        r["supplied"] = r["natural"] - r["measured"]
        r["demand"] = r["natural"] - r["inner"]
        r["give"] = [(k[2], nat[k] - a) for a, k in r["per"]]
        keep.append(r)
    rows = keep
    print(f"{len(pdfs)} exports, {len(rows)} cell lines usable "
          f"({skipped['has non-fullwidth']} skipped: not pure fullwidth, "
          f"{skipped['face/size seen < 3 times']}: face/size too rare to calibrate)")
    print(f"natural-advance table: {len(nat)} (face, size, char) entries")

    # supplied against demand -- the whole question
    bins = defaultdict(list)
    for r in rows:
        d = r["demand"]
        key = ("demand <= -3pt (room to spare)" if d <= -3 else
               "-3 .. 0pt (just fits)" if d <= 0 else
               "0 .. +1em" if d <= r["em"] else
               "+1em .. +2em" if d <= 2 * r["em"] else "> +2em")
        bins[key].append(r["supplied"])
    order = ["demand <= -3pt (room to spare)", "-3 .. 0pt (just fits)",
             "0 .. +1em", "+1em .. +2em", "> +2em"]
    print(f"\n{'demand':<32}{'n':>6}{'supplied median':>17}{'max':>9}{'>0.5pt':>9}")
    for k in order:
        v = sorted(bins.get(k, []))
        if not v:
            continue
        n = len(v)
        print(f"{k:<32}{n:>6}{v[n // 2]:>17.2f}{v[-1]:>9.2f}"
              f"{sum(1 for x in v if x > 0.5) / n:>8.0%}")

    comp = [r for r in rows if r["supplied"] > 0.5 and r["whole"]]
    print(f"\ncompressed lines (supplied > 0.5pt, last advance known): {len(comp)}")
    if comp:
        print(f"  demand>0 on {sum(1 for r in comp if r['demand'] > -0.5)}/{len(comp)}")
        by = defaultdict(list)
        for r in comp:
            for c, g in r["give"]:
                if g > 0.2:
                    by[klass(c)].append(g)
        print(f"  {'class':<12}{'chars that gave up':>20}{'median':>9}{'max':>9}")
        for k, v in sorted(by.items(), key=lambda kv: -len(kv[1])):
            v = sorted(v)
            print(f"  {k:<12}{len(v):>20}{v[len(v) // 2]:>9.2f}{v[-1]:>9.2f}")
        docs = Counter(r["doc"] for r in comp)
        print("  documents:", dict(docs.most_common(8)))
    if show:
        for r in sorted(comp, key=lambda r: -r["supplied"])[:15]:
            print(f"   {r['doc']} inner={r['inner']:6.2f} nat={r['natural']:6.2f} "
                  f"got={r['measured']:6.2f} demand={r['demand']:+5.2f} "
                  f"supplied={r['supplied']:5.2f}  {r['text']!r}")


if __name__ == "__main__":
    main()
