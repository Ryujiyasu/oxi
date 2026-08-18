# -*- coding: utf-8 -*-
"""When Word COULD squeeze one more character into a cell line, does it?

Everything about the capacity is measured: the budget, the two 約物 types, the
CJK/Latin gap, the ideographic space, and the trigger (compressPunctuation, and
either justification or a legacy compatibility mode). What is not measured is how
much of that capacity Word actually spends — and it is not "all of it". 34140b and
a47e squeeze; d77a58 declines on lines that carry the same trigger and the same
marks, and ends up exactly between Oxi's two arms, one character either way.

So take the decision itself as the observation. For every line inside a cell:

    slack  = the cell's content width - what Word drew
    need   = the next line's first character, minus that slack
    pool   = half an em if the line carries a closing mark, plus half an em for
             each opening bracket (the derived model)

A line where `need <= pool` is one Word COULD have squeezed and did not; a line
already over its natural width is one Word DID squeeze. Printing the two
populations side by side is what a discriminator has to separate.

    python _cw_decide.py 34140b a47e d77a58
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

CLOSING = "）」』〕】》〉｝］、。，．・"
OPENING = "（「『〔【《〈｛［"


def cellmar(docx):
    for part in ("word/document.xml", "word/styles.xml"):
        try:
            x = zipfile.ZipFile(docx).read(part).decode("utf-8")
        except Exception:
            continue
        m = re.search(r"<w:tblCellMar>(.*?)</w:tblCellMar>", x, re.S)
        if m:
            def side(tag):
                mm = re.search(r'<w:%s w:w="(-?\d+)"' % tag, m.group(1))
                return int(mm.group(1)) / 20.0 if mm else 5.4
            return side("left"), side("right")
    return 5.4, 5.4


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


def fullwidth(c):
    o = ord(c)
    return (0x3000 <= o <= 0x303F or 0x3040 <= o <= 0x30FF or 0x4E00 <= o <= 0x9FFF
            or 0x3400 <= o <= 0x4DBF or 0xFF01 <= o <= 0xFF60)


def span_naturals(obs):
    """The natural advance of each SPAN, from its own fullwidth advances.

    A span is one run of one formatting, so its median fullwidth advance already
    has that run's `w:spacing` in it. Spans too short to have a median fall back to
    the document-wide table."""
    per = defaultdict(list)
    for adv, key, sid in obs:
        if fullwidth(key[2]):
            per[sid].append(adv)
    out = {}
    for sid, v in per.items():
        if len(v) >= 5:
            v.sort()
            out[sid] = v[len(v) // 2]
    return out


def natural_table(obs):
    """The advance a fullwidth character takes when nothing squeezes it.

    Keying the mode per CHARACTER needs a sample per character, and rare ones came
    back with whatever their one compressed occurrence was -- that is what produced
    "squeezed" lines with a pool of zero (公表事項, a URL, ***,***,***). Every
    fullwidth glyph in these faces is one em, so the mode over ALL of them at a
    given (face, size) is the natural advance for all of them, spacing included,
    and it has the whole document behind it."""
    per_face = defaultdict(Counter)
    for adv, (face, size, ch), _sid in obs:
        if fullwidth(ch):
            per_face[(face, size)][round(adv, 2)] += 1
    face_nat = {k: c.most_common(1)[0][0] for k, c in per_face.items()
                if sum(c.values()) >= 8}
    # halfwidth and Latin still need their own mode; they are not one em
    seen = defaultdict(Counter)
    for adv, key, _sid in obs:
        if not fullwidth(key[2]):
            seen[key][round(adv, 2)] += 1
    out = {k: c.most_common(1)[0][0] for k, c in seen.items() if sum(c.values()) >= 3}
    for (face, size), v in face_nat.items():
        out[("*FULL*", face, size)] = v
    return out


def nat_of(nat, key, spans=None, sid=None):
    face, size, ch = key
    if fullwidth(ch):
        if spans is not None and sid in spans:
            return spans[sid]
        return nat.get(("*FULL*", face, size))
    return nat.get(key)


def read(pdf, marl, marr):
    import fitz
    raw, pages = [], []
    for pg in fitz.open(pdf):
        rules = rules_of(pg)
        got = []
        for b in pg.get_text("rawdict")["blocks"]:
            for ln in b.get("lines", []):
                # ★Tag each character with WHICH SPAN it came from. A span is a
                # formatting run, so its own advances carry that run's w:spacing --
                # and a document-wide (face, size) mode does not. Reading the
                # natural advance per (face, size) made every run with a stronger
                # negative spacing look "compressed": the lines it reported as
                # squeezed were only 73% full, which no demand-driven compression
                # can produce.
                cs = [(c["c"], c["bbox"][0], round(sp["size"], 2), sp["font"], si)
                      for si, sp in enumerate(ln["spans"]) for c in sp["chars"]]
                sid = (round(ln["bbox"][1], 1),)
                if len(cs) < 2:
                    continue
                y0, y1 = ln["bbox"][1], ln["bbox"][3]
                near = sorted({x for x, a, b_ in rules if a <= y0 + 1 and b_ >= y1 - 1})
                left = [x for x in near if x < cs[0][1]]
                right = [x for x in near if x > cs[-1][1]]
                if not left or not right:
                    continue
                if any(cs[0][1] < x < cs[-1][1] for x in near):
                    continue                       # welded two cells of one row
                got.append((round(y0, 1), left[-1] + marl, right[0] - marr,
                            [(c, x, sz, f, sid + (si,)) for c, x, sz, f, si in cs]))
        got.sort()
        pages.append(got)
        for _, _, _, cs in got:
            for i in range(len(cs) - 1):
                if cs[i][4] != cs[i + 1][4]:
                    continue                       # advance across a span edge
                raw.append((cs[i + 1][1] - cs[i][1],
                            (cs[i][3], cs[i][2], cs[i][0]), cs[i][4]))
    return pages, raw


def main():
    import fitz  # noqa: F401
    for pref in sys.argv[1:]:
        docx = sorted(glob.glob(os.path.join(DOCS, pref + "*.docx")))
        pdf = [p for p in glob.glob(os.path.join(DOCS, pref + "*_rt.pdf"))]
        if not docx or not pdf:
            print("%-10s no cached export" % pref)
            continue
        marl, marr = cellmar(docx[0])
        pages, raw = read(pdf[0], marl, marr)
        nat = natural_table(raw)
        squeezed, declined, blocked = [], [], []
        for got in pages:
            for gi, (y, cl, cr, cs) in enumerate(got):
                body = [c for c in cs if c[0].strip()]
                if len(body) < 3:
                    continue
                keys = [(c[3], c[2], c[0]) for c in body]
                vals = [nat_of(nat, k) for k in keys]
                if any(v is None for v in vals):
                    continue
                drawn = body[-1][1] - body[0][1] + vals[-1]
                natural = sum(vals)
                inner = cr - cl
                slack = inner - drawn
                em = body[0][2]
                text = "".join(c[0] for c in body)
                pool = (em * 0.5 if any(c in CLOSING for c in text) else 0.0) \
                    + em * 0.5 * sum(1 for c in text if c in OPENING)
                if natural - drawn > 0.3:
                    squeezed.append((natural - drawn, pool, text))
                    continue
                nxt = got[gi + 1][3] if gi + 1 < len(got) else None
                if not nxt:
                    continue
                nc = nxt[0]
                nk = (nc[3], nc[2], nc[0])
                nv = nat_of(nat, nk)
                if nv is None:
                    continue
                need = nv - slack
                if 0 < need <= pool:
                    declined.append((need, pool, text))
                elif need > pool:
                    blocked.append((need, pool, text))
        print("\n=== %s  (cellMar %.2f/%.2f)" % (pref, marl, marr))
        print("    squeezed  %4d lines   median given %.2fpt" %
              (len(squeezed), sorted(s for s, _, _ in squeezed)[len(squeezed) // 2]
               if squeezed else 0))
        print("    DECLINED  %4d lines   (need <= pool, Word did not take it)" %
              len(declined))
        print("    blocked   %4d lines   (need > pool, could not have)" % len(blocked))
        for tag, rows in (("squeezed", squeezed), ("declined", declined)):
            if not rows:
                continue
            r = sorted(rows)[:2] + sorted(rows)[-1:]
            for v, p, t in r:
                print("      %-9s need/given %5.2f  pool %5.2f  %s" % (tag, v, p, t[:38]))


if __name__ == "__main__":
    main()
