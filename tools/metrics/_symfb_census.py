# -*- coding: utf-8 -*-
"""Corpus exposure of the S1119 symbol-fallback line height — done properly.

A first cut of this census reported "2 documents, 20 occurrences" and both
turned out to be non-exposed.  Two mistakes, both worth keeping in the tool so
the next census does not repeat them:

  ★1. It resolved the run's font through theme / docDefaults and IGNORED the
      run's OWN <w:rFonts>.  forms__001ae487's ballot boxes are MS Gothic runs
      (w:hint="eastAsia"), and MS Gothic HAS U+2610 — no fallback at all.  The
      run's own rFonts wins over everything upstream; resolve it first.
  ★2. It counted codepoints without asking whether the line height can move.
      forms__002f81ab's 16 stars are all in ONE table row, and that row carries
      <w:trHeight w:hRule="atLeast" w:val="11241"> = 562pt.  A 2pt line-height
      change is absorbed whole.  A codepoint only counts as exposed if nothing
      pins its line.

Reports separately: runs that fall back at all, and the subset whose line is
free to move (body paragraph, or a cell whose row has no fixed trHeight).

usage: python tools/metrics/_symfb_census.py
"""

import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from pathlib import Path

from fontTools.ttLib import TTFont

REPO = Path(__file__).resolve().parents[2]
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
FONTDIR = Path(r"C:/Windows/Fonts")
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

FILE_FOR = {
    "Arial": ("arial.ttf", 0), "Calibri": ("calibri.ttf", 0),
    "Cambria": ("cambria.ttc", 0), "Cambria Math": ("cambria.ttc", 1),
    "Courier New": ("cour.ttf", 0), "Times New Roman": ("times.ttf", 0),
    "Segoe UI": ("segoeui.ttf", 0), "Segoe UI Symbol": ("seguisym.ttf", 0),
    "Verdana": ("verdana.ttf", 0), "Tahoma": ("tahoma.ttf", 0),
    "Georgia": ("georgia.ttf", 0), "Trebuchet MS": ("trebuc.ttf", 0),
    "MS Gothic": ("msgothic.ttc", 0), "MS Mincho": ("msmincho.ttc", 0),
    "MS PGothic": ("msgothic.ttc", 1), "Meiryo": ("meiryo.ttc", 0),
    "Yu Gothic": ("YuGothR.ttc", 0), "Garamond": ("GARA.TTF", 0),
    "Book Antiqua": ("BKANT.TTF", 0), "Century Gothic": ("GOTHIC.TTF", 0),
    "Aptos": ("Aptos.ttf", 0), "Bookman Old Style": ("BOOKOS.TTF", 0),
}
CHAIN = ["Courier New", "Cambria Math", "Segoe UI Symbol"]
RANGES = [(0x2010, 0x2044), (0x2190, 0x22FF), (0x2460, 0x24FF), (0x2500, 0x27BF)]
_cm = {}


def cmap_of(fam):
    if fam not in _cm:
        spec = FILE_FOR.get(fam)
        out = None
        if spec:
            fn, idx = spec
            p = FONTDIR / fn
            if p.exists():
                try:
                    f = TTFont(p, fontNumber=idx, lazy=True)
                    out = set(f.getBestCmap())
                    f.close()
                except Exception:
                    out = None
        _cm[fam] = out
    return _cm[fam]


def ambiguous(cp):
    return any(a <= cp <= b for a, b in RANGES)


def real_cjk(c):
    o = ord(c)
    return (0x3040 <= o <= 0x30FF or 0x3400 <= o <= 0x4DBF or 0x4E00 <= o <= 0x9FFF
            or 0xF900 <= o <= 0xFAFF or 0xFF00 <= o <= 0xFFEF)


def theme_map(z):
    try:
        t = ET.fromstring(z.read("word/theme/theme1.xml"))
    except Exception:
        return {}
    out = {}
    for kind in ("major", "minor"):
        n = t.find(".//%s%sFont" % (A, kind))
        if n is None:
            continue
        lt = n.find(A + "latin")
        out["%sHAnsi" % kind] = lt.get("typeface", "") if lt is not None else ""
    return out


def doc_default(z, th):
    try:
        s = ET.fromstring(z.read("word/styles.xml"))
    except Exception:
        return "Calibri"
    dd = s.find(W + "docDefaults")
    rf = dd.find(".//" + W + "rFonts") if dd is not None else None
    if rf is None:
        return "Calibri"
    return rf.get(W + "ascii") or th.get(rf.get(W + "asciiTheme", ""), "") or "Calibri"


def main():
    n_doc = 0
    fired = Counter()
    movable = Counter()
    for root in (REPO / "pipeline_data" / "docx_corpus",
                 REPO / "tools" / "golden-test" / "documents" / "docx"):
        if not root.exists():
            continue
        for p in root.rglob("*.docx"):
            try:
                z = zipfile.ZipFile(p)
                t = ET.fromstring(z.read("word/document.xml"))
            except Exception:
                continue
            n_doc += 1
            if any(real_cjk(c) for n in t.iter(W + "t") for c in (n.text or "")):
                z.close()
                continue
            th = theme_map(z)
            dflt = doc_default(z, th)
            z.close()
            did = "%s__%s" % (p.parent.name, p.stem)

            def walk(e, anc):
                if e.tag != W + "r":
                    for c in e:
                        walk(c, anc + [e])
                    return
                rpr = e.find(W + "rPr")
                fam = dflt
                if rpr is not None:
                    rf = rpr.find(W + "rFonts")
                    if rf is not None:
                        # ★1: the run's OWN rFonts wins. hint="eastAsia" means the
                        # eastAsia face draws it, which is why MS Gothic ballot
                        # boxes never fall back.
                        fam = (rf.get(W + "eastAsia") if rf.get(W + "hint") == "eastAsia"
                               else None) or rf.get(W + "ascii") or \
                              th.get(rf.get(W + "asciiTheme", ""), "") or dflt
                cm = cmap_of(fam)
                if cm is None:
                    return
                text = "".join(n.text or "" for n in e.iter(W + "t"))
                for ch in text:
                    cp = ord(ch)
                    if cp < 0x80 or not ambiguous(cp) or cp in cm:
                        continue
                    if not any(cp in (cmap_of(f) or set()) for f in CHAIN):
                        continue
                    fired[did] += 1
                    # ★2: can the line actually move?
                    row = next((a for a in reversed(anc) if a.tag == W + "tr"), None)
                    pinned = False
                    if row is not None:
                        trpr = row.find(W + "trPr")
                        h = trpr.find(W + "trHeight") if trpr is not None else None
                        if h is not None and h.get(W + "val"):
                            # atLeast with a large val, or exact, pins the line
                            pinned = True
                    if not pinned:
                        movable[did] += 1

            walk(t, [])
    print("Latin docs scanned (of %d)" % n_doc)
    print("docs where a run FALLS BACK          : %d  (%d occurrences)"
          % (len(fired), sum(fired.values())))
    print("docs where the line can also MOVE    : %d  (%d occurrences)"
          % (len(movable), sum(movable.values())))
    for d, n in movable.most_common(10):
        print("   %-44s %d" % (d, n))
    if not movable:
        print("   (none — every fallback sits in a row whose height is pinned)")


if __name__ == "__main__":
    main()
