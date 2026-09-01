# -*- coding: utf-8 -*-
"""How many corpus floats did the page-edge clamp move?  (S1268)

Reads sectPr + every wp:anchor straight out of the .docx and reports the
anchors whose RAW resolved rect leaves the page, i.e. exactly the ones
`resolve_textbox_position` / `resolve_floating_image_position` used to slide.

The column reference is approximated by the left margin (what Oxi resolves for
a column-1 anchor); an anchor flowing in column 2 overflows at least as much,
so the count is a LOWER bound.

Usage: _pb_floatclamp_census.py <dir> [<dir> ...]
"""
import os
import re
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
EMU = 12700.0
TW = 20.0


def sect(doc):
    m = re.search(r'<w:pgSz w:w="(\d+)" w:h="(\d+)"([^/]*)/>', doc)
    if not m:
        return None
    w, h = int(m.group(1)) / TW, int(m.group(2)) / TW
    if "landscape" in m.group(3):
        w, h = max(w, h), min(w, h)
    mm = re.search(r'<w:pgMar ([^/]*)/>', doc)
    left = right = 72.0
    if mm:
        a = dict(re.findall(r'w:(\w+)="(-?\d+)"', mm.group(1)))
        left = int(a.get("left", 1440)) / TW
        right = int(a.get("right", 1440)) / TW
    return w, h, left, right


def anchors(doc):
    for m in re.finditer(r"<wp:anchor\b.*?</wp:anchor>", doc, re.S):
        s = m.group(0)
        h = re.search(r'<wp:positionH relativeFrom="([^"]+)">(.*?)</wp:positionH>', s, re.S)
        v = re.search(r'<wp:positionV relativeFrom="([^"]+)">(.*?)</wp:positionV>', s, re.S)
        e = re.search(r'<wp:extent cx="(\d+)" cy="(\d+)"', s)
        if not (h and v and e):
            continue

        def off(block):
            o = re.search(r"<wp:posOffset>(-?\d+)</wp:posOffset>", block)
            return None if o is None else int(o.group(1)) / EMU

        yield dict(
            hrel=h.group(1), hoff=off(h.group(2)),
            halign=(re.search(r"<wp:align>(\w+)</wp:align>", h.group(2)) or [None, None])[1]
            if re.search(r"<wp:align>", h.group(2)) else None,
            vrel=v.group(1), voff=off(v.group(2)),
            w=int(e.group(1)) / EMU, h=int(e.group(2)) / EMU,
            is_tb="<w:txbxContent" in s,
        )


def scan(root):
    docs = 0
    hit_docs = 0
    n_anch = 0
    n_x = 0
    n_y = 0
    worst = []
    for dirpath, _dirs, files in os.walk(root):
        for fn in files:
            if not fn.endswith(".docx") or fn.startswith("~$"):
                continue
            p = os.path.join(dirpath, fn)
            try:
                doc = zipfile.ZipFile(p).read("word/document.xml").decode("utf-8", "replace")
            except Exception:
                continue
            g = sect(doc)
            if not g:
                continue
            pw, ph, ml, _mr = g
            docs += 1
            hits = 0
            for a in anchors(doc):
                n_anch += 1
                if a["hoff"] is None or a["vrel"] not in ("page", "margin", "paragraph", "line"):
                    pass
                if a["hoff"] is not None:
                    ref = 0.0 if a["hrel"] == "page" else ml
                    x = ref + a["hoff"]
                    if x + a["w"] > pw:
                        n_x += 1
                        hits += 1
                        worst.append((x + a["w"] - pw, fn, "x", a["w"], a["is_tb"]))
                if a["voff"] is not None and a["vrel"] == "page":
                    if a["voff"] + a["h"] > ph:
                        n_y += 1
                        hits += 1
                        worst.append((a["voff"] + a["h"] - ph, fn, "y", a["h"], a["is_tb"]))
            if hits:
                hit_docs += 1
    print("%-46s docs=%4d  anchors=%5d  x_overflow=%4d  y_overflow=%3d  docs_hit=%3d (%.1f%%)"
          % (root, docs, n_anch, n_x, n_y, hit_docs, 100.0 * hit_docs / max(docs, 1)))
    worst.sort(reverse=True)
    for over, fn, ax, size, is_tb in worst[:8]:
        print("      %6.1fpt past the %s edge  %-42s size=%.1f %s"
              % (over, ax, fn[:42], size, "textbox" if is_tb else "image/shape"))


for d in sys.argv[1:]:
    scan(d)
