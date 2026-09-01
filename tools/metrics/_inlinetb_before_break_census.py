# -*- coding: utf-8 -*-
"""Paragraphs that draw an inline object and THEN break the page.

    <w:p>
      <w:r>…<wp:inline>…<w:txbxContent>…</w:txbxContent>…</w:r>   <- drawn here
      <w:r><w:br w:type="page"/></w:r>                            <- then break
    </w:p>

Word draws the object on the CURRENT page and starts a new one after it.
Oxi lifts the drawing out of the run list into the text-box list, so the only
thing left in the paragraph is the `\\x0C` the parser writes for the break --
the page closes FIRST and the box lands on the NEXT page with it.

legal__02f84965dccfe4db p4: Word puts the 103.35pt box at y=676.2 and ends the
page; Oxi ends the page at y=545.5 (225pt of unused column) and draws the box
at the top of p5. Deleting that one `<w:br>` takes Oxi from 11 pages to 10.

Counts how many corpus paragraphs have this shape, so a fix can be scoped.
Nesting-safe: `<w:p>` is matched by walking tags, not by a non-greedy regex
(which breaks on the `<w:p>` inside `<w:txbxContent>` -- that mistake hid 2 of
this document's 7 page breaks).

Usage: _inlinetb_before_break_census.py <dir> [<dir> ...]
"""
import os
import re
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

TAG = re.compile(r"<(/?)w:p(?:\s[^>]*)?(/?)>|<w:drawing\b|<w:pict\b|"
                 r'<w:br w:type="page"\s*/>')


def top_level_paragraphs(xml):
    """Yield (start, end) of every TOP-LEVEL w:p, tracking nesting depth."""
    depth = 0
    start = None
    for m in re.finditer(r"<w:p(?:\s[^>]*)?>|</w:p>|<w:p(?:\s[^>]*)?/>", xml):
        t = m.group(0)
        if t.endswith("/>") and not t.startswith("</"):
            continue  # empty <w:p/>
        if not t.startswith("</"):
            if depth == 0:
                start = m.start()
            depth += 1
        else:
            depth -= 1
            if depth == 0 and start is not None:
                yield start, m.end()
                start = None


def scan(path):
    try:
        xml = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8", "replace")
    except Exception:
        return 0, 0
    hits = 0
    total_br = xml.count('<w:br w:type="page"/>')
    for s, e in top_level_paragraphs(xml):
        seg = xml[s:e]
        bi = seg.find('<w:br w:type="page"/>')
        if bi < 0:
            continue
        # a drawing (inline or anchored) that starts BEFORE the break
        di = min([i for i in (seg.find("<w:drawing"), seg.find("<w:pict")) if i >= 0]
                 or [-1])
        if 0 <= di < bi:
            hits += 1
    return hits, total_br


targets = []
for a in sys.argv[1:]:
    if os.path.isdir(a):
        for dp, _d, fs in os.walk(a):
            targets += [os.path.join(dp, f) for f in fs
                        if f.endswith(".docx") and not f.startswith("~$")]
    else:
        targets.append(a)

ndoc = 0
npara = 0
for p in sorted(targets):
    h, _t = scan(p)
    if h:
        ndoc += 1
        npara += h
        print("  %-44s %d paragraph(s)" % (os.path.basename(p)[:44], h))
print("\n%d paragraphs in %d of %d docs draw an inline object then break the page"
      % (npara, ndoc, len(targets)))
