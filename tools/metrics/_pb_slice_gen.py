# -*- coding: utf-8 -*-
"""Cut a faithful slice of a real document: N body paragraphs, every part kept.

A hand-written repro of `legal__001a2c7f07cd358f`'s shape does NOT reproduce its
behaviour (Word advances 23.04pt between paragraphs that declare no spacing at
all; the repro advances 11.40), so stop guessing which attribute matters and
narrow the real file instead: keep styles.xml, numbering.xml, settings.xml,
theme and fonts verbatim, and keep only a window of body paragraphs.

    python tools/metrics/_pb_slice_gen.py <docx> <first_para> <count> [out.docx]
"""
import os, re, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# Self-closing <w:p .../> is a paragraph too. Matching only the paired form
# silently drops them (36 of 1001 in legal__001a2c7f07cd358f) and shifts
# every index after the first one.
PARA = re.compile(r"<w:p(?: [^>]*?)?/>|<w:p(?: [^>]*)?>.*?</w:p>", re.S)


def slice_doc(src, first, count, out):
    zin = zipfile.ZipFile(src)
    items = [(it, zin.read(it.filename)) for it in zin.infolist()]
    doc = next(d for it, d in items
               if it.filename == "word/document.xml").decode("utf-8")
    body_open = doc.index("<w:body>") + len("<w:body>")
    body_close = doc.index("</w:body>")
    body = doc[body_open:body_close]
    sect = ""
    m = re.search(r"<w:sectPr(?: [^>]*)?>.*?</w:sectPr>\s*$", body, re.S)
    if m:
        sect = m.group(0)
    paras = list(PARA.finditer(body))
    keep = "".join(p.group(0) for p in paras[first:first + count])
    new_doc = doc[:body_open] + keep + sect + doc[body_close:]
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for it, data in items:
            z.writestr(it, new_doc.encode("utf-8")
                       if it.filename == "word/document.xml" else data)
    kept = [re.sub(r"<[^>]+>", "", p.group(0))[:34] for p in paras[first:first + count]]
    print("sliced %d paragraphs -> %s" % (len(kept), out))
    for k, t in enumerate(kept):
        print("   %3d %r" % (first + k, t))


if __name__ == "__main__":
    src, first, count = sys.argv[1], int(sys.argv[2]), int(sys.argv[3])
    out = sys.argv[4] if len(sys.argv) > 4 else r"C:\tmp\slice.docx"
    slice_doc(src, first, count, out)
