# -*- coding: utf-8 -*-
"""RENDER-TRUTH for correspondence__000407cd: export to PDF via Word, dump
line bboxes AND image bboxes so the inline OLE + inline picture row can be
compared against Oxi's stacked lines."""
import os
import sys
import tempfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else
                       "pipeline_data/docx_corpus/en/correspondence/000407cd6c79442c.docx")
PDF = os.path.join(tempfile.gettempdir(), os.path.basename(DOCX) + ".truth.pdf")

if not os.path.exists(PDF) or "--reexport" in sys.argv:
    import win32com.client as win32
    w = win32.DispatchEx("Word.Application")
    w.Visible = False
    try:
        d = w.Documents.Open(DOCX, ReadOnly=True)
        print("Word pages:", d.ComputeStatistics(2))
        d.ExportAsFixedFormat(PDF, 17)
        d.Close(False)
    finally:
        w.Quit()
    print("exported", PDF)

import fitz  # noqa: E402

doc = fitz.open(PDF)
print("pdf pages:", doc.page_count)
for pi in range(doc.page_count):
    pg = doc[pi]
    print("=== page", pi, "size", pg.rect)
    rows = []
    for blk in pg.get_text("dict")["blocks"]:
        if blk.get("type", 0) != 0:
            continue
        for ln in blk.get("lines", []):
            txt = "".join(s["text"] for s in ln["spans"])
            if not txt.strip():
                txt = "<ws:%d>" % len(txt)
            y0 = min(s["bbox"][1] for s in ln["spans"])
            x0 = min(s["bbox"][0] for s in ln["spans"])
            x1 = max(s["bbox"][2] for s in ln["spans"])
            sz = max(s["size"] for s in ln["spans"])
            rows.append((y0, "T", "x %7.2f..%7.2f sz %4.1f %r" % (x0, x1, sz, txt[:44])))
    for im in pg.get_image_info():
        b = im["bbox"]
        rows.append((b[1], "I", "x %7.2f..%7.2f  h %6.2f w %6.2f" % (b[0], b[2], b[3] - b[1], b[2] - b[0])))
    for d in pg.get_drawings():
        b = d["rect"]
        if b.height < 0.4 and b.width < 0.4:
            continue
        rows.append((b.y0, "D", "x %7.2f..%7.2f  h %6.2f" % (b.x0, b.x1, b.height)))
    rows.sort(key=lambda r: (r[0], r[1]))
    prev = None
    for y, k, s in rows:
        pitch = ("%6.2f" % (y - prev)) if prev is not None else "     -"
        print("  %-8.2f %s %s  %s" % (y, k, pitch, s))
        prev = y
