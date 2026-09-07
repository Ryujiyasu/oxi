# -*- coding: utf-8 -*-
"""Where does Word put an underline?  The document is exported through Word to
PDF; every thin horizontal rule in the PDF is paired with the text line whose
baseline sits just above it, and the offset and thickness are printed as
fractions of the run's font size -- the two numbers a renderer needs.

Usage: python tools/metrics/_pb_underline_geom.py <docx> [max-pages]
"""
import glob
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import fitz  # noqa: E402

docx = sys.argv[1]
if not os.path.exists(docx):
    docx = glob.glob("tools/golden-test/documents/docx/%s*.docx" % docx)[0]
maxp = int(sys.argv[2]) if len(sys.argv) > 2 else 3
pdf = os.path.join("pipeline_data", "_ul_" + os.path.splitext(os.path.basename(docx))[0] + ".pdf")
if not os.path.exists(pdf) or os.path.getmtime(pdf) < os.path.getmtime(docx):
    import win32com.client as w

    app = w.Dispatch("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        d.ExportAsFixedFormat(os.path.abspath(pdf), 17)
        d.Close(False)
    finally:
        app.Quit()
print("pdf:", pdf)
doc = fitz.open(pdf)
rows = []
for pno in range(min(maxp, len(doc))):
    page = doc[pno]
    spans = []
    for b in page.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                if s["text"].strip():
                    spans.append(s)
    rules = []
    for dr in page.get_drawings():
        r = dr["rect"]
        if r.height <= 2.5 and r.width >= 6 and r.height >= 0.05:
            rules.append((r, dr.get("fill") or dr.get("color")))
    for r, col in rules:
        # the span whose baseline is just above this rule and whose x range overlaps it
        best = None
        for s in spans:
            base = s["origin"][1]
            if base > r.y0 + 0.6 or base < r.y0 - 6.0:
                continue
            ox = min(s["bbox"][2], r.x1) - max(s["bbox"][0], r.x0)
            if ox <= 0:
                continue
            d = r.y0 - base
            if best is None or d < best[0]:
                best = (d, s, ox)
        if best is None:
            continue
        d, s, ox = best
        # a table border sits far below the baseline and usually runs wider than
        # the text: keep only rules that hug the run they underline
        if d > 0.30 * s["size"] or ox < 0.6 * (r.x1 - r.x0):
            continue
        rows.append((pno + 1, s["font"], round(s["size"], 2), round(d, 3), round(r.height, 3),
                     round(d / s["size"], 4), round(r.height / s["size"], 4), s["text"][:8]))
print("page  font                       size   dy    thick   dy/size  th/size  text")
for r in rows[:60]:
    print("  %-4d %-26s %5.1f %6.3f %6.3f  %7.4f %7.4f  %s" % r)
if rows:
    import statistics

    by_font = {}
    for r in rows:
        by_font.setdefault((r[1], r[2]), []).append((r[5], r[6]))
    print("\nper font/size medians (dy/size, thickness/size, n):")
    for (f, sz), v in sorted(by_font.items()):
        print("  %-26s %5.1fpt  %.4f  %.4f  n=%d" % (f, sz, statistics.median(x[0] for x in v),
                                                     statistics.median(x[1] for x in v), len(v)))
