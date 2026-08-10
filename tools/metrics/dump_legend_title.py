# -*- coding: utf-8 -*-
"""Measure the LEGEND origin from the Word line-chart PDFs that carry an
explicit <c:title> (chart_title_line / chart_title_line2).

Dumps all vector drawings (items) + text spans (rawdict), focusing on the
right-hand legend region so we can read Word's legend row y / label origins
and compare against the renderer's frame-relative legend_y0.
"""
import sys, os
sys.stdout.reconfigure(encoding="utf-8")
import fitz

def dump(pdf):
    print("=" * 70)
    print("PDF:", pdf)
    doc = fitz.open(pdf)
    page = doc[0]
    print("page size:", round(page.rect.width, 2), "x", round(page.rect.height, 2))

    print("--- drawings (%d) ---" % len(page.get_drawings()))
    for i, d in enumerate(page.get_drawings()):
        r = d["rect"]
        fill = tuple(round(c, 3) for c in d["fill"]) if d.get("fill") is not None else None
        stroke = tuple(round(c, 3) for c in d["color"]) if d.get("color") is not None else None
        kinds = {}
        for it in d["items"]:
            kinds[it[0]] = kinds.get(it[0], 0) + 1
        print("[%d] rect=(%.2f,%.2f,%.2f,%.2f) fill=%s stroke=%s w=%s kinds=%s"
              % (i, r[0], r[1], r[2], r[3], fill, stroke, d.get("width"), kinds))
        for it in d["items"]:
            if it[0] in ("l", "c", "re"):
                print("   ", it)

    print("--- text spans (rawdict) ---")
    raw = page.get_text("rawdict")
    for b in raw["blocks"]:
        if b["type"] != 0:
            continue
        for ln in b.get("lines", []):
            for sp in ln["spans"]:
                txt = "".join(ch["c"] for ch in sp["chars"])
                print("   origin=(%.2f,%.2f) bbox=(%.2f,%.2f,%.2f,%.2f) "
                      "font=%s size=%.2f text=%r"
                      % (sp["origin"][0], sp["origin"][1],
                         sp["bbox"][0], sp["bbox"][1], sp["bbox"][2], sp["bbox"][3],
                         sp["font"], sp["size"], txt))
    doc.close()

dump(r"pipeline_data\pptx_probes\chart_title_line\chart_title_line.pdf")
dump(r"pipeline_data\pptx_probes\chart_title_line2\chart_title_line2.pdf")
