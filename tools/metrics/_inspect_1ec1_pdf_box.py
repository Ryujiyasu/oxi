# -*- coding: utf-8 -*-
"""Inspect each □ in 1ec1 PDF — y position, surrounding text, page."""
import sys, fitz
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
PDF = "pipeline_data/1ec1_actual_pdf_measure/1ec1.pdf"

d = fitz.open(PDF)
print(f"Pages: {d.page_count}")
for pi in range(d.page_count):
    page = d[pi]
    insts = page.search_for("□")
    print(f"\nPage {pi+1}: {len(insts)} □ instances, page size {page.rect}")
    for i, inst in enumerate(insts):
        # Get text in surrounding rect
        ctx_rect = fitz.Rect(inst.x0 - 5, inst.y0, inst.x1 + 200, inst.y1)
        text = page.get_text("text", clip=ctx_rect).replace('\n', ' ')[:50]
        print(f"  □#{i+1}: x=[{inst.x0:.2f}, {inst.x1:.2f}] y=[{inst.y0:.2f}, {inst.y1:.2f}] | text: {text!r}")
