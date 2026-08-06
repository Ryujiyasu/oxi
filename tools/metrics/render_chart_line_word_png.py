# -*- coding: utf-8 -*-
"""Rasterize the Word line-chart PDF page 1 to a PNG for visual inspection."""
import sys, os
sys.stdout.reconfigure(encoding="utf-8")
import fitz

pdf = r"pipeline_data\pptx_probes\chart_line\chart_line.pdf"
out = r"scratchpad\chart_line_word.png"
os.makedirs("scratchpad", exist_ok=True)

doc = fitz.open(pdf)
page = doc[0]
# 150 dpi like the renderer (1pt = 2.0833px)
mat = fitz.Matrix(150 / 72, 150 / 72)
pix = page.get_pixmap(matrix=mat)
pix.save(out)
print("saved", out, pix.width, "x", pix.height)
