# -*- coding: utf-8 -*-
"""Read pie_label_probe.pdf label spans via fitz (per-slide)."""
import os, sys
sys.stdout.reconfigure(encoding="utf-8")
import fitz

name = "pie_label_probe"
pdf = os.path.join(os.path.abspath(r"pipeline_data\pptx_probes"), name, name + ".pdf")

doc = fitz.open(pdf)
for pi, page in enumerate(doc):
    d = page.get_text("dict")
    spans = []
    for blk in d["blocks"]:
        for line in blk.get("lines", []):
            for sp in line["spans"]:
                t = sp["text"].strip()
                if not t:
                    continue
                x0, y0, x1, y1 = sp["bbox"]
                spans.append((t, sp["origin"], (round(x0,2), round(y0,2), round(x1,2), round(y1,2)),
                             round(sp["size"],2)))
    print(f"=== page {pi} ===")
    for t, org, bb, sz in spans:
        print(f"  '{t}' origin=({org[0]:.2f},{org[1]:.2f}) bbox={bb} size={sz}")
doc.close()
