# -*- coding: utf-8 -*-
"""Re-dump ALL text from chart1.pdf and chart2.pdf with positions, to check
whether a legend ('Series 1' / 'Revenue' / 'Cost') is really rendered."""
import fitz

for name in ["chart1", "chart2"]:
    doc = fitz.open(rf"pipeline_data\pptx_probes\{name}\{name}.pdf")
    print(f"=== {name}.pdf ===")
    print("full text:", repr(doc[0].get_text("text")))
    for b in doc[0].get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                print(f"  '{s['text']}'  font={s['font']} size={s['size']:.2f} origin=({s['origin'][0]:.2f},{s['origin'][1]:.2f}) bbox={tuple(round(v,2) for v in s['bbox'])}")
    doc.close()
