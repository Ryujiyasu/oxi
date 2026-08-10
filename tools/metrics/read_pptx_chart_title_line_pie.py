# -*- coding: utf-8 -*-
"""Read chart_title_line / chart_title_line2 / chart_title_pie /
chart_title_pie2 Word PDFs: dump text spans (esp. the explicit title
'Quarterly Revenue') and vector drawings to derive the title baseline, the
plot-top shift (line) and the circle-top shift (pie) vs the auto-title
geometry the renderer already handles."""
import sys
sys.stdout.reconfigure(encoding="utf-8")
import fitz


def read(base_name):
    doc = fitz.open(rf"pipeline_data\pptx_probes\{base_name}\{base_name}.pdf")
    page = doc[0]
    print(f"=== {base_name} page.rect = {page.rect} ===")
    print("\n=== TEXT SPANS (dict, origin=baseline) ===")
    d = page.get_text("dict")
    for block in d["blocks"]:
        if "lines" not in block:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                t = span["text"].strip()
                if t:
                    o = span.get("origin")
                    o = (o[0], o[1]) if o else (float("nan"), float("nan"))
                    b = span["bbox"]
                    print(
                        f"'{t}' origin=({o[0]:.2f},{o[1]:.2f}) "
                        f"bbox=({b[0]:.2f},{b[1]:.2f},{b[2]:.2f},{b[3]:.2f}) "
                        f"size={span['size']:.2f} "
                        f"color=#{span['color']:06x} font={span['font']}"
                    )
    print("\n=== VECTOR DRAWINGS ===")
    for dr in page.get_drawings():
        r = dr["rect"]
        print(
            f"rect=({r.x0:.1f},{r.y0:.1f},{r.x1:.1f},{r.y1:.1f}) "
            f"w={r.width:.1f} h={r.height:.1f} fill={dr['fill']} "
            f"stroke={dr['color']} n_items={len(dr['items'])}"
        )
        for it in dr["items"]:
            if it[0] == "re":
                rr = it[1]
                print(f"    ('re', Rect({rr.x0:.2f}, {rr.y0:.2f}, {rr.x1:.2f}, {rr.y1:.2f}))")
            elif it[0] == "l":
                p0, p1 = it[1], it[2]
                print(f"    ('l', Point({p0.x:.2f},{p0.y:.2f}), Point({p1.x:.2f},{p1.y:.2f}))")
    doc.close()


for name in ("chart_title_line", "chart_title_line2", "chart_title_pie", "chart_title_pie2"):
    read(name)
