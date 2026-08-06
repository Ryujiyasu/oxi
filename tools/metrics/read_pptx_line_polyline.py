# -*- coding: utf-8 -*-
"""Dump ALL drawings for chart_line2 / chart_line3 to confirm the polyline
color (Word draws the connecting polyline in the series BORDER color, like
every other stroke). No filtering, so we see the diagonal line segments."""
import sys, os
sys.stdout.reconfigure(encoding="utf-8")
import fitz

for name in ("chart_line2", "chart_line3"):
    pdf = os.path.abspath(os.path.join(r"pipeline_data\pptx_probes", name, name + ".pdf"))
    doc = fitz.open(pdf)
    page = doc[0]
    print("=" * 70)
    print(name, "page rect", page.rect)
    draws = page.get_drawings()
    for di, d in enumerate(draws):
        rect = d["rect"]
        items = d.get("items", [])
        fill = d.get("fill")
        color = d.get("color")
        n_l = sum(1 for it in items if it[0] == "l")
        n_c = sum(1 for it in items if it[0] == "c")
        n_re = sum(1 for it in items if it[0] == "re")
        # print every drawing with >=2 line items (a polyline = several line
        # items in one drawing) OR any diagonal line
        diag = 0
        for it in items:
            if it[0] == "l":
                p1, p2 = it[1], it[2]
                if abs(p2.x - p1.x) > 0.5 and abs(p2.y - p1.y) > 0.5:
                    diag += 1
        if n_l >= 2 or diag > 0:
            print(f"d{di} rect=({rect.x0:.2f},{rect.y0:.2f},{rect.x1:.2f},{rect.y1:.2f}) "
                  f"items l={n_l} c={n_c} re={n_re} fill={fill} color={color}")
            for it in items:
                if it[0] == "l":
                    p1, p2 = it[1], it[2]
                    print(f"    line ({p1.x:.2f},{p1.y:.2f})->({p2.x:.2f},{p2.y:.2f})")
                elif it[0] == "c":
                    print(f"    bezier " + " ".join(f"({p.x:.2f},{p.y:.2f})" for p in it[1:5]))
                elif it[0] == "re":
                    print(f"    rect {it[1]}")
    doc.close()
