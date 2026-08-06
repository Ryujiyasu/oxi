#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Measure the pie circle geometry for each page of chart_pie3.pdf.

For each accent wedge, extract:
  - center point (the shared 'l' endpoint every wedge closes back to)
  - the arc start/end points (first/last Point of the 'c' runs)
From East (which starts at 12 o'clock) we get top = cy - r, so r = cy - top.
"""
import json

PDF = r"C:\Users\ryuji\oxi-main\pipeline_data\pptx_probes\chart_pie3\chart_pie3.pdf"

ACCENT = {
    (0.31, 0.506, 0.741): "a1",
    (0.753, 0.314, 0.302): "a2",
    (0.608, 0.733, 0.349): "a3",
}


def norm_color(c):
    if not c:
        return None
    return (round(c[0], 3), round(c[1], 3), round(c[2], 3))


def main():
    import fitz
    doc = fitz.open(PDF)
    out = []
    for pno in range(len(doc)):
        wedges = {}
        for d in doc[pno].get_drawings():
            fill = norm_color(d.get("fill"))
            if fill not in ACCENT:
                continue
            csegs = []
            lsegs = []
            for it in d["items"]:
                kind = it[0]
                if kind == "c":
                    csegs.append((tuple(it[1]), tuple(it[4])))
                elif kind == "l":
                    lsegs.append((tuple(it[1]), tuple(it[2])))
            if not csegs:
                continue
            arc_start = csegs[0][0]
            arc_end = csegs[-1][1]
            center = None
            if len(lsegs) >= 2:
                a, b = lsegs[0]
                c, dd = lsegs[1]
                if a == c or a == dd:
                    center = a
                elif b == c or b == dd:
                    center = b
            if center is None and lsegs:
                center = lsegs[0][0]
            wedges[ACCENT[fill]] = {
                "center": [round(v, 2) for v in center],
                "arc_start": [round(v, 2) for v in arc_start],
                "arc_end": [round(v, 2) for v in arc_end],
            }
        if "a1" not in wedges:
            out.append({"page": pno, "no_a1": True})
            continue
        top = wedges["a1"]["arc_start"][1]
        cx, cy = wedges["a1"]["center"]
        r = round(cy - top, 2)
        out.append({
            "page": pno,
            "center": [cx, cy],
            "top": round(top, 2),
            "r": r,
            "bottom": round(cy + r, 2),
            "wedges": wedges,
        })
    print(json.dumps(out, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
