#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Raw dump of wedge items on page0/page2 to see exact fitz item shapes."""
import fitz

PDF = r"pipeline_data\pptx_probes\chart_pie2\chart_pie2.pdf"

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
    doc = fitz.open(PDF)
    for pno in (0, 2):
        print(f"===== page{pno} =====")
        for d in doc[pno].get_drawings():
            fill = norm_color(d.get("fill"))
            if fill not in ACCENT:
                continue
            print(f"-- {ACCENT[fill]} fill={fill}")
            for it in d["items"]:
                kind = it[0]
                rest = it[1:]
                # show raw structure
                print("   kind", kind, "len(rest)", len(rest))
                for r in rest:
                    if isinstance(r, (list, tuple)):
                        print("     point", [round(float(v), 2) for v in r])
                    else:
                        print("     val", r)


if __name__ == "__main__":
    main()
