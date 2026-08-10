"""Measure chart_stacked100.pdf with fitz — text spans + vector drawings.

Prints per-item so the 100%-STACKING rule (segment heights relative to the
category SUM), per-series colours, the value-axis scale (0..100?), auto-title
and legend are fully exposed.
"""
import os
import sys

import fitz

sys.stdout.reconfigure(encoding="utf-8")


def main():
    base = r"pipeline_data\pptx_probes\chart_stacked100"
    pdf_path = os.path.join(base, "chart_stacked100.pdf")

    doc = fitz.open(pdf_path)
    page = doc[0]
    print("page.rect =", page.rect)

    print("\n=== TEXT SPANS ===")
    raw = page.get_text("rawdict")
    for block in raw["blocks"]:
        if block.get("type", 0) != 0:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                txt = "".join(c["c"] for c in span["chars"])
                if not txt.strip():
                    continue
                o = span["origin"]
                col = span.get("color", 0)
                print(
                    f"'{txt}' x={o[0]:.2f} y={o[1]:.2f} "
                    f"size={span['size']:.2f} color=#{col:06x} font={span['font']}"
                )

    print("\n=== VECTOR DRAWINGS ===")
    drawings = page.get_drawings()
    print("n_paths =", len(drawings))
    for p in drawings:
        r = p["rect"]
        fill = p.get("fill")
        stroke = p.get("color")
        items = p["items"]
        print(
            f"rect=({r.x0:.1f},{r.y0:.1f},{r.x1:.1f},{r.y1:.1f}) "
            f"w={r.width:.1f} h={r.height:.1f} "
            f"fill={fill} stroke={stroke} n_items={len(items)}"
        )
        for it in items:
            print("   ", it)


if __name__ == "__main__":
    main()
