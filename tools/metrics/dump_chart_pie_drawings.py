#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Dump ALL drawing paths on chart_pie.pdf page 1: per-path colour + item types,
to see whether pie slices carry an outline / how the fill paths are structured."""
import fitz

PDF = r"pipeline_data\pptx_probes\chart_pie\chart_pie.pdf"


def main():
    doc = fitz.open(PDF)
    page = doc[0]
    draws = page.get_drawings()
    print(f"page 0: {len(draws)} drawing paths")
    for i, d in enumerate(draws):
        fill = d.get("fill")
        stroke = d.get("color")
        w = d.get("width")
        even_odd = d.get("even_odd")
        types = [it[0] for it in d["items"]]
        bbox = d["rect"]
        # count item types
        from collections import Counter

        cnt = Counter(types)
        print(
            f"  [{i}] fill={fill} stroke={stroke} width={w} even_odd={even_odd} "
            f"items={dict(cnt)} rect=({round(bbox.x0,1)},{round(bbox.y0,1)},",
            f"{round(bbox.x1,1)},{round(bbox.y1,1)})",
        )
        # print the first path's full item detail for the biggest fill
        if fill and d["rect"].width > 50:
            for it in d["items"][:3]:
                print(f"      item {it[0]}: {it}")


if __name__ == "__main__":
    main()
