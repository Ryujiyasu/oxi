#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Read the pie-chart PDF (Word render truth) with fitz.

Draws chart vectors (get_drawings) and text (rawdict -> chars) to stdout as JSON.
12700 EMU = 1pt. The pie sectors appear as path items (re/curve/c/...).
"""
import json
import sys

import fitz

PDF = r"pipeline_data\pptx_probes\chart_pie\chart_pie.pdf"


def main():
    doc = fitz.open(PDF)
    page = doc[0]
    print("PAGE rect:", page.rect, file=sys.stderr)

    drawings = []
    for d in page.get_drawings():
        item_list = []
        for it in d["items"]:
            kind = it[0]
            if kind == "re":
                item_list.append(("re", [round(v, 2) for v in it[1]]))
            elif kind == "qu":
                item_list.append(("qu", [round(v, 2) for v in it[1]]))
            elif kind == "c":
                item_list.append(("c", [round(v, 2) for v in it[1]]))
            elif kind == "l":
                item_list.append(("l", [round(v, 2) for v in it[1]]))
            elif kind == "e":
                item_list.append(("e", [round(v, 2) for v in it[1]]))
            elif kind == "p":
                pts = [[round(p[0], 2), round(p[1], 2)] for p in it[1]]
                item_list.append(("p", pts))
            else:
                item_list.append((kind, it[1]))
        drawings.append({
            "rect": [round(v, 2) for v in d["rect"]],
            "fill": d.get("fill"),
            "color": d.get("color"),
            "items": item_list,
        })

    texts = []
    for blk in page.get_text("rawdict")["blocks"]:
        for line in blk.get("lines", []):
            for span in line["spans"]:
                chars = "".join(ch["c"] for ch in span["chars"])
                texts.append({
                    "text": chars,
                    "origin": [round(span["origin"][0], 2), round(span["origin"][1], 2)],
                    "size": round(span["size"], 2),
                    "font": span["font"],
                })

    out = {"drawings": drawings, "texts": texts}
    print(json.dumps(out, ensure_ascii=False))
    print(f"\n# drawings: {len(drawings)} / # text spans: {len(texts)}", file=sys.stderr)


if __name__ == "__main__":
    main()
