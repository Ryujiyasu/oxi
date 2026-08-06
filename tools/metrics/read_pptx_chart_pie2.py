#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Read the multi-slide pie-chart PDF (Word render truth) with fitz.

Per page: chart vectors (get_drawings p['items'] per-item) + text spans.
"""
import json
import sys

import fitz

PDF = r"pipeline_data\pptx_probes\chart_pie2\chart_pie2.pdf"


def read_page(page):
    drawings = []
    for d in page.get_drawings():
        item_list = []
        for it in d["items"]:
            kind = it[0]
            if kind in ("re", "qu"):
                item_list.append((kind, [round(v, 2) for v in it[1]]))
            elif kind in ("c", "l", "e"):
                item_list.append((kind, [round(v, 2) for v in it[1]]))
            elif kind == "p":
                item_list.append((kind, [[round(p[0], 2), round(p[1], 2)] for p in it[1]]))
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
    return {"drawings": drawings, "texts": texts}


def main():
    doc = fitz.open(PDF)
    out = {f"page{i}": read_page(doc[i]) for i in range(len(doc))}
    print(json.dumps(out, ensure_ascii=False))
    for i in range(len(doc)):
        print(f"# page{i}: drawings {len(out[f'page{i}']['drawings'])} / texts {len(out[f'page{i}']['texts'])}",
              file=sys.stderr)


if __name__ == "__main__":
    main()
