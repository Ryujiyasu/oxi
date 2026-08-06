#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Dump chart2.xml <c:title> subtree + page1 (B) rendered text via rawdict."""
import zipfile
import re
import fitz

PPTX = r"pipeline_data\pptx_probes\chart_pie3\chart_pie3.pptx"
PDF = r"pipeline_data\pptx_probes\chart_pie3\chart_pie3.pdf"


def main():
    z = zipfile.ZipFile(PPTX)
    xml = z.read("ppt/charts/chart2.xml").decode("utf-8", "ignore")
    m = re.search(r"<c:title>.*?</c:title>", xml, re.S)
    print("== chart2 <c:title> ==")
    print(m.group(0) if m else "NO <c:title>")

    print("\n== page1 (B) rendered text (rawdict) ==")
    doc = fitz.open(PDF)
    page = doc[1]
    td = page.get_text("rawdict")
    for block in td["blocks"]:
        if block["type"] != 0:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                txt = "".join(c["c"] for c in span.get("chars", []))
                if not txt.strip():
                    continue
                font = span["font"]
                size = round(span["size"], 2)
                ox = round(span["origin"][0], 2)
                oy = round(span["origin"][1], 2)
                print(f"  '{txt}' font={font} size={size} origin=({ox},{oy})")


if __name__ == "__main__":
    main()
