#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Compare chart_pie / chart_pie2 / chart_pie3 XML: title elem, autoTitleDeleted,
ser count, pieChart — to settle the title-draw condition."""
import zipfile
import re

PROBES = [
    r"pipeline_data\pptx_probes\chart_pie\chart_pie.pptx",
    r"pipeline_data\pptx_probes\chart_pie2\chart_pie2.pptx",
    r"pipeline_data\pptx_probes\chart_pie3\chart_pie3.pptx",
]


def pdf_text(pdf):
    import fitz
    doc = fitz.open(pdf)
    out = []
    for pi in range(len(doc)):
        page = doc[pi]
        d = page.get_text("rawdict")
        spans = []
        for block in d["blocks"]:
            for line in block.get("lines", []):
                for span in line["spans"]:
                    t = "".join(ch["c"] for ch in span.get("chars", []))
                    if t.strip():
                        spans.append((round(span["origin"][0], 2), round(span["origin"][1], 2),
                                     span.get("font"), round(span.get("size", 0), 2), t))
        out.append(spans)
    return out


def main():
    for pptx in PROBES:
        print("====", pptx)
        pdf = pptx.replace(".pptx", ".pdf")
        import os
        if os.path.exists(pdf):
            print("  --- PDF text per page:")
            for pi, spans in enumerate(pdf_text(pdf)):
                print(f"  p{pi+1}:", spans[:6])
        z = zipfile.ZipFile(pptx)
        charts = sorted(n for n in z.namelist() if "/charts/chart" in n and n.endswith(".xml"))
        for c in charts:
            xml = z.read(c).decode("utf-8", "ignore")
            pie = "<c:pieChart>" in xml
            title_elem = "<c:title>" in xml
            legend_elem = "<c:legend>" in xml
            m = re.search(r"autoTitleDeleted.{0,20}val=\"(\d)\"", xml)
            atd = m.group(1) if m else None
            ser_count = xml.count("<c:ser>")
            vals = re.findall(r"<c:v>([^<]*)</c:v>", xml)
            print(c, "pie=", pie, "title=", title_elem, "legend=", legend_elem,
                  "autoTitleDeleted=", atd, "ser_count=", ser_count, "vals=", vals[:8])


if __name__ == "__main__":
    main()
