# -*- coding: utf-8 -*-
"""Read chart_datalabel_line / chart_datalabel_pie Word PDFs: dump text
spans (the data labels '19' '21' '17' etc. and the axis/category labels) so
the data-label placement / format / font can be measured for the line and
pie branches."""
import sys
sys.stdout.reconfigure(encoding="utf-8")
import fitz


def read(base_name, n_pages):
    doc = fitz.open(rf"pipeline_data\pptx_probes\{base_name}\{base_name}.pdf")
    for pi in range(min(n_pages, doc.page_count)):
        page = doc[pi]
        print(f"=== {base_name} page {pi} rect = {page.rect} ===")
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
    doc.close()


read("chart_datalabel_line", 3)
read("chart_datalabel_pie", 3)
