# -*- coding: utf-8 -*-
"""Dump chart1.xml and chart2.xml full text to files, then grep for
title / legend / autoTitleDeleted related elements."""
import zipfile, os, re

base = r"pipeline_data\pptx_probes"
for name in ("chart1", "chart2"):
    pptx = os.path.join(base, name, name + ".pptx")
    with zipfile.ZipFile(pptx) as z:
        xml = z.read("ppt/charts/chart1.xml").decode("utf-8")
    out = os.path.join(base, name, name + "_xml_full.txt")
    with open(out, "w", encoding="utf-8") as f:
        f.write(xml)
    print(f"=== {name} chart1.xml len={len(xml)}")
    # print everything before <c:plotArea> (title/legend area)
    m = re.search(r"<c:plotArea", xml)
    head = xml[:m.start()] if m else xml
    print("--- pre-plotArea ---")
    print(head)
    print()
