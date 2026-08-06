# -*- coding: utf-8 -*-
"""Compare chart1.xml vs chart2.xml legend presence (spec derivation)."""
import zipfile

for name, path in [("chart1", r"pipeline_data\pptx_probes\chart1\chart1.pptx"),
                   ("chart2", r"pipeline_data\pptx_probes\chart2\chart2.pptx")]:
    z = zipfile.ZipFile(path)
    xml = z.read("ppt/charts/chart1.xml").decode("utf-8")
    has_legend = "<c:legend" in xml
    n_ser = xml.count("<c:ser>")
    print(f"{name}: legend={'YES' if has_legend else 'NO'}  n_ser={n_ser}")
    # print a compact strip around legend if present
    idx = xml.find("<c:legend")
    if idx >= 0:
        print("   legend xml:", xml[idx:idx+400])
