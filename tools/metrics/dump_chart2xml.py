# -*- coding: utf-8 -*-
"""Dump chart2.pptx internal chart XML (for spec derivation)."""
import zipfile
z = zipfile.ZipFile(r"pipeline_data\pptx_probes\chart2\chart2.pptx")
names = [x for x in z.namelist() if "chart" in x.lower()]
print("chart parts:", names)
for name in names:
    print("=== " + name + " ===")
    print(z.read(name).decode("utf-8"))
