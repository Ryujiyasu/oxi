# -*- coding: utf-8 -*-
"""Dump chart1.xml from chart1.pptx fully (check legend / series name)."""
import zipfile
z = zipfile.ZipFile(r"pipeline_data\pptx_probes\chart1\chart1.pptx")
names = [x for x in z.namelist() if "chart" in x.lower()]
print("parts:", names)
print(z.read("ppt/charts/chart1.xml").decode("utf-8"))
