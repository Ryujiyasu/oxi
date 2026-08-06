#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Inspect chart_pie3 chart parts: pie / title element / autoTitleDeleted / ser count."""
import zipfile
import re

PPTX = r"pipeline_data\pptx_probes\chart_pie3\chart_pie3.pptx"


def main():
    z = zipfile.ZipFile(PPTX)
    charts = sorted(n for n in z.namelist() if "/charts/chart" in n and n.endswith(".xml"))
    for c in charts:
        xml = z.read(c).decode("utf-8", "ignore")
        pie = "<c:pieChart>" in xml
        title_elem = "<c:title>" in xml
        m = re.search(r"autoTitleDeleted.{0,20}val=\"(\d)\"", xml)
        atd = m.group(1) if m else None
        ser_count = xml.count("<c:ser>")
        vals = re.findall(r"<c:v>([^<]*)</c:v>", xml)
        print(c, "pie=", pie, "title_elem=", title_elem, "autoTitleDeleted=", atd,
              "ser_count=", ser_count, "first_vals=", vals[:8])


if __name__ == "__main__":
    main()
