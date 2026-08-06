#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Inspect the chart parts of chart_pie2.pptx: series/category counts + catAx/valAx."""
import json
import re
import zipfile

PPTX = r"pipeline_data\pptx_probes\chart_pie2\chart_pie2.pptx"


def main():
    z = zipfile.ZipFile(PPTX)
    out = {}
    for n in sorted(z.namelist()):
        if "charts" in n and n.endswith(".xml"):
            xml = z.read(n).decode("utf-8")
            ser_names = re.findall(
                r"<c:tx>\s*<c:strRef>.*?</c:strRef>\s*</c:tx>|"
                r"<c:tx>\s*<c:v>([^<]*)</c:v>\s*</c:tx>", xml, re.S
            )
            # series name via strCache
            str_cache = re.findall(
                r"<c:strCache>.*?<c:pt idx=\"0\">\s*<c:v>([^<]+)</c:v>", xml, re.S
            )
            num_pts = re.findall(
                r"<c:numCache>.*?<c:ptCount val=\"(\d+)\"", xml, re.S
            )
            cat_count = re.findall(r"<c:ptCount val=\"(\d+)\"", xml)
            out[n] = {
                "ser_count": xml.count("<c:ser>"),
                "ser_names": str_cache,
                "has_barDir": "<c:barDir" in xml,
                "has_pieChart": "<c:pieChart" in xml,
                "pt_counts": num_pts,
            }
    print(json.dumps(out, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
