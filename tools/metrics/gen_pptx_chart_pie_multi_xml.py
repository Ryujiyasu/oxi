#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Manually inject a SECOND <c:ser> into a single-series pie chart pptx.

python-pptx cannot emit multi-series PIE (it drops add_series #2), so we
rewrite ppt/charts/chart1.xml inside the pptx zip.  The embedded cache
(strCache/numCache) carries both series; Word renders the cached values.

OUT: pipeline_data/pptx_probes/chart_pie_multi/chart_pie_multi.pptx
"""
import shutil
import zipfile
import os

SRC = r"pipeline_data\pptx_probes\chart_pie_multi\chart_pie_multi.pptx"
OUT = SRC


def build_multi_xml(single_xml: str) -> str:
    # Insert a second <c:ser> right before the closing </c:pieChart>.
    ser2 = (
        '<c:ser><c:idx val="1"/><c:order val="1"/>'
        "<c:tx><c:strRef><c:f>Sheet1!$D$1</c:f><c:strCache><c:ptCount val=\"1\"/>"
        '<c:pt idx="0"><c:v>Cost</c:v></c:pt></c:strCache></c:strRef></c:tx>'
        "<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f><c:strCache>"
        '<c:ptCount val="3"/><c:pt idx="0"><c:v>East</c:v></c:pt>'
        '<c:pt idx="1"><c:v>West</c:v></c:pt><c:pt idx="2"><c:v>Midwest</c:v></c:pt>'
        "</c:strCache></c:strRef></c:cat>"
        "<c:val><c:numRef><c:f>Sheet1!$D$2:$D$4</c:f><c:numCache>"
        '<c:formatCode>General</c:formatCode><c:ptCount val="3"/>'
        '<c:pt idx="0"><c:v>10.5</c:v></c:pt><c:pt idx="1"><c:v>11.2</c:v></c:pt>'
        '<c:pt idx="2"><c:v>8.5</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>'
    )
    assert "</c:pieChart>" in single_xml, "pieChart close tag not found"
    return single_xml.replace("</c:pieChart>", ser2 + "</c:pieChart>")


def main():
    tmp = OUT + ".tmp"
    with zipfile.ZipFile(SRC, "r") as zin:
        names = zin.namelist()
        chart_name = next(n for n in names if n.endswith("chart1.xml"))
        single = zin.read(chart_name).decode("utf-8")
        multi = build_multi_xml(single)
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for n in names:
                if n == chart_name:
                    zout.writestr(n, multi)
                else:
                    zout.writestr(n, zin.read(n))
    shutil.move(tmp, OUT)
    with zipfile.ZipFile(OUT) as z:
        xml = z.read("ppt/charts/chart1.xml").decode("utf-8")
        print("ser count:", xml.count("<c:ser>"))
    print(f"Saved {OUT}")


if __name__ == "__main__":
    main()
