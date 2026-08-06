# -*- coding: utf-8 -*-
"""Dump chart_line.xml axis/series config to understand why category labels
are absent and how the plot area is configured."""
import sys, zipfile, re
sys.stdout.reconfigure(encoding="utf-8")

pptx = r"pipeline_data\pptx_probes\chart_line\chart_line.pptx"
z = zipfile.ZipFile(pptx)
names = [n for n in z.namelist() if "chart" in n]
print("chart parts:", names)

chart = [n for n in names if n.endswith(".xml") and "charts" in n]
for n in chart:
    xml = z.read(n).decode("utf-8")
    print("==== %s (%d bytes) ====" % (n, len(xml)))
    # pretty-print loosely
    xml2 = re.sub(r"><", ">\n<", xml)
    # only print structural elements (skip data caches to reduce noise)
    for line in xml2.splitlines():
        if any(k in line for k in [
            "lineChart", "barDir", "grouping", "varyColors", "ser", "catAx",
            "valAx", "axId", "axPos", "tickLblPos", "tickMarkSkip",
            "catLbl", "numFmt", "crosses", "majorTickMark", "minorTickMark",
            "delete", "layout", "manualLayout", "autoTitleDeleted", "legend",
            "showLegendKey", "spPr", "ln", "marker", "symbol", "size",
            "scaling", "max", "min", "txPr", "defRPr", "lblAlgn", "lblOffset",
            "crossBetween", "majorGridlines", "minorGridlines",
        ]):
            # strip cache value lists
            if "c:pt" in line or "c:v>" in line and "numCache" in n:
                line = re.sub(r"<c:v>[^<]*</c:v>", "<c:v>...</c:v>", line)
            print(line)
