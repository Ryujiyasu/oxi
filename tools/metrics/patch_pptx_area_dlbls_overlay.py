# -*- coding: utf-8 -*-
"""python-pptx writes <c:overlay val="0"/> even for a bare
`chart.has_legend = True`, so slide D7 cannot exercise the OVERLAY legend
path.  Rewrite chart7.xml in place so its legend becomes the bare
self-closing <c:legend/> form (no overlay child) that pie/doughnut showed
Word treats as an overlay (no band, plot stays frame-centred)."""
import sys, os, re, shutil, zipfile

sys.stdout.reconfigure(encoding="utf-8")

src = r"pipeline_data\pptx_probes\chart_area_dlbls\chart_area_dlbls.pptx"
tmp = src + ".tmp"

zin = zipfile.ZipFile(src)
items = [(i, zin.read(i.filename)) for i in zin.infolist()]
zin.close()

hit = False
with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for info, blob in items:
        if info.filename == "ppt/charts/chart7.xml":
            s = blob.decode("utf-8")
            s2 = re.sub(r"<c:legend>.*?</c:legend>", "<c:legend/>", s, flags=re.S)
            hit = s2 != s
            blob = s2.encode("utf-8")
        zout.writestr(info, blob)

assert hit, "chart7.xml legend not rewritten"
shutil.move(tmp, src)
print("patched D7 legend -> bare <c:legend/>:", src, os.path.getsize(src))
