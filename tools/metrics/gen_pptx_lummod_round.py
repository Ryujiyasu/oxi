# -*- coding: utf-8 -*-
"""Probe: which byte does PowerPoint draw for an HSL modifier (lumMod / lumOff / satMod / hueOff)?

Each arm is a rectangle whose spPr solidFill is `<a:srgbClr val=BASE>` + the arm's modifiers,
written straight into the slide XML (PowerPoint only authors the boxes). PowerPoint exports
the PDF; the fill colour of each rectangle is read back and printed next to the exact HSL
value, so ties (x.5) and near-ties can be classified.

NEVER run this while the renderer is producing PNGs (`pptx_com_render_must_not_overlap`).

Usage: python tools/metrics/gen_pptx_lummod_round.py [sat|grey]
"""
import json
import sys
import re
import shutil
import zipfile
from pathlib import Path

import pymupdf
import win32com.client

REPO = Path(__file__).resolve().parents[2]
OUT = REPO / "pipeline_data" / "pptx_probes" / "lummod_round"
OUT.mkdir(parents=True, exist_ok=True)
RAW = OUT / "raw.pptx"
PPTX = OUT / "probe_lummod_round.pptx"
PDF = OUT / "probe_lummod_round.pdf"

BASES = ["FFFFFF", "000000", "808080", "7F7F7F", "A5A5A5", "4472C4", "156082", "E97132",
         "0E2841", "FFC000", "70AD47", "E8E8E8", "44546A", "196B24", "0F9ED5", "A02B93"]
MODS = [
    [("lumMod", 50000)], [("lumMod", 75000)], [("lumMod", 85000)], [("lumMod", 95000)],
    [("lumMod", 90000)], [("lumMod", 65000)], [("lumMod", 25000)], [("lumMod", 10000)],
    [("lumMod", 50000), ("lumOff", 50000)], [("lumMod", 75000), ("lumOff", 25000)],
    [("lumMod", 65000), ("lumOff", 35000)], [("lumMod", 85000), ("lumOff", 15000)],
    [("lumMod", 95000), ("lumOff", 5000)], [("lumMod", 20000), ("lumOff", 80000)],
    [("lumMod", 40000), ("lumOff", 60000)], [("lumMod", 60000), ("lumOff", 40000)],
    [("lumMod", 10000), ("lumOff", 90000)], [("lumMod", 50000), ("lumOff", 25000)],
    [("lumOff", 10000)], [("lumOff", -10000)], [("satMod", 50000)], [("satMod", 150000)],
    [("hueOff", 1800000)], [("hueOff", 7200000)], [("lum", 50000)], [("sat", 50000)],
]
COLS, PER_SLIDE = 13, 52  # 4 rows of 13

if len(sys.argv) > 1 and sys.argv[1] == "sat":  # achromatic / hue-0 inputs under sat & hue modifiers
    OUT = REPO / "pipeline_data" / "pptx_probes" / "satmod_hue"
    OUT.mkdir(parents=True, exist_ok=True)
    RAW, PPTX, PDF = OUT / "raw.pptx", OUT / "probe_satmod_hue.pptx", OUT / "probe_satmod_hue.pdf"
    BASES = ["808080", "404040", "C0C0C0", "E8E8E8", "FF0000", "800000", "FF8080", "00FF00", "0000FF", "FFFF00"]
    MODS = [[("sat", 25000)], [("sat", 50000)], [("sat", 75000)], [("sat", 100000)],
            [("sat", 50000), ("hueOff", 3600000)], [("sat", 50000), ("hueOff", 7200000)],
            [("sat", 50000), ("hueOff", 10800000)], [("sat", 50000), ("hue", 2400000)],
            [("hue", 7200000)], [("hueOff", 21000000)], [("satMod", 200000)], [("satOff", 30000)],
            [("satOff", -30000)]]

if len(sys.argv) > 1 and sys.argv[1] == "grey":  # hue sweep under sat on achromatic bases
    OUT = REPO / "pipeline_data" / "pptx_probes" / "grey_hue"
    OUT.mkdir(parents=True, exist_ok=True)
    RAW, PPTX, PDF = OUT / "raw.pptx", OUT / "probe_grey_hue.pptx", OUT / "probe_grey_hue.pdf"
    BASES = ["C0C0C0", "E8E8E8", "808080", "C0C0C1"]
    MODS = ([[("sat", 50000), ("hueOff", k * 900000)] for k in range(24)]
            + [[("hue", k * 1800000), ("sat", 50000)] for k in range(12)]
            + [[("sat", 20000), ("hueOff", k * 1800000)] for k in range(12)])

arms = [(b, m) for b in BASES for m in MODS]

app = win32com.client.Dispatch("PowerPoint.Application")
try:
    pres = app.Presentations.Add(WithWindow=False)
    try:
        for i, _ in enumerate(arms):
            si, k = divmod(i, PER_SLIDE)
            if k == 0:
                pres.Slides.Add(si + 1, 12)
            s = pres.Slides(si + 1)
            r, c = divmod(k, COLS)
            t = s.Shapes.AddShape(1, 10 + c * 72, 20 + r * 120, 60, 100)
            t.Name = f"arm_{i}"
        pres.SaveAs(str(RAW))
    finally:
        pres.Close()

    tmp = OUT / "tmp.pptx"
    with zipfile.ZipFile(RAW) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if re.match(r"ppt/slides/slide\d+\.xml$", item.filename):
                x = data.decode("utf-8")

                def fix(m):
                    sp = m.group(0)
                    nm = re.search(r'name="arm_(\d+)"', sp)
                    if not nm:
                        return sp
                    base, mods = arms[int(nm.group(1))]
                    inner = "".join(f'<a:{n} val="{v}"/>' for n, v in mods)
                    fill = f'<a:solidFill><a:srgbClr val="{base}">{inner}</a:srgbClr></a:solidFill><a:ln><a:noFill/></a:ln>'
                    sp = re.sub(r"(</a:prstGeom>)", r"\1" + fill, sp, count=1)
                    sp = re.sub(r"<p:style>.*?</p:style>", "", sp, flags=re.S)
                    return sp

                x = re.sub(r"<p:sp>.*?</p:sp>", fix, x, flags=re.S)
                data = x.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, PPTX)

    pres = app.Presentations.Open(str(PPTX), WithWindow=False)
    try:
        pres.SaveAs(str(PDF), 32)
    finally:
        pres.Close()
finally:
    app.Quit()

doc = pymupdf.open(PDF)
rows = []
for si, pg in enumerate(doc):
    rects = []
    for d in pg.get_drawings():
        if d.get("fill") is None or d["rect"].width < 30:
            continue
        rects.append((round(d["rect"].y0), round(d["rect"].x0), d["fill"]))
    rects.sort()
    for k, (_, _, f) in enumerate(rects):
        i = si * PER_SLIDE + k
        if i >= len(arms):
            break
        base, mods = arms[i]
        rows.append({"base": base, "mods": mods, "pdf": [round(v * 255, 3) for v in f]})
(OUT / "result.json").write_text(json.dumps(rows, indent=1))
print(len(rows), "arms of", len(arms))
