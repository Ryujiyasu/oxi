# -*- coding: utf-8 -*-
"""Derive Word's beside/below cutoff M for a wrapTight TEXTBOX lane.

Takes reports__0013bcb8's real page-1 title textbox verbatim and sweeps only its
`wp:extent cx`, so the free lane to its right moves through 24..340pt while every
other property (distL/distR = 9.0pt, wrapTight bothSides with a RECTANGULAR
wrapPolygon, positionH column, the anchoring paragraph, the section geometry)
stays byte-identical to the real document.

Read-out: the x0 of the "Research Article" line.  x0 == left margin (70.85) means
Word pushed it BELOW the band; a larger x0 means Word kept it BESIDE in the lane.

  python tools/metrics/_pb_tightlane_gen.py gen     # build the arm docx files
  python tools/metrics/_pb_tightlane_gen.py bake    # one Word instance, all arms
  python tools/metrics/_pb_tightlane_gen.py read    # x0 per arm + the flip
"""
import os
import re
import shutil
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8")

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
SRC = os.path.join(
    REPO, "pipeline_data", "docx_corpus", "en", "reports", "0013bcb8b619da89.docx"
)
OUT = os.path.join(REPO, "pipeline_data", "_pb_tightlane")

# real geometry: content 453.8pt, box left 0.65pt, distR 9.0pt
# lane = 453.8 - 0.65 - cx_pt - 9.0  =>  cx_pt = 444.15 - lane
CONTENT = 453.8
BOX_LEFT = 0.65
DIST_R = 9.0
ORIG_CX = 5046345  # EMU (397.35pt), the real box
# The full sweep that was baked: Word puts the paragraph BELOW the band in
# EVERY one of these arms — the side lane is never used, however wide.
LANES = [24.0, 30.0, 36.0, 40.0, 44.0, 46.8, 50.0, 56.0, 64.0, 76.0,
         80.0, 84.0, 88.0, 96.0, 110.0, 130.0, 180.0, 260.0, 340.0]


def cx_for(lane):
    return int(round((CONTENT - BOX_LEFT - DIST_R - lane) * 12700))


def gen():
    os.makedirs(OUT, exist_ok=True)
    blob = open(SRC, "rb").read()
    with zipfile.ZipFile(SRC) as z:
        doc = z.read("word/document.xml").decode("utf8")
    assert f'cx="{ORIG_CX}"' in doc, "the real extent is gone - re-derive"
    for lane in LANES:
        cx = cx_for(lane)
        name = f"lane{lane:04.1f}".replace(".", "_")
        dst = os.path.join(OUT, name + ".docx")
        new = doc.replace(f'cx="{ORIG_CX}"', f'cx="{cx}"', 1)
        shutil.copyfile(SRC, dst)
        # rewrite document.xml in place
        tmp = dst + ".tmp"
        with zipfile.ZipFile(dst) as zin, zipfile.ZipFile(
            tmp, "w", zipfile.ZIP_DEFLATED
        ) as zout:
            for it in zin.infolist():
                data = zin.read(it.filename)
                if it.filename == "word/document.xml":
                    data = new.encode("utf8")
                zout.writestr(it, data)
        os.replace(tmp, dst)
        print(f"  {name}.docx  lane={lane:5.1f}  cx={cx}  ({cx/12700:.2f}pt)")


def bake():
    import win32com.client as win32

    app = win32.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        for f in sorted(os.listdir(OUT)):
            if not f.endswith(".docx"):
                continue
            src = os.path.join(OUT, f)
            pdf = src[:-5] + ".pdf"
            d = app.Documents.Open(src, ReadOnly=True)
            d.ExportAsFixedFormat(OutputFileName=pdf, ExportFormat=17)
            d.Close(False)
            print("  baked", os.path.basename(pdf))
    finally:
        app.Quit()


def read():
    import fitz

    margin = 70.85
    rows = []
    for f in sorted(os.listdir(OUT)):
        if not f.endswith(".pdf"):
            continue
        lane = float(f[4:-4].replace("_", "."))
        d = fitz.open(os.path.join(OUT, f))
        x0 = None
        for b in d[0].get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    if s["text"].strip().startswith("Research Article"):
                        x0 = s["bbox"][0]
        d.close()
        rows.append((lane, x0))
    rows.sort()
    prev = None
    print(f"{'lane':>7} {'x0':>8}  verdict")
    for lane, x0 in rows:
        if x0 is None:
            print(f"{lane:7.1f} {'-':>8}  (not found)")
            continue
        v = "BELOW" if abs(x0 - margin) < 1.0 else "beside"
        mark = ""
        if prev is not None and prev != v:
            mark = "   <<< FLIP"
        print(f"{lane:7.1f} {x0:8.2f}  {v}{mark}")
        prev = v


if __name__ == "__main__":
    {"gen": gen, "bake": bake, "read": read}[sys.argv[1]]()
