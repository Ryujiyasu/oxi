# -*- coding: utf-8 -*-
"""Does Word's cursor LOSE the pixel remainder, or only round for display?

Both readings fit a single paragraph:
  EXACT   the cursor accumulates the true multiplied height h and only the
          reported/painted position is rounded to the 96dpi pixel (0.75pt),
          so the displayed y never drifts more than 0.375pt from exact
  SNAP    each paragraph advances by round(n*h/0.75)*0.75, i.e. the remainder
          is DISCARDED at every paragraph, so the error accumulates

They separate over a RUN of single-line paragraphs whose h has a large
fractional pixel part.  Arial 12 x1.15: h = 15.87 = 21.16px -> EXACT predicts
20 paragraphs span 317.4, SNAP predicts 20 x 21px = 315.0 (2.4pt apart, far
outside Info6's 0.05 rounding).

Each combo gets its own page so the run starts at a page top.

  python _pb_pxrun_gen.py gen
  python _pb_pxrun_gen.py read
"""
import os
import re
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_pxgrid")
DOCX = os.path.join(OUT, "pxrun.docx")
PX = 0.75
NRUN = 20

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS, STYLES  # noqa: E402

# (font, half-point size, multiplier) -- the fractional pixel part of h is
# what makes a combo informative, so keep a spread of them.
COMBOS = [
    ("Arial", 24, 276),             # 15.87 = 21.16px   frac .16
    ("Arial", 24, 240),             # 13.80 = 18.40px   frac .40
    ("Arial", 22, 276),             # 14.55 = 19.39px   frac .39
    ("Times New Roman", 24, 360),   # 20.70 = 27.60px   frac .60
    ("Times New Roman", 24, 259),   # 14.89 = 19.86px   frac .86
    ("Times New Roman", 22, 240),   # 12.65 = 16.87px   frac .87
    ("Calibri", 22, 276),
    ("Calibri", 24, 240),
]


def para(tag, font, sz, ml):
    rpr = ('<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/><w:sz w:val="%d"/>'
           '<w:szCs w:val="%d"/></w:rPr>' % (font, font, font, sz, sz))
    return ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="%d" w:lineRule="auto"/>'
            '<w:widowControl w:val="0"/>%s</w:pPr><w:r>%s<w:t>%s</w:t></w:r></w:p>'
            % (ml, rpr, rpr, tag))


def sect(last):
    inner = "" if last else '<w:type w:val="nextPage"/>'
    s = ('<w:sectPr>%s<w:pgSz w:w="11906" w:h="16838"/>'
         '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
         'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>' % inner)
    return s if last else '<w:p><w:pPr>%s</w:pPr></w:p>' % s


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ci, (f, sz, ml) in enumerate(COMBOS):
        for k in range(NRUN):
            body.append(para("R%02dK%02d" % (ci, k), f, sz, ml))
        body.append(sect(ci == len(COMBOS) - 1))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           '><w:body>' + "".join(body) + "</w:body></w:document>")
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", doc)
    print("wrote", DOCX, len(COMBOS) * NRUN, "paragraphs")


def read():
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    d = app.Documents.Open(DOCX, ReadOnly=True)
    rows = {}
    try:
        d.Repaginate()
        for i in range(1, d.Paragraphs.Count + 1):
            rng = d.Paragraphs(i).Range
            m = re.match(r"(R\d+K\d+)", rng.Text)
            if not m:
                continue
            c = d.Range(rng.Start, rng.Start)
            rows[m.group(1)] = (c.Information(3), round(c.Information(6), 2))
    finally:
        d.Close(False)
        app.Quit()

    print("%-16s %4s %4s %8s %8s %9s %9s %s"
          % ("font", "sz", "mult", "step1", "span", "exact h", "snap h", "verdict"))
    for ci, (f, sz, ml) in enumerate(COMBOS):
        pts = [rows.get("R%02dK%02d" % (ci, k)) for k in range(NRUN)]
        if not pts[0]:
            continue
        pg0 = pts[0][0]
        last = max(k for k in range(NRUN) if pts[k] and pts[k][0] == pg0)
        if last < 4:
            continue
        span = pts[last][1] - pts[0][1]
        step1 = pts[1][1] - pts[0][1] if pts[1] else float("nan")
        # EXACT: span = last*h (+-0.375).  SNAP: span = last*round(h/PX)*PX.
        h_exact = span / last
        h_snap = step1
        pred_snap = last * step1
        verdict = ("SNAP" if abs(span - pred_snap) < 0.4 else "EXACT") \
            if abs(span - pred_snap) > 0.4 or abs(h_exact - h_snap) < 0.02 else "?"
        print("%-16s %4.1f %4d %8.2f %8.2f %9.4f %9.2f  span-vs-snap %+6.2f  %s"
              % (f, sz / 2.0, ml, step1, span, h_exact, h_snap,
                 span - pred_snap, verdict))


if __name__ == "__main__":
    {"gen": gen, "read": read}[sys.argv[1]]()
