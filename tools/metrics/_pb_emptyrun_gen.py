# -*- coding: utf-8 -*-
"""Does Word advance an EMPTY paragraph by the exact multiplied height, or a
rounded one?

S1091 gives an empty Latin paragraph the same exact height its text siblings
get (Calibri 11 x1.15 = 15.442); the legacy path advances a flat 15.50.  The
two documents that measured it disagree at the document scale:

  correspondence__000f9471  Word steps average 15.42  -> exact
  ukframework               Word's cursor at the p20 'informing the department'
                            boundary is 746.25, which is the LEGACY arm's
                            746.20; the exact arm is 732.76 (one line high)

A single Info6 advance cannot separate 15.442 from 15.50 (the report is
quantised to 0.75pt), so this probe measures a RUN: 40 consecutive empty
paragraphs bracketed by two text markers.  exact predicts marker span
40*15.442 + h_marker, the flat-0.5 model 40*15.50 + h_marker -- 2.3pt apart,
far outside the +-0.75 of the two endpoints.

  python _pb_emptyrun_gen.py gen
  python _pb_emptyrun_gen.py read
"""
import os
import re
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_pxgrid")
DOCX = os.path.join(OUT, "emptyrun.docx")
NRUN = 40

sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, DRELS, NS, RELS, STYLES  # noqa: E402

# (label, font, half-point size, multiplier)
ARMS = [
    ("cal11x115", "Calibri", 22, 276),
    ("cal11x108", "Calibri", 22, 259),
    ("ari12x115", "Arial", 24, 276),
    ("ari10x100", "Arial", 20, 240),
    ("tnr12x150", "Times New Roman", 24, 360),
]


def rpr(font, sz):
    return ('<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (font, font, font, sz, sz))


def ppr(font, sz, ml):
    return ('<w:pPr><w:spacing w:before="0" w:after="0" w:line="%d" w:lineRule="auto"/>'
            '<w:widowControl w:val="0"/>%s</w:pPr>' % (ml, rpr(font, sz)))


def empty(font, sz, ml):
    return '<w:p>%s</w:p>' % ppr(font, sz, ml)


def marker(tag, font, sz, ml):
    return ('<w:p>%s<w:r>%s<w:t>%s</w:t></w:r></w:p>'
            % (ppr(font, sz, ml), rpr(font, sz), tag))


def sect(last):
    inner = "" if last else '<w:type w:val="nextPage"/>'
    s = ('<w:sectPr>%s<w:pgSz w:w="11906" w:h="16838"/>'
         '<w:pgMar w:top="1440" w:right="1440" w:bottom="360" w:left="1440" '
         'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>' % inner)
    return s if last else '<w:p><w:pPr>%s</w:pPr></w:p>' % s


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (label, font, sz, ml) in enumerate(ARMS):
        body.append(marker("E%02dSTART" % ai, font, sz, ml))
        for _ in range(NRUN):
            body.append(empty(font, sz, ml))
        body.append(marker("E%02dEND" % ai, font, sz, ml))
        body.append(sect(ai == len(ARMS) - 1))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           '><w:body>' + "".join(body) + "</w:body></w:document>")
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", doc)
    print("wrote", DOCX, len(ARMS), "arms x", NRUN, "empties")


def natural():
    from fontTools.ttLib import TTFont
    files = {"Arial": "arial.ttf", "Calibri": "calibri.ttf",
             "Times New Roman": "times.ttf"}
    out = {}
    for name, fn in files.items():
        f = TTFont(os.path.join(os.environ["WINDIR"], "Fonts", fn),
                   fontNumber=0, lazy=True)
        up = f["head"].unitsPerEm
        hh = f["hhea"]
        out[name] = (hh.ascender - hh.descender + hh.lineGap) / up
    return out


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
            m = re.match(r"(E\d+(?:START|END))", rng.Text)
            if not m:
                continue
            c = d.Range(rng.Start, rng.Start)
            rows[m.group(1)] = (c.Information(3), round(c.Information(6), 2))
    finally:
        d.Close(False)
        app.Quit()
    nat = natural()
    print("%-10s %-16s %5s %5s %9s %9s %9s %9s %s"
          % ("arm", "font", "sz", "mult", "span", "per empty", "exact", "ceil .5", "verdict"))
    for ai, (label, font, sz, ml) in enumerate(ARMS):
        a, b = rows.get("E%02dSTART" % ai), rows.get("E%02dEND" % ai)
        if not a or not b or a[0] != b[0]:
            print("%-10s (crosses a page or missing)" % label)
            continue
        h = nat[font] * (sz / 2.0) * (ml / 240.0)
        per = (b[1] - a[1]) / (NRUN + 1)          # NRUN empties + the START marker
        ceil5 = (int(h / 0.5) + (0 if abs(h / 0.5 - round(h / 0.5)) < 1e-9 else 1)) * 0.5
        v = "EXACT" if abs(per - h) < abs(per - ceil5) else "CEIL0.5"
        print("%-10s %-16s %5.1f %5d %9.2f %9.4f %9.4f %9.4f %s"
              % (label, font, sz / 2.0, ml, b[1] - a[1], per, h, ceil5, v))


if __name__ == "__main__":
    {"gen": gen, "read": read}[sys.argv[1]]()
