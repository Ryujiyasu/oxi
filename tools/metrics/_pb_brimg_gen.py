# -*- coding: utf-8 -*-
"""How tall is a paragraph whose run holds a <w:br/> and an INLINE image?

reference__0042471c's p28 is exactly that -- a run with one <w:br/> and one
inline picture 130.7 x 24.2pt, no text -- and it is where Oxi's page-1 cursor
gains the ~13.5pt that eventually costs the doc a page-boundary flip:

    Word  640.50 -> 678.75 = 38.25   (a 13.5pt line, then a 24.75pt line)
    Oxi   645.03 -> 696.83 = 51.80   (13.80 + 13.80, then the picture on its own)

i.e. Word puts the picture ON the line the break started, Oxi emits an empty
text line and then the picture.  This probe derives the rule instead of
patching the specimen: it sweeps the picture height across five run shapes and
reads Word's own advance for each.

Each arm is one page (pageBreakBefore on its start marker) holding REPEAT
copies of the subject paragraph between two markers, so the per-copy height is
(span - span_of_the_empty_control) / REPEAT and the +-0.375pt Info6 reporting
noise divides down with it.

  python _pb_brimg_gen.py gen
  python _pb_brimg_gen.py read              # Word COM truth
  python _pb_brimg_gen.py oxi               # Oxi's own advance, same arms
"""
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_brimg")
DOCX = os.path.join(OUT, "brimg.docx")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_emptyrun_gen import natural  # noqa: E402

FONT, SZ = "Calibri", 22               # 11pt -> natural 13.4277
REPEAT = 8
EMU = 12700
PNG = (b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
       b"\x08\x06\x00\x00\x00\x1f\x15\xc4\x89\x00\x00\x00\nIDATx\x9cc\x00\x01"
       b"\x00\x00\x05\x00\x01\r\n-\xb4\x00\x00\x00\x00IEND\xaeB`\x82")

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Default Extension="png" ContentType="image/png"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/img.png"/></Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + '>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/>'
          '</w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
          '</w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
          '<w:name w:val="Normal"/></w:style></w:styles>')

# (label, runs) -- each run is the list of items inside one <w:r>.
# The "R:" shapes put every item in its OWN run, which is how the specimen is
# written (`<w:r><w:br/></w:r><w:r><w:drawing/></w:r>`); the first five pack
# them into a single run, which is legal OOXML and turned out to be handled
# differently.
SHAPES = [
    ("br_img", [["br", "img"]]),
    ("img", [["img"]]),
    ("txt_br_img", [["txt", "br", "img"]]),
    ("img_br_txt", [["img", "br", "txt"]]),
    ("txt_img", [["txt", "img"]]),
    ("R:br|img", [["br"], ["img"]]),
    ("R:txt|img", [["txt"], ["img"]]),
    ("R:txt|br|img", [["txt"], ["br"], ["img"]]),
    ("R:img|br|txt", [["img"], ["br"], ["txt"]]),
]
HEIGHTS = [6.0, 12.0, 18.0, 24.2, 36.0, 60.0]     # picture cy, pt


def arms():
    return [(s, h) for s in SHAPES for h in HEIGHTS]


def rpr():
    return ('<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (FONT, FONT, FONT, SZ, SZ))


def ppr(pbb=False):
    return ("<w:pPr>%s<w:widowControl w:val=\"0\"/>"
            "<w:spacing w:before=\"0\" w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/>%s</w:pPr>"
            % ("<w:pageBreakBefore/>" if pbb else "", rpr()))


def pic(idx, h_pt):
    cx, cy = int(round(60 * EMU)), int(round(h_pt * EMU))
    a = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    return ('<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
            '<wp:extent cx="%d" cy="%d"/><wp:docPr id="%d" name="P%d"/>'
            '<a:graphic %s><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:nvPicPr><pic:cNvPr id="%d" name="img.png"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>'
            % (cx, cy, idx, idx, a, idx, cx, cy))


def subject(runs, h_pt, idx):
    item = {"txt": lambda: '<w:t xml:space="preserve">x</w:t>',
            "br": lambda: "<w:br/>",
            "img": lambda: pic(idx, h_pt)}
    out = []
    for run in runs:
        out.append("<w:r>%s%s</w:r>" % (rpr(), "".join(item[k]() for k in run)))
    return "<w:p>%s%s</w:p>" % (ppr(), "".join(out))


def marker(tag, pbb=False):
    return ('<w:p>%s<w:r>%s<w:t>%s</w:t></w:r></w:p>' % (ppr(pbb), rpr(), tag))


def gen():
    os.makedirs(OUT, exist_ok=True)
    body, idx = [], 100
    # arm 0 is the control: markers with nothing between them
    body.append(marker("M00S", pbb=True))
    body.append(marker("M00E"))
    for ai, (shape, h) in enumerate(arms(), start=1):
        body.append(marker("M%02dS" % ai, pbb=True))
        for _ in range(REPEAT):
            idx += 1
            body.append(subject(shape[1], h, idx))
        body.append(marker("M%02dE" % ai))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440" '
           'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>')
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/media/img.png", PNG)
        z.writestr("word/document.xml", doc)
    print("wrote", DOCX, len(arms()), "arms x", REPEAT, "copies")


def report(spans, who):
    nat = natural()[FONT] * (SZ / 2.0)
    base = spans.get(0)
    if base is None:
        raise SystemExit("control arm missing")
    print("%s   natural line = %.4f   control span (one marker) = %.2f"
          % (who, nat, base))
    print("%-12s %6s %9s %9s %9s  %s"
          % ("shape", "img_h", "per copy", "1line+img", "2 lines", "verdict"))
    for ai, (shape, h) in enumerate(arms(), start=1):
        span = spans.get(ai)
        if span is None:
            print("%-12s %6.1f   MISSING" % (shape[0], h))
            continue
        per = (span - base) / REPEAT
        one = max(nat, h)                     # picture sits on one line
        two = nat + max(nat, h)               # a text line, then the picture
        pick = "1line+img" if abs(per - one) < abs(per - two) else "2 lines"
        print("%-12s %6.1f %9.3f %9.3f %9.3f  %s (%+.2f)"
              % (shape[0], h, per, one, two, pick,
                 per - (one if pick == "1line+img" else two)))


def read():
    import re
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.ScreenUpdating = False
    d = app.Documents.Open(DOCX, ReadOnly=True)
    ys = {}
    try:
        d.Repaginate()
        for p in d.Paragraphs:
            rng = p.Range
            m = re.match(r"M(\d\d)([SE])", rng.Text)
            if not m:
                continue
            c = d.Range(rng.Start, rng.Start)
            ys[(int(m.group(1)), m.group(2))] = (c.Information(3), round(c.Information(6), 2))
    finally:
        d.Close(False)
        app.Quit()
    spans = {}
    for ai in range(0, len(arms()) + 1):
        s, e = ys.get((ai, "S")), ys.get((ai, "E"))
        if s and e and s[0] == e[0]:
            spans[ai] = e[1] - s[1]
    report(spans, "WORD")


def oxi(envs=""):
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    import json
    out = os.path.join(tempfile.gettempdir(), "brimg_oxi.json")
    subprocess.run([GDI, DOCX, os.path.join(tempfile.gettempdir(), "brimg"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    ys = {}
    for pg in json.load(open(out, encoding="utf-8"))["pages"]:
        for e in pg["elements"]:
            t = (e.get("text") or "").strip()
            if len(t) == 4 and t.startswith("M") and t[3] in "SE" and t[1:3].isdigit():
                ys.setdefault((int(t[1:3]), t[3]), (pg["page"], e["y"]))
    spans = {}
    for ai in range(0, len(arms()) + 1):
        s, e = ys.get((ai, "S")), ys.get((ai, "E"))
        if s and e and s[0] == e[0]:
            spans[ai] = e[1] - s[1]
    report(spans, "OXI  " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        {"gen": gen, "read": read}[sys.argv[1]]()
