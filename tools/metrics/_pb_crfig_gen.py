# -*- coding: utf-8 -*-
"""The creative figure block as it actually is: no list, no autospacing.

creative__0158c02a loses ~1.4pt per figure block (p61 drift table). The
_pb_imgnum probe chased a numPr+beforeAutospacing host — but the real doc has
that form ONCE (p1058 "Amir Siraj"); the other 28 figure hosts are PLAIN
paragraphs.  The real block (Fig2, p139-143) is:

    [empty p, line=276] [host: " "+inline img, line=360] [empty p, line=360]
    [multi-line caption, line=276] [" " p, line=276]

all runs/marks Arial 12pt, docDefaults `after=160 line=259`, pPr overriding
only w:line.  Untested elements vs the arms that already match Word:
S1179's (f-1)*natural term at f=1.5/2.0 (derived @240/276 only), inherited
after=160 on an image host, empty-para height at line=360.

    python _pb_crfig_gen.py gen
    python _pb_crfig_gen.py read     # Word COM truth
    python _pb_crfig_gen.py oxi [ENV=..]
"""
import os
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_crfig")
# "grid" as a trailing CLI arg switches every mode to the no-type
# `<w:docGrid w:linePitch="360"/>` variant (creative's sectPr): Word then
# quantizes the cumulative paragraph cursor to whole 96dpi pixels (0.75pt),
# a DIFFERENT engine from the exact-hhea no-grid path the base arms measure.
GRID = "grid" in sys.argv[2:] or (len(sys.argv) > 1 and sys.argv[1] == "grid")
DOCX = os.path.join(OUT, "crfig_grid.docx" if GRID else "crfig.docx")
GDI = os.path.join(REPO, "tools", "oxi-gdi-renderer", "target", "release",
                   "oxi-gdi-renderer.exe")

sys.path.insert(0, HERE)
from _pb_brimg_gen import PNG, pic  # noqa: E402  (1x1 png + wp:inline builder)

FONT, SZ = "Arial", 24              # the block's 12pt Arial

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
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/img.png"/>'
         '</Relationships>')
# docDefaults mirror the specimen: after=160, line=259 auto.
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + '>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/>'
          '</w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr>'
          '</w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
          '<w:name w:val="Normal"/></w:style></w:styles>')

CAP = ("Figure 2 – WICK ROTATION|: “The complex plane reveals a special "
       "relationship with rotation. Whenever a point on the complex plane is "
       "multiplied by a factor the resulting point is the original rotated a "
       "quarter turn about the origin, and repeating the multiplication walks "
       "the point around the circle.”")  # | = bold/plain split, ~3 lines

# (label, kind, opts) — kind builders below; opts: line, img_h, repeat, after,
#   font, sz (half-points).
#   ("after": explicit w:after twips in pPr, None = inherit docDefaults 160)
ARMS = [
    ("txt@276",       "txt",   {"line": 276}),
    ("empty@276",     "empty", {"line": 276}),
    ("empty@360",     "empty", {"line": 360}),
    ("img36@276",     "img",   {"line": 276, "img_h": 36.0}),
    ("img36@360",     "img",   {"line": 360, "img_h": 36.0}),
    ("img60@360",     "img",   {"line": 360, "img_h": 60.0}),
    ("img36@480",     "img",   {"line": 480, "img_h": 36.0}),
    ("img36@360_a0",  "img",   {"line": 360, "img_h": 36.0, "after": 0}),
    ("cap@276",       "cap",   {"line": 276}),
    ("blk60",         "blk",   {"img_h": 60.0, "repeat": 3}),
    ("blk212",        "blk",   {"img_h": 212.25, "repeat": 1}),
    # ---- empty-vs-text natural sweep (the 13.5-vs-13.83 wall) ----
    ("txtC11@276",    "txt",   {"line": 276, "font": "Calibri", "sz": 22}),
    ("emptyC11@276",  "empty", {"line": 276, "font": "Calibri", "sz": 22}),
    ("txtT12@276",    "txt",   {"line": 276, "font": "Times New Roman", "sz": 24}),
    ("emptyT12@276",  "empty", {"line": 276, "font": "Times New Roman", "sz": 24}),
    ("txtA10@276",    "txt",   {"line": 276, "font": "Arial", "sz": 20}),
    ("emptyA10@276",  "empty", {"line": 276, "font": "Arial", "sz": 20}),
    ("txtA14@360",    "txt",   {"line": 360, "font": "Arial", "sz": 28}),
    ("emptyA14@360",  "empty", {"line": 360, "font": "Arial", "sz": 28}),
    ("txtA12@240",    "txt",   {"line": 240}),
    ("emptyA12@240",  "empty", {"line": 240}),
]
REPEAT_DEFAULT = 8


def rpr(bold=False, font=None, sz=None):
    font = font or FONT
    sz = sz or SZ
    return ('<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>%s'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>'
            % (font, font, font, "<w:b/><w:bCs/>" if bold else "", sz, sz))


def ppr(line, pbb=False, after=None, marker=False, font=None, sz=None):
    # like the specimen: pPr sets only w:line (+ mark rPr); markers pin all.
    if marker:
        sp = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
    elif after is not None:
        sp = ('<w:spacing w:after="%d" w:line="%d" w:lineRule="auto"/>'
              % (after, line))
    else:
        sp = '<w:spacing w:line="%d" w:lineRule="auto"/>' % line
    return ("<w:pPr>%s<w:widowControl w:val=\"0\"/>%s%s</w:pPr>"
            % ("<w:pageBreakBefore/>" if pbb else "", sp, rpr(font=font, sz=sz)))


def para(line, inner, after=None, font=None, sz=None):
    return "<w:p>%s%s</w:p>" % (ppr(line, after=after, font=font, sz=sz), inner)


def txt_p(line, text, after=None, font=None, sz=None):
    return para(line, "<w:r>%s<w:t xml:space=\"preserve\">%s</w:t></w:r>"
                % (rpr(font=font, sz=sz), text), after=after, font=font, sz=sz)


def img_p(line, h, idx, after=None):
    return para(line,
                '<w:r>%s<w:t xml:space="preserve"> </w:t></w:r><w:r>%s%s</w:r>'
                % (rpr(), rpr(), pic(idx, h)), after=after)


def cap_p(line):
    bold, plain = CAP.split("|")
    return para(line,
                '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>'
                '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>'
                % (rpr(bold=True), bold, rpr(), plain))


def subject(kind, idx, opts):
    line = opts.get("line", 276)
    after = opts.get("after")
    font = opts.get("font")
    sz = opts.get("sz")
    if kind == "txt":
        return txt_p(line, "body text line", after=after, font=font, sz=sz)
    if kind == "empty":
        return para(line, "", after=after, font=font, sz=sz)
    if kind == "img":
        return img_p(line, opts["img_h"], idx, after=after)
    if kind == "cap":
        return cap_p(line)
    # blk = the real five-paragraph figure block
    return (para(276, "") + img_p(360, opts["img_h"], idx) + para(360, "")
            + cap_p(276) + txt_p(276, " "))


def marker(tag, pbb=False):
    return ('<w:p>%s<w:r>%s<w:t>%s</w:t></w:r></w:p>'
            % (ppr(240, pbb=pbb, marker=True), rpr(), tag))


def gen():
    os.makedirs(OUT, exist_ok=True)
    body, idx = [], 100
    body.append(marker("M00S", pbb=True))
    body.append(marker("M00E"))
    for ai, (_lbl, kind, opts) in enumerate(ARMS, start=1):
        body.append(marker("M%02dS" % ai, pbb=True))
        for _ in range(opts.get("repeat", REPEAT_DEFAULT)):
            idx += 1
            body.append(subject(kind, idx, opts))
        body.append(marker("M%02dE" % ai))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           "><w:body>" + "".join(body) +
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440" '
           'w:header="708" w:footer="708" w:gutter="0"/>'
           + ('<w:docGrid w:linePitch="360"/>' if GRID else '')
           + '</w:sectPr></w:body></w:document>')
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/media/img.png", PNG)
        z.writestr("word/document.xml", doc)
    print("wrote", DOCX, len(ARMS), "arms")


def report(spans, who):
    base = spans.get(0)
    if base is None:
        raise SystemExit("control arm missing")
    print("%s  control span (one marker) = %.2f" % (who, base))
    print("%-14s %9s" % ("arm", "per copy"))
    for ai, (lbl, _kind, opts) in enumerate(ARMS, start=1):
        span = spans.get(ai)
        n = opts.get("repeat", REPEAT_DEFAULT)
        print("%-14s %9s" % (lbl, "MISSING" if span is None
                             else "%.3f" % ((span - base) / n)))


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
    for ai in range(0, len(ARMS) + 1):
        s, e = ys.get((ai, "S")), ys.get((ai, "E"))
        if s and e and s[0] == e[0]:
            spans[ai] = e[1] - s[1]
    report(spans, "WORD")


def oxi(envs=""):
    import json
    env = dict(os.environ)
    for kv in [s for s in envs.split(",") if s]:
        k, _, v = kv.partition("=")
        env[k] = v or "1"
    out = os.path.join(tempfile.gettempdir(), "crfig_oxi.json")
    subprocess.run([GDI, DOCX, os.path.join(tempfile.gettempdir(), "crfig"),
                    "--dump-layout=" + out], check=True, capture_output=True, env=env)
    ys = {}
    for pg in json.load(open(out, encoding="utf-8"))["pages"]:
        for e in pg["elements"]:
            t = (e.get("text") or "").strip()
            if len(t) == 4 and t.startswith("M") and t[3] in "SE" and t[1:3].isdigit():
                ys.setdefault((int(t[1:3]), t[3]), (pg["page"], e["y"]))
    spans = {}
    for ai in range(0, len(ARMS) + 1):
        s, e = ys.get((ai, "S")), ys.get((ai, "E"))
        if s and e and s[0] == e[0]:
            spans[ai] = e[1] - s[1]
    report(spans, "OXI " + (envs or "(default)"))


if __name__ == "__main__":
    if sys.argv[1] == "oxi":
        oxi(sys.argv[2] if len(sys.argv) > 2 else "")
    else:
        {"gen": gen, "read": read}[sys.argv[1]]()
