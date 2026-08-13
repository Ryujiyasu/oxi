# -*- coding: utf-8 -*-
"""What does a numbering SYMBOL run contribute to a line's height?

policies__000f7115 page 26 is 13 single-line Symbol-bulleted Arial 12 items at
line=276.  Word's own span over them gives 16.6875pt per item (+-0.06), while

  Arial 12 x1.15 alone              15.8687   (too short)
  Symbol asc + Arial desc, x1.15    16.7995   (what Oxi computes today)
  Symbol hhea asc+desc, x1.15       16.9063   (the whole symbol box)

so neither the plain text line nor the full symbol box is what Word uses, and
Oxi's 0.11pt/line excess is exactly what cancels its 0.37pt/paragraph empty
under-height there -- the compensating balance that blocks S1091.

Each arm is 20 single-line list paragraphs in its own section, so the span
between the first and last pins the per-item height to +-0.04pt (Info6 is
quantised to the 96dpi pixel; see _pb_pxrun_gen.py).

  python _pb_bullet_gen.py gen
  python _pb_bullet_gen.py read
"""
import os
import re
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_pxgrid")
DOCX = os.path.join(OUT, "bullet.docx")
PX = 0.75
NRUN = 20

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + '>'
          '<w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="24"/>'
          '</w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
          '</w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
          '<w:name w:val="Normal"/></w:style></w:styles>')

# (label, bullet font, bullet char, bullet half-point size or None = inherit,
#  body half-point size, multiplier)
ARMS = [
    ("sym12",      "Symbol",      "", None, 24, 276),
    ("sym12x1",    "Symbol",      "", None, 24, 240),
    ("sym12s20",   "Symbol",      "", 20,   24, 276),
    ("sym12s28",   "Symbol",      "", 28,   24, 276),
    ("sym10",      "Symbol",      "", None, 20, 276),
    ("wing12",     "Wingdings",   "", None, 24, 276),
    ("cour12",     "Courier New", "o",      None, 24, 276),
    ("none12",     None,          None,     None, 24, 276),   # plain paragraphs
]

# Is the ascent-overflow shape specific to the numbering symbol, or is it the
# general mixed-run line-height rule?  Each of these is a PLAIN paragraph whose
# body run is Arial `body` with one extra run in (font, size) appended.
#   overflow model  h = nat(body)*mult + max(0, asc(extra) - asc(body) - gap(body))
#   classic model   h = (max asc + max desc) * mult
RUN_ARMS = [
    ("runsym12", "Symbol", 24, 24, 276),
    ("runa14",   "Arial", 28, 24, 276),
    ("runa20",   "Arial", 40, 24, 276),
    ("runtnr12", "Times New Roman", 24, 24, 276),
    ("runa14x1", "Arial", 28, 24, 240),
]


def numbering():
    lvls = []
    for i, (label, font, ch, bsz, _body, _ml) in enumerate(ARMS):
        if font is None:
            fmt, txt = "none", ""
            rpr = ""
        else:
            fmt, txt = "bullet", ch
            rpr = '<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:hint="default"/>%s</w:rPr>' % (
                font, font, ("<w:sz w:val=\"%d\"/>" % bsz) if bsz else "")
        lvls.append(
            '<w:abstractNum w:abstractNumId="%d"><w:multiLevelType w:val="singleLevel"/>'
            '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="%s"/>'
            '<w:lvlText w:val="%s"/><w:lvlJc w:val="left"/>'
            '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>%s</w:lvl></w:abstractNum>'
            % (i, fmt, txt, rpr))
    nums = "".join('<w:num w:numId="%d"><w:abstractNumId w:val="%d"/></w:num>' % (i + 1, i)
                   for i in range(len(ARMS)))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ' + NS + '>'
            + "".join(lvls) + nums + '</w:numbering>')


def para(tag, num_id, body_sz, ml):
    rpr = ('<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
           '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (body_sz, body_sz))
    return ('<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="%d"/></w:numPr>'
            '<w:spacing w:before="0" w:after="0" w:line="%d" w:lineRule="auto"/>'
            '<w:widowControl w:val="0"/>%s</w:pPr><w:r>%s<w:t>%s</w:t></w:r></w:p>'
            % (num_id, ml, rpr, rpr, tag))


def sect(last):
    inner = "" if last else '<w:type w:val="nextPage"/>'
    s = ('<w:sectPr>%s<w:pgSz w:w="11906" w:h="16838"/>'
         '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
         'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>' % inner)
    return s if last else '<w:p><w:pPr>%s</w:pPr></w:p>' % s


def run_para(tag, extra_font, extra_sz, body_sz, ml):
    def rpr(font, sz):
        return ('<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>'
                '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (font, font, font, sz, sz))
    return ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="%d" w:lineRule="auto"/>'
            '<w:widowControl w:val="0"/>%s</w:pPr>'
            '<w:r>%s<w:t xml:space="preserve">%s </w:t></w:r>'
            '<w:r>%s<w:t>Z</w:t></w:r></w:p>'
            % (ml, rpr("Arial", body_sz), rpr("Arial", body_sz), tag,
               rpr(extra_font, extra_sz)))


def gen():
    os.makedirs(OUT, exist_ok=True)
    body = []
    for ai, (label, _f, _c, _bs, body_sz, ml) in enumerate(ARMS):
        for k in range(NRUN):
            body.append(para("B%02dK%02d" % (ai, k), ai + 1, body_sz, ml))
        body.append(sect(False))
    for ri, (label, font, esz, body_sz, ml) in enumerate(RUN_ARMS):
        for k in range(NRUN):
            body.append(run_para("R%02dK%02d" % (ri, k), font, esz, body_sz, ml))
        body.append(sect(ri == len(RUN_ARMS) - 1))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS +
           '><w:body>' + "".join(body) + "</w:body></w:document>")
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/numbering.xml", numbering())
        z.writestr("word/document.xml", doc)
    print("wrote", DOCX, len(ARMS) * NRUN, "paragraphs")


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
            m = re.match(r"([BR]\d+K\d+)", rng.Text)
            if not m:
                continue
            c = d.Range(rng.Start, rng.Start)
            rows[m.group(1)] = (c.Information(3), round(c.Information(6), 2))
    finally:
        d.Close(False)
        app.Quit()

    metrics = font_metrics()

    def measured(prefix, ai):
        pts = [rows.get("%s%02dK%02d" % (prefix, ai, k)) for k in range(NRUN)]
        if not pts[0]:
            return None
        pg0 = pts[0][0]
        last = max(k for k in range(NRUN) if pts[k] and pts[k][0] == pg0)
        if last < 4:
            return None
        return (pts[last][1] - pts[0][1]) / last, 0.75 / last

    print("%-10s %-16s %5s %5s %5s %9s %9s %9s %s"
          % ("arm", "extra font", "esz", "body", "mult", "measured", "overflow",
             "classic", "verdict"))
    for ai, (label, font, _c, bsz, body_sz, ml) in enumerate(ARMS):
        got = measured("B", ai)
        if not got:
            continue
        h, tol = got
        esz = bsz or body_sz
        ov, cl = models(metrics, font, esz, body_sz, ml)
        print("%-10s %-16s %5.1f %5.1f %5d %9.4f %9.4f %9.4f %s"
              % (label, font or "-", esz / 2.0, body_sz / 2.0, ml, h, ov, cl,
                 verdict(h, ov, cl, tol)))
    for ri, (label, font, esz, body_sz, ml) in enumerate(RUN_ARMS):
        got = measured("R", ri)
        if not got:
            continue
        h, tol = got
        ov, cl = models(metrics, font, esz, body_sz, ml)
        print("%-10s %-16s %5.1f %5.1f %5d %9.4f %9.4f %9.4f %s"
              % (label, font, esz / 2.0, body_sz / 2.0, ml, h, ov, cl,
                 verdict(h, ov, cl, tol)))


def verdict(h, ov, cl, tol):
    tol = max(tol, 0.05)
    a, b = abs(h - ov) <= tol, abs(h - cl) <= tol
    if a and b:
        return "both"
    if a:
        return "OVERFLOW"
    if b:
        return "classic"
    return "NEITHER"


def font_metrics():
    import os as _os
    from fontTools.ttLib import TTFont
    files = {"Arial": "arial.ttf", "Symbol": "symbol.ttf",
             "Wingdings": "wingding.ttf", "Courier New": "cour.ttf",
             "Times New Roman": "times.ttf"}
    out = {}
    for name, fn in files.items():
        f = TTFont(_os.path.join(_os.environ["WINDIR"], "Fonts", fn),
                   fontNumber=0, lazy=True)
        up = f["head"].unitsPerEm
        hh = f["hhea"]
        out[name] = (hh.ascender / up, -hh.descender / up, hh.lineGap / up)
    return out


def models(m, extra_font, extra_hsz, body_hsz, ml):
    """(ascent-overflow height, classic max-asc+max-desc height)."""
    mult = ml / 240.0
    ba, bd, bg = m["Arial"]
    bsz = body_hsz / 2.0
    nat = (ba + bd + bg) * bsz
    if not extra_font:
        return nat * mult, nat * mult
    ea, ed, eg = m[extra_font]
    esz = extra_hsz / 2.0
    overflow = nat * mult + max(0.0, ea * esz - (ba + bg) * bsz)
    classic = (max(ba * bsz, ea * esz) + max(bd * bsz, ed * esz)
               + max(bg * bsz, eg * esz)) * mult
    return overflow, classic


if __name__ == "__main__":
    {"gen": gen, "read": read}[sys.argv[1]]()
