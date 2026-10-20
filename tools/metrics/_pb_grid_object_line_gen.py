# -*- coding: utf-8 -*-
"""How tall is a line holding only an inline object on a typed `lines` grid?

golden probeqobject: a 144x72pt VML w:object alone in a paragraph, grid pitch 18,
Normal.dotm line spacing (multiple 1.158): Word's object line is ~75pt, Oxi rounds it
up to 5 cells (90) per S1418 (technical__90f5b9d4: a 244pt picture line spans 14 cells).

Arms: object kind (VML w:object / DrawingML inline picture) x height (72 / 80 / 90 / 100)
x line spacing (auto 240 / auto 278 / none = inherited) x grid (lines 18 / none).
Per arm: Info(6) of the paragraph after the object minus Info(6) of the object paragraph
(spacing before/after 0).

    python tools/metrics/_pb_grid_object_line_gen.py
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")
REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_grid_object_line")
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:v="urn:schemas-microsoft-com:vml" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Default Extension="png" ContentType="image/png"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/></w:pPr></w:style></w:styles>')
# the golden probe's own 74-byte PNG (a hand-made 1x1 one made Word ignore the VML size)
PNG = open(os.path.join(REPO, "pipeline_data", "_pb_grid_object_line_img.png"), "rb").read() if os.path.exists(
    os.path.join(REPO, "pipeline_data", "_pb_grid_object_line_img.png")) else b""
TEXT = "本項に定める事項については、関係法令及び本規程の趣旨に照らし。"


def ppr(sp):
    s = '<w:spacing w:before="0" w:after="0"%s/>' % ('' if sp is None else ' w:line="%d" w:lineRule="auto"' % sp)
    return '<w:pPr>%s</w:pPr>' % s


def obj(kind, h):
    if kind == "vml":
        return ('<w:r><w:object w:dxaOrig="2880" w:dyaOrig="%d"><v:shape id="O1" style="width:144pt;height:%dpt">'
                '<v:imagedata r:id="rId20"/></v:shape></w:object></w:r>' % (h * 20, h))
    cx, cy = 144 * 12700, h * 12700
    return ('<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="%d" cy="%d"/>'
            '<wp:docPr id="1" name="P1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="P1"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="rId20"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>' % (cx, cy, cx, cy))


def build(kind, h, sp, grid):
    p = lambda inner: '<w:p>%s%s</w:p>' % (ppr(sp), inner)
    t = lambda s: '<w:r><w:t>%s</w:t></w:r>' % s
    body = p(t("A" + TEXT)) + p(t("B" + TEXT)) + p(obj(kind, h)) + p(t("C" + TEXT)) + p(t("D" + TEXT))
    g = '<w:docGrid w:type="lines" w:linePitch="360"/>' if grid else ''
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418"/>%s</w:sectPr></w:body></w:document>'
           % (NS, body, g))
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, "go_%s_h%d_s%s_g%d.docx" % (kind, h, sp, grid))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS); z.writestr("word/styles.xml", STYLES)
        z.writestr("word/media/image1.png", PNG); z.writestr("word/document.xml", doc)
    return path


if __name__ == "__main__":
    import win32com.client
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
    try:
        if os.environ.get("PB_SWEEP"):
            for kind in ("dml", "vml"):
                for h in range(60, 112, 2):
                    d = w.Documents.Open(build(kind, h, 240, 1), ReadOnly=True)
                    y = [d.Range(d.Paragraphs(i).Range.Start, d.Paragraphs(i).Range.Start).Information(6) for i in (3, 4)]
                    d.Close(False)
                    print("%s h%d obj %.2f extra %.2f" % (kind, h, y[1] - y[0], y[1] - y[0] - h), flush=True)
            sys.exit(0)
        for grid in (1, 0):
            for sp in (None, 240, 278, 360):
                for kind in ("vml", "dml"):
                    row = []
                    for h in (72, 80, 90, 100):
                        d = w.Documents.Open(build(kind, h, sp, grid), ReadOnly=True)
                        y = [d.Range(d.Paragraphs(i).Range.Start, d.Paragraphs(i).Range.Start).Information(6) for i in (1, 2, 3, 4, 5)]
                        d.Close(False)
                        row.append("h%d: text %.2f obj %.2f" % (h, y[1] - y[0], y[3] - y[2]))
                    print("grid %d sp %-4s %s  %s" % (grid, sp, kind, " | ".join(row)), flush=True)
    finally:
        w.Quit()
