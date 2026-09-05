# -*- coding: utf-8 -*-
"""What colour does Word give text WITHOUT a w:color inside a filled wps shape?

Oxi drops such runs when the fill is "dark" (r+g+b < 600) -- a rule derived on
1ec1's 4472C4 heading box, whose visible runs carry an explicit FFFFFF and whose
dropped runs were really overflow lines. reports__167853's thirteen 969696
banners (black text in Word's PDF) fall under the same rule and lose their
titles. Sweep fill x shape-style (a wps:style with a lt1 fontRef is what Word
writes for its own preset shapes) x explicit run colour, and read the glyph
colour from Word's PDF export.

    python _pb_txbxfill_gen.py gen
    python _pb_txbxfill_gen.py pdf      # Word truth (COM -> PDF spans + colours)
"""
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_txbxfill")
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, HERE)
from _pb_pxgrid_gen import CT, NS, RELS  # noqa: E402

# (label, fill hex, wps:style with lt1 fontRef?, explicit run colour or None)
ARMS = [
    ("f969696_plain", "969696", False, None),
    ("f4472C4_plain", "4472C4", False, None),
    ("f000000_plain", "000000", False, None),
    ("fFF0000_plain", "FF0000", False, None),
    ("fFFFFFF_plain", "FFFFFF", False, None),
    ("f969696_style", "969696", True, None),
    ("f4472C4_style", "4472C4", True, None),
    ("f000000_style", "000000", True, None),
    ("f4472C4_red", "4472C4", False, "FF0000"),
    ("f000000_white", "000000", False, "FFFFFF"),
    # HSL-lightness sweep for the auto colour flip (L = (max+min)/2)
    ("f404040_plain", "404040", False, None),   # L 0.25
    ("f7F7F7F_plain", "7F7F7F", False, None),   # L 0.498
    ("f808080_plain", "808080", False, None),   # L 0.502
    ("f000080_plain", "000080", False, None),   # navy, L 0.25, lum 0.06
    ("f00FF00_plain", "00FF00", False, None),   # L 0.5, lum 0.59
    ("f0000FF_plain", "0000FF", False, None),   # L 0.5, lum 0.11
    # grey / dark-colour sweep for the flip threshold (Rec.601 luma of 255)
    ("f101010_plain", "101010", False, None),   # 16
    ("f202020_plain", "202020", False, None),   # 32
    ("f303030_plain", "303030", False, None),   # 48
    ("f383838_plain", "383838", False, None),   # 56
    ("f800000_plain", "800000", False, None),   # 38
    ("f008000_plain", "008000", False, None),   # 75
    ("f004000_plain", "004000", False, None),   # 38
    ("f400000_plain", "400000", False, None),   # 19
    ("f606060_plain", "606060", False, None),   # 96
    ("f3A3A3A_plain", "3A3A3A", False, None),   # 58
    ("f3C3C3C_plain", "3C3C3C", False, None),   # 60
    ("f3E3E3E_plain", "3E3E3E", False, None),   # 62
    ("f3F3F3F_plain", "3F3F3F", False, None),   # 63
    ("f00C000_plain", "00C000", False, None),   # luma 113, L 0.38
    ("fC00000_plain", "C00000", False, None),   # luma 57, L 0.38
    ("fC80000_plain", "C80000", False, None),   # luma 60
]
DNS = ('xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
       'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
       'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
       'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
       'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"')
STYLE = ('<wps:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef>'
         '<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>'
         '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>'
         '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></wps:style>')


def docx(label):
    return os.path.join(OUT, "txbxfill_%s.docx" % label)


def shape(i, fill, styled, run_color, text):
    rpr = '<w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:hint="eastAsia"/>'
    if run_color:
        rpr += '<w:color w:val="%s"/>' % run_color
    rpr += '<w:sz w:val="32"/></w:rPr>'
    return ('<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
            'relativeHeight="%d" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>'
            '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
            '<wp:extent cx="4000000" cy="480000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
            '<wp:docPr id="%d" name="box%d"/><wp:cNvGraphicFramePr/>'
            '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
            '<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="480000"/></a:xfrm>'
            '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
            '<a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>%s'
            '<wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>%s<w:t>%s</w:t></w:r></w:p>'
            '</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="74295" tIns="8890" '
            'rIns="74295" bIns="8890" anchor="t" upright="1"><a:noAutofit/></wps:bodyPr></wps:wsp>'
            '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
            % (251660000 + i, 10 + i, i, fill, STYLE if styled else "", rpr, text))


def gen():
    os.makedirs(OUT, exist_ok=True)
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + NS + ">"
              '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="MS Mincho" w:hAnsi="Century"/>'
              '<w:sz w:val="21"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>')
    for i, (label, fill, styled, run_color) in enumerate(ARMS):
        text = "MIDASHI%d" % i
        body = ("<w:p><w:r><w:t>mae</w:t></w:r></w:p>"
                "<w:p>%s</w:p>" % shape(i, fill, styled, run_color, text)
                + "".join("<w:p><w:r><w:t>gyo%d</w:t></w:r></w:p>" % k for k in range(4)))
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + " " + DNS
               + "><w:body>" + body
               + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                 '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/>'
                 "</w:sectPr></w:body></w:document>")
        with zipfile.ZipFile(docx(label), "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", CT)
            z.writestr("_rels/.rels", RELS)
            z.writestr("word/_rels/document.xml.rels",
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                       '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                       'relationships/styles" Target="styles.xml"/></Relationships>')
            z.writestr("word/styles.xml", styles)
            z.writestr("word/document.xml", doc)
    print("wrote %d arms into %s" % (len(ARMS), OUT))


def pdf():
    import fitz
    import win32com.client as w
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        for label, *_ in ARMS:
            d = app.Documents.Open(docx(label), ReadOnly=True, AddToRecentFiles=False)
            try:
                d.SaveAs2(docx(label)[:-5] + ".word.pdf", 17)
            finally:
                d.Close(False)
    finally:
        app.Quit()
    print("== WORD (PDF): the MIDASHI span's PDF text colour (present at all?) and the drawn fills ==")
    for label, fill, styled, run_color in ARMS:
        pg = fitz.open(docx(label)[:-5] + ".word.pdf")[0]
        found = []
        for b in pg.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for sp in l["spans"]:
                    t = "".join(c["c"] for c in sp["chars"])
                    if "MIDASHI" in t:
                        found.append((t, "%06X" % sp["color"], round(sp["size"], 1)))
        fills = []
        for dr in pg.get_drawings():
            f = dr.get("fill")
            if f:
                fills.append("%02X%02X%02X" % (round(f[0] * 255), round(f[1] * 255), round(f[2] * 255)))
        print("%-16s fill=%s style=%-5s run=%-6s -> text=%s  drawn_fills=%s"
              % (label, fill, styled, run_color, found, sorted(set(fills))[:3]))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "gen"
    if cmd == "pdf":
        pdf()
    else:
        gen()
