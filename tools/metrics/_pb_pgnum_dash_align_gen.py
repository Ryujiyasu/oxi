# -*- coding: utf-8 -*-
"""Where does a numberInDash PAGE field sit in its footer line?

zh dev correspondence__015abfaf: a left-aligned footer text box holds PAGE with
the section's pgNumType numberInDash; Word draws «- 1 -» from the box's left
edge, Oxi lays out the placeholder «1» and widens it to «- 1 -» around its
centre (layout/mod.rs S534 post-pass), so the leading «-» lands 14pt left of
the box and is clipped.  Arms: footer paragraph jc left / center / right / both,
a centre tab and a right tab before the field, in the body flow and in a text
box; pages 1 and 12 (one and two digits).  Read: the x of each «-» and digit.

    python _pb_pgnum_dash_align_gen.py gen
    python _pb_pgnum_dash_align_gen.py word
    python _pb_pgnum_dash_align_gen.py oxi EXE
"""
import json
import os
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(REPO, "pipeline_data", "_pb_pgnum_dash_align")
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>')
# (label, jc, tab xml, text before the field, in a text box)
ARMS = [
    ("left", "left", "", "", False),
    ("center", "center", "", "", False),
    ("right", "right", "", "", False),
    ("both", "both", "", "", False),
    ("ctab", "left", '<w:tabs><w:tab w:val="center" w:pos="4513"/></w:tabs>', "\t", False),
    ("rtab", "left", '<w:tabs><w:tab w:val="right" w:pos="9026"/></w:tabs>', "\t", False),
    ("box_left", "left", "", "", True),
    ("box_center", "center", "", "", True),
    ("box_right", "right", "", "", True),
]
RPR = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="18"/></w:rPr>'


def field_para(jc, tabs, pre, label):
    t = ('<w:r>%s<w:tab/></w:r>' % RPR) if pre else ""
    return ('<w:p><w:pPr>%s<w:jc w:val="%s"/>%s</w:pPr>%s'
            '<w:r>%s<w:t xml:space="preserve">%s </w:t></w:r>'
            '<w:r>%s<w:fldChar w:fldCharType="begin"/></w:r><w:r>%s<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
            '<w:r>%s<w:fldChar w:fldCharType="separate"/></w:r><w:r>%s<w:t>1</w:t></w:r><w:r>%s<w:fldChar w:fldCharType="end"/></w:r></w:p>'
            % (tabs, jc, RPR, t, RPR, label, RPR, RPR, RPR, RPR, RPR))


def box(inner, idx):
    return ('<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
            '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="%d" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>%d</wp:posOffset></wp:positionH>'
            '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
            '<wp:extent cx="1270000" cy="254000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
            '<wp:docPr id="%d" name="Box %d"/><wp:cNvGraphicFramePr/>'
            '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
            '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="254000"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>'
            '<wps:txbx><w:txbxContent>%s</w:txbxContent></wps:txbx>'
            '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:noAutofit/></wps:bodyPr>'
            '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r></w:p>'
            % (idx + 1, 3000000, idx + 10, idx + 10, inner))


def gen():
    os.makedirs(OUT, exist_ok=True)
    paras = []
    for i, (label, jc, tabs, pre, inbox) in enumerate(ARMS):
        p = field_para(jc, tabs, pre, label.upper())
        if inbox:
            p = box(field_para(jc, tabs, pre, ""), i)
            # the label stays readable in the flow beside the box
            p = p.replace("<w:p><w:r><mc:", '<w:p><w:r>%s<w:t>%s</w:t></w:r><w:r><mc:' % (RPR, label.upper()), 1)
        paras.append(p)
    ftr = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ' + NS + '>' + "".join(paras) + '</w:ftr>')
    body = "".join('<w:p>%s<w:r><w:t>page %d</w:t></w:r></w:p>' % ('<w:r><w:br w:type="page"/></w:r>' if k else "", k + 1)
                   for k in range(12))
    sect = ('<w:sectPr><w:footerReference w:type="default" r:id="rIdF"/><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="4000" w:left="1440" w:header="720" w:footer="400" w:gutter="0"/>'
            '<w:pgNumType w:fmt="numberInDash"/></w:sectPr>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + '><w:body>' + body + sect
           + '</w:body></w:document>')
    p = os.path.join(OUT, "pgnum_dash.docx")
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DRELS)
        z.writestr("word/document.xml", doc)
        z.writestr("word/footer1.xml", ftr)
    print(p)


def rows_from_words(words):
    """[(label, [x of '-' / digit glyph runs])] per footer line, label first."""
    lines = {}
    for x0, y0, x1, y1, t in words:
        lines.setdefault(round(y1, 0), []).append((x0, x1, t))
    out = []
    for y, ws in sorted(lines.items()):
        ws.sort()
        lab = [w for w in ws if w[2].isalpha() and w[2].isupper() or "_" in w[2]]
        nums = [(round(w[0], 2), round(w[1], 2), w[2]) for w in ws if w not in lab]
        if nums:
            out.append((" ".join(w[2] for w in lab), nums))
    return out


def word():
    import fitz
    import win32com.client as w
    src = os.path.join(OUT, "pgnum_dash.docx")
    pdf = os.path.join(OUT, "pgnum_dash_word.pdf")
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    try:
        d = app.Documents.Open(src, ReadOnly=True)
        d.ExportAsFixedFormat(pdf, 17)
        d.Close(False)
    finally:
        app.Quit()
    doc = fitz.open(pdf)
    res = {}
    for pi in (0, 11):
        ws = [(w[0], w[1], w[2], w[3], w[4]) for w in doc[pi].get_text("words") if w[1] > 600]
        res[pi + 1] = rows_from_words(ws)
    json.dump(res, open(os.path.join(OUT, "word.json"), "w", encoding="utf-8"), indent=1)
    for k, v in res.items():
        print("page", k)
        for r in v:
            print("  ", r)


def oxi(exe):
    src = os.path.join(OUT, "pgnum_dash.docx")
    dump = os.path.join(OUT, "oxi.json")
    subprocess.run([exe, src, os.path.join(OUT, "oxi"), "--dump-layout=" + dump], capture_output=True)
    d = json.load(open(dump, encoding="utf-8"))
    for pi in (0, 11):
        print("page", pi + 1)
        for e in d["pages"][pi]["elements"]:
            if e["type"] == "text" and e["y"] > 600:
                print("   %.2f..%.2f %r" % (e["x"], e["x"] + e["w"], e.get("text")))


if __name__ == "__main__":
    {"gen": gen, "word": word}.get(sys.argv[1], lambda: oxi(sys.argv[2]))()
