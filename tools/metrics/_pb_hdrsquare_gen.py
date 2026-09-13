# -*- coding: utf-8 -*-
"""A header float with SQUARE wrap that spans the column: where does the body start?

correspondence__00595cd73be1bcd0 (blindD50): the first-page header holds a
133pt VML picture (`<w:object>` / v:shape position:absolute margin-top 42.75pt,
width 505pt > the 451pt column, w10:wrap square) and a 62pt DrawingML anchor
(wrapTopAndBottom at 36.5). Word PDF: pictures at 42.75..148.71 and 36.5..98.75;
the first body line (Arial 12, line=336) at 175.5. Oxi: 231.8.

Arms (Calibri 11 body of six lines unless noted, header distance 0, top margin
709 twips = 35.45, A4, margins 1411 = 70.55 so the column is 454pt):
  A_vml_square_wide     v:rect 505pt wide, top 42.75, h 133, wrap square, body line=336 Arial 12
  B_vml_square_wide240  same shape, body line=240 Calibri 11
  C_vml_square_narrow   v:rect 200pt wide (lane stays open)
  D_dml_square_wide     DrawingML wps rect, wrapSquare, same geometry as A
  E_vml_topbottom_wide  v:rect 505pt, wrap topAndBottom
Readout: Word COM Information(6) of body paragraphs 1..3; Oxi dump.
"""
import os, sys, json, subprocess, tempfile, zipfile
from pathlib import Path
GDI = Path(os.environ.get("OXI_GDI_EXE") or "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
OUT = Path('tests/fixtures/hdrsquare'); OUT.mkdir(parents=True, exist_ok=True)
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
      'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
      '</Types>')
ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DOC_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="MS Mincho" w:hAnsi="Calibri" w:cs="Times New Roman"/>'
          '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
          '<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/></w:style>'
          '</w:styles>')

def vml(width, top, wrap):
    return (f'<w:r><w:pict><v:rect id="r1" style="position:absolute;margin-left:-10.55pt;margin-top:{top}pt;width:{width}pt;height:133pt;z-index:251658240;'
            'mso-position-horizontal-relative:text;mso-position-vertical-relative:text" fillcolor="#9999ff" stroked="f">'
            f'<w10:wrap type="{wrap}"/></v:rect></w:pict></w:r>')
def dml(width_pt, top_pt):
    cx = int(width_pt * 12700); cy = 133 * 12700; off = int(top_pt * 12700)
    return ('<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>-133985</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>{off}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/>'
            '<wp:docPr id="1" name="Rect 1"/><wp:cNvGraphicFramePr/>'
            '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
            f'<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="9999FF"/></a:solidFill></wps:spPr>'
            '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')
def header(inner):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {NS}>'
            f'<w:p><w:pPr><w:pStyle w:val="Header"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr></w:pPr>{inner}</w:p></w:hdr>')
def document(line336):
    ppr = '<w:pPr><w:spacing w:line="336" w:lineRule="auto"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr></w:pPr>' if line336 else ''
    rpr = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr>' if line336 else ''
    body = ''.join(f'<w:p>{ppr}<w:r>{rpr}<w:t>Body line {i} of the letter</w:t></w:r></w:p>' for i in range(1, 7))
    sect = ('<w:sectPr><w:headerReference w:type="default" r:id="rId2"/><w:pgSz w:w="11907" w:h="16840"/>'
            '<w:pgMar w:top="709" w:right="1411" w:bottom="1560" w:left="1411" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}{sect}</w:body></w:document>'
arms = {
    "A_vml_square_wide": (vml(505, 42.75, "square"), True),
    "B_vml_square_wide240": (vml(505, 42.75, "square"), False),
    "C_vml_square_narrow": (vml(200, 42.75, "square"), True),
    "D_dml_square_wide": (dml(505, 42.75), True),
    "E_vml_topbottom_wide": (vml(505, 42.75, "topAndBottom"), True),
}
import win32com.client
app = win32com.client.DispatchEx("Word.Application")
try: app.Visible = False
except Exception: pass
try:
  for name, (inner, l336) in arms.items():
    at = OUT / f"{name}.docx"
    with zipfile.ZipFile(at, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS); z.writestr("word/styles.xml", STYLES)
        z.writestr("word/header1.xml", header(inner)); z.writestr("word/document.xml", document(l336))
    d = app.Documents.Open(str(at.resolve()), False, True)
    try:
        ys = []
        for i in range(1, 4):
            r = d.Paragraphs(i).Range; c = d.Range(r.Start, r.Start); ys.append(round(c.Information(6), 2))
        hr = d.Sections(1).Headers(1).Range
        shp = [(round(hr.ShapeRange(i).Top, 2), round(hr.ShapeRange(i).Height, 2), round(hr.ShapeRange(i).Left, 2)) for i in range(1, hr.ShapeRange.Count + 1)]
    finally: d.Close(False)
    with tempfile.TemporaryDirectory() as t:
        dump = Path(t) / 'l.json'
        subprocess.run([str(GDI), str(at), str(Path(t) / 'p'), "96", f"--dump-layout={dump}"], capture_output=True)
        dd = json.loads(dump.read_text(encoding='utf-8')) if dump.exists() else {'pages': [{'elements': []}]}
    oxi = sorted({round(el['y'], 2) for el in dd['pages'][0]['elements'] if el.get('type') == 'text' and 'Body' in el.get('text', '')})
    print(f"=== {name}: WORD body {ys}  shapes {shp}   OXI body {oxi[:3]}")
finally: app.Quit()
