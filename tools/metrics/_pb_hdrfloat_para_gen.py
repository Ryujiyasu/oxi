# -*- coding: utf-8 -*-
"""A header's PARAGRAPH-relative wrapTopAndBottom float: where does the body start?

reports__003862302b660a86 (blindD50): the first-page header holds one empty
paragraph whose anchored picture (positionV relativeFrom="paragraph"
posOffset=-35.45pt, cy=100.5pt, wrapTopAndBottom) spans 0..100.5 on the page,
and the section's top margin is NEGATIVE (-1438 twips = fixed 71.9). Word COM
starts the body at 100.5 -- the float band, not the fixed margin. S1010 prices
only PAGE-relative header bands; S1267 pins a negative margin.

Arms (Calibri 11 body, header distance 709 twips = 35.45pt, a 100.5pt-tall
wps shape with wrapTopAndBottom anchored in the header's only paragraph):
  A_neg_margin_para   top=-1438, positionV paragraph -35.45pt      -> 100.5 ?
  B_pos_margin_para   top=1438,  same anchor                        -> band + host line ?
  C_neg_margin_noimg  top=-1438, no anchor                          -> 71.9 (S1267 control)
  D_pos_margin_noimg  top=1438,  no anchor, empty header            -> 71.9 (S843 control)
  E_neg_margin_page   top=-1438, positionV page 0                   -> 100.5 (S1010 shape)
Readout: COM Information(6) of the first body paragraph; PDF first line top.
"""
import os, sys, json, subprocess, tempfile, zipfile
from pathlib import Path
GDI = Path(os.environ.get("OXI_GDI_EXE") or "tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
OUT = Path('tests/fixtures/hdrfloat_para'); OUT.mkdir(parents=True, exist_ok=True)
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"')
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

def anchor(v_rel, v_off_emu):
    return ('<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="{v_rel}"><wp:posOffset>{v_off_emu}</wp:posOffset></wp:positionV>'
            '<wp:extent cx="3600000" cy="1276350"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapTopAndBottom/>'
            '<wp:docPr id="1" name="Rect 1"/><wp:cNvGraphicFramePr/>'
            '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
            '<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3600000" cy="1276350"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="9999FF"/></a:solidFill></wps:spPr>'
            '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')

def header(anch):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {NS}>'
            f'<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr>{anch}</w:p></w:hdr>')

def document(top):
    body = ''.join(f'<w:p><w:r><w:t>Body line {i}</w:t></w:r></w:p>' for i in range(1, 5))
    sect = (f'<w:sectPr><w:headerReference w:type="default" r:id="rId2"/><w:pgSz w:w="11906" w:h="16838"/>'
            f'<w:pgMar w:top="{top}" w:right="1797" w:bottom="1440" w:left="1797" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}{sect}</w:body></w:document>'

arms = {
    "A_neg_margin_para": (-1438, anchor("paragraph", -450215)),
    "B_pos_margin_para": (1438, anchor("paragraph", -450215)),
    "C_neg_margin_noimg": (-1438, ''),
    "D_pos_margin_noimg": (1438, ''),
    "E_neg_margin_page": (-1438, anchor("page", 0)),
}
import pymupdf, win32com.client
app = win32com.client.DispatchEx("Word.Application")
try: app.Visible = False
except Exception: pass
try:
  for name, (top, anch) in arms.items():
    at = OUT / f"{name}.docx"
    with zipfile.ZipFile(at, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS); z.writestr("word/styles.xml", STYLES)
        z.writestr("word/header1.xml", header(anch)); z.writestr("word/document.xml", document(top))
    pdf = str(at)[:-5] + '.pdf'
    d = app.Documents.Open(str(at.resolve()), False, True)
    try:
        d.ExportAsFixedFormat(OutputFileName=str(Path(pdf).resolve()), ExportFormat=17)
        r = d.Paragraphs(1).Range; c = d.Range(r.Start, r.Start)
        body_y = round(c.Information(6), 2)
    finally: d.Close(False)
    pg = pymupdf.open(pdf)[0]
    lines = sorted((round(l['bbox'][1], 2), ''.join(s['text'] for s in l['spans'])[:12]) for b in pg.get_text('dict')['blocks'] for l in b.get('lines', []))
    with tempfile.TemporaryDirectory() as t:
        dump = Path(t) / 'l.json'
        subprocess.run([str(GDI), str(at), str(Path(t) / 'p'), "96", f"--dump-layout={dump}"], capture_output=True)
        dd = json.loads(dump.read_text(encoding='utf-8')) if dump.exists() else {'pages': [{'elements': []}]}
    oxi = sorted({round(el['y'], 2) for el in dd['pages'][0]['elements'] if el.get('type') == 'text' and 'Body' in el.get('text', '')})
    print(f"=== {name}: WORD body Info6 {body_y}  PDF lines {lines[:3]}  OXI body {oxi[:2]}")
finally: app.Quit()
