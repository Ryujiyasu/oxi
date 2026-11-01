# -*- coding: utf-8 -*-
"""Does a run of default tabs wrap when the next stop lies past the right margin,
and how does a tab cross a wrapSquare picture?

en/educational__0071453c p1: a paragraph of ten TNR-10 tabs beside a wrapTight
logo (column x 240.3, 128.25 wide, dist 9). Word COM Information(5): tabs 1-6 at
57..237, tab 7 jumps to 433.5 (right of the logo), 8-9 at 453/489, tab 10 wraps to
a second line at 57 -- the paragraph is two lines (26.25), Oxi one (14.6), and the
whole page sits 11.65 high.

Arms: N tabs then «X» (no picture) for N in 12..16 (text width 481.9, default stop
36); and the same with a 128x42 wrapSquare picture at column x 240. Readout from
Word's PDF: «X» x and line, and the next paragraph's baseline.
"""
import os, sys, zipfile, io
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/tab_float_cross'; os.makedirs(OUT, exist_ok=True)
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Default Extension="png" ContentType="image/png"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/i.png"/>'
         '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')


def settings(c):
    cs = '' if c is None else (f'<w:compat><w:compatSetting w:name="compatibilityMode" '
                               f'w:uri="http://schemas.microsoft.com/office/word" w:val="{c}"/></w:compat>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {NS}>'
            f'<w:defaultTabStop w:val="720"/>{cs}</w:settings>')
RPR = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/></w:rPr>'


def png():
    from PIL import Image
    b = io.BytesIO(); Image.new('RGB', (40, 40), (0, 0, 0)).save(b, 'PNG'); return b.getvalue()


def pic():
    return ('<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="1" behindDoc="0" '
            'locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>3048000</wp:posOffset></wp:positionH>'
            '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1625600" cy="533400"/>'
            '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="p"/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="i.png"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="rId5"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1625600" cy="533400"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')


def document(n, with_pic):
    tabs = ''.join(f'<w:r>{RPR}<w:tab/></w:r>' for _ in range(n))
    body = (f'<w:p>{pic() if with_pic else ""}{tabs}<w:r>{RPR}<w:t>X</w:t></w:r></w:p>'
            f'<w:p><w:r>{RPR}<w:t>NEXTPARA</w:t></w:r></w:p>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1140" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


if __name__ == '__main__':
    import win32com.client, pymupdf
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        print('text 57.0 .. 541.3 ; default stops every 36 from 57; picture keep-out 288.0 .. 434.2')
        arms = [(None, False, n) for n in (13, 14)] + [(c, True, n) for c in (None, 14, 15) for n in (6, 7, 8, 9, 10)]
        for c, with_pic, n in arms:
            name = f'c{c or "x"}_{"pic" if with_pic else "nopic"}_t{n}'
            at = os.path.abspath(os.path.join(OUT, name + '.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/media/i.png', png())
                z.writestr('word/settings.xml', settings(c))
                z.writestr('word/document.xml', document(n, with_pic))
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                r = d.Paragraphs(1).Range
                tabs = [round(d.Range(i, i).Information(5), 1) for i in range(r.Start, r.End) if d.Range(i, i + 1).Text == '\t']
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            P = pymupdf.open(pdf)[0]
            ws = P.get_text('words')
            x = [(round(w[0], 2), round(w[3], 2)) for w in ws if w[4] == 'X']
            nx = [round(w[3], 2) for w in ws if w[4] == 'NEXTPARA']
            print(f'{name:16} tabs {tabs}  X {x}  next {nx}', flush=True)
    finally:
        app.Quit()
