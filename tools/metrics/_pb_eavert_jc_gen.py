# -*- coding: utf-8 -*-
"""Where does a jc=center / right paragraph start in an eaVert text box?

ja/reports__67cf8704: the eaVert label text boxes centre their text along the
column in Word (the first character ~20pt lower than Oxi's top-aligned column).
Arms: jc left / center / right / both x one short line («あいうえお») and a line
that wraps (30 chars), box 60 x 200pt, insets default. Readout: Word PDF y range
of the characters.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/eavert_jc'; os.makedirs(OUT, exist_ok=True)
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
TEXTS = {'short': 'あいうえお', 'long': 'あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほ'}


def document(jc, text):
    rpr = '<w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="22"/></w:rPr>'
    tb = (f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
          '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>1270000</wp:posOffset></wp:positionH>'
          '<wp:positionV relativeFrom="margin"><wp:posOffset>635000</wp:posOffset></wp:positionV><wp:extent cx="762000" cy="2540000"/>'
          '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="1" name="tb"/>'
          '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>'
          '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="762000" cy="2540000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
          '<a:noFill/><a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>'
          f'<wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="{jc}"/></w:pPr><w:r>{rpr}<w:t>{text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>'
          '<wps:bodyPr rot="0" vert="eaVert" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
          '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body><w:p>{tb}</w:p>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


if __name__ == '__main__':
    import win32com.client, pymupdf
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        print('box top 122 .. bottom 322 (200pt), x 172..232')
        for tname, text in TEXTS.items():
            for jc in ('left', 'center', 'right', 'both'):
                name = f'{tname}_{jc}'
                at = os.path.abspath(os.path.join(OUT, name + '.docx'))
                with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                    z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                    z.writestr('word/document.xml', document(jc, text))
                pdf = at[:-5] + '.pdf'
                d = app.Documents.Open(at, ReadOnly=True)
                try:
                    d.ExportAsFixedFormat(pdf, 17)
                finally:
                    d.Close(False)
                P = pymupdf.open(pdf)[0]
                cols = {}
                for b in P.get_text('rawdict')['blocks']:
                    for l in b.get('lines', []):
                        for s in l['spans']:
                            for c in s['chars']:
                                if c['c'].strip():
                                    cols.setdefault(round(c['bbox'][0]), []).append((c['bbox'][1], c['bbox'][3]))
                desc = '  '.join(f'x{x}: y {min(v[0] for v in ys):.2f}..{max(v[1] for v in ys):.2f} n{len(ys)}' for x, ys in sorted(cols.items(), reverse=True))
                print(f'{name:14} {desc}', flush=True)
    finally:
        app.Quit()
