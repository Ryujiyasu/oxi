# -*- coding: utf-8 -*-
"""Does a cell's `nil` border suppress the table's insideH / insideV on a shared edge?

ja/forms__031af703 p1: the box «使用する麻酔薬の名称及び量» sits under a full-width
row whose cell declares bottom nil, beside a vMerge margin column whose cells
declare top nil. Word draws the box's top and left rules; Oxi draws neither.

Arms (compat 15, tblBorders all single sz4): 2 or 3 rows x 2 columns.
  upper_nil   row 1 = one gridSpan-2 cell with bottom nil
  lower_nil   row 2 cells with top nil
  both_nil    both
  vm_topnil   column 1 vMerge restart in row 2 with top nil (continue in row 3)
  vm_plain    the same without nil
Readout: Word PDF horizontal rules at the row-1/row-2 edge and vertical rules
between the columns in rows 2-3.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/nil_edge'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
SET = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat>'
       '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')


def tc(w, text, extra='', borders=''):
    b = f'<w:tcBorders>{borders}</w:tcBorders>' if borders else ''
    return f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{extra}{b}</w:tcPr><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>'


def arm_rows(name):
    if name in ('upper_nil', 'lower_nil', 'both_nil'):
        up = '<w:bottom w:val="nil"/>' if name in ('upper_nil', 'both_nil') else ''
        lo = '<w:top w:val="nil"/>' if name in ('lower_nil', 'both_nil') else ''
        return [tc(6000, 'TOP', '<w:gridSpan w:val="2"/>', up), tc(3000, 'L', '', lo) + tc(3000, 'R', '', lo)]
    nil = '<w:top w:val="nil"/>' if name == 'vm_topnil' else ''
    return [tc(1000, 'a') + tc(5000, 'b'),
            tc(1000, '', '<w:vMerge w:val="restart"/>', nil) + tc(5000, 'BOX'),
            tc(1000, '', '<w:vMerge/>', nil) + tc(5000, 'BOX2')]


def document(name):
    b = ''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>' for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
    grid = '<w:gridCol w:w="3000"/><w:gridCol w:w="3000"/>' if 'nil' in name and not name.startswith('vm') else '<w:gridCol w:w="1000"/><w:gridCol w:w="5000"/>'
    rows = ''.join(f'<w:tr><w:trPr><w:trHeight w:val="600"/></w:trPr>{r}</w:tr>' for r in arm_rows(name))
    tbl = f'<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblBorders>{b}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>{grid}</w:tblGrid>{rows}</w:tbl>'
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body><w:p/>{tbl}<w:p/>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


if __name__ == '__main__':
    import win32com.client, pymupdf
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for name in ('upper_nil', 'lower_nil', 'both_nil', 'vm_topnil', 'vm_plain'):
            at = os.path.abspath(os.path.join(OUT, name + '.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/settings.xml', SET)
                z.writestr('word/document.xml', document(name))
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            P = pymupdf.open(pdf)[0]
            hs = sorted((round(g['rect'].y0, 1), round(g['rect'].x0, 1), round(g['rect'].x1, 1)) for g in P.get_drawings() if g['rect'].height < 2 and g['rect'].width > 2)
            vs = sorted((round(g['rect'].x0, 1), round(g['rect'].y0, 1), round(g['rect'].y1, 1)) for g in P.get_drawings() if g['rect'].width < 2 and g['rect'].height > 2)
            print(f'{name:10} H {hs}\n           V {vs}', flush=True)
    finally:
        app.Quit()
