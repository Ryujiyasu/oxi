# -*- coding: utf-8 -*-
"""How does Word size an AUTO table that declares no widths at all (empty
<w:gridCol/>, no tcW, tblW auto)?

en/reference__00f2bb61 (METEOR export): 4 columns, Word's text starts 60.86 /
97.82 / 443.71 / 510.58 -- the short columns wrap at their widest word, the
«Obligation» column is exactly its widest line. Hypothesis: min-content
(widest word) + the leftover shared in proportion to max - min, or every max
when the maxima fit.

Arms (Calibri 11, tblCellMar 108 -- a styles-less package has no Normal
Table margins otherwise -- compat 15, 468pt column):
  fit    : «A» / «Bee cee» / «Dee» -- the maxima fit
  long   : «Seq No.» / long sentence x3 / «Mandatory» / «Max occurs»
  two    : two long sentences of different length + «Short words here»
Readout: Word PDF x of each column's first text line in row 1 and the right
table border.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/auto_content_cols'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
SETTINGS = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
LONG = 'This data element refers to whether the counsellors underwent any formal training during the reporting period.'
ARMS = {
    'fit': [['A', 'Bee cee', 'Dee'], ['1', 'two', 'three']],
    'long': [['Seq No.', 'Metadata item', 'Obligation', 'Max occurs'], ['1', LONG, 'Mandatory', '1'], ['2', LONG + ' ' + LONG, 'Conditional', '99']],
    'two': [['First column header', 'Second'], [LONG, LONG[:40]], ['Short words here', 'x']],
}


def document(rows):
    ppr = '<w:pPr><w:spacing w:after="0"/></w:pPr>'
    rpr = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>'
    b = ''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>' for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
    trs = ''.join('<w:tr>' + ''.join(f'<w:tc><w:p>{ppr}<w:r>{rpr}<w:t xml:space="preserve">{t}</w:t></w:r></w:p></w:tc>' for t in r) + '</w:tr>' for r in rows)
    grid = '<w:gridCol/>' * len(rows[0])
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
            f'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>{b}</w:tblBorders><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>{grid}</w:tblGrid>{trs}</w:tbl><w:p/>'
            '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


if __name__ == '__main__':
    import win32com.client, pymupdf
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for name, rows in ARMS.items():
            at = os.path.abspath(os.path.join(OUT, name + '.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/settings.xml', SETTINGS)
                z.writestr('word/document.xml', document(rows))
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            P = pymupdf.open(pdf)[0]
            vx = sorted({round(it[1].x0, 2) for dr in P.get_drawings() for it in dr['items']
                         if it[0] == 're' and it[1].width < it[1].height and it[1].height > 5})
            print(f'{name:6} vertical rules {vx}', flush=True)
    finally:
        app.Quit()
