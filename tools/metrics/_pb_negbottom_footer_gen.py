# -*- coding: utf-8 -*-
"""Does a footer WITH TEXT push the body up when the bottom margin is negative?

ja/correspondence__04a3e3e1 (pgMar bottom -284, footer 510, the footer prints the
page number): Word's page 1 runs to 825.7 (= 841.9 - 14.2 - 2.2), over the
footer's number at 805; Oxi stopped above the footer and pushed two empty
paragraphs to page 2. S1267's probe (-284 vs +284 indistinguishable) had an
ink-free footer.

Arms: bottom -284 / +284 / -1134 x footer empty / one text line / three lines;
80 numbered 12pt exact lines. Readout: Word PDF last line on page 1.
"""
import os, sys, zipfile, itertools
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/negbottom_footer'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>')
PPR = '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>'


def footer(n):
    paras = ''.join(f'<w:p>{PPR}<w:r><w:t>FOOTER{i}</w:t></w:r></w:p>' for i in range(n)) or f'<w:p>{PPR}</w:p>'
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr {W}>{paras}</w:ftr>'


def document(bottom):
    body = ''.join(f'<w:p>{PPR}<w:r><w:t>L{i:02d}</w:t></w:r></w:p>' for i in range(80))
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}'
            '<w:sectPr><w:footerReference w:type="default" r:id="rId9"/><w:pgSz w:w="11906" w:h="16838"/>'
            f'<w:pgMar w:top="1440" w:right="1440" w:bottom="{bottom}" w:left="1440" w:header="851" w:footer="510" w:gutter="0"/></w:sectPr></w:body></w:document>')


if __name__ == '__main__':
    import win32com.client, pymupdf
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for bottom, nf in itertools.product((-284, 284, -1134), (0, 1, 3)):
            name = f'b{bottom}_f{nf}'
            at = os.path.abspath(os.path.join(OUT, name + '.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/footer1.xml', footer(nf))
                z.writestr('word/document.xml', document(bottom))
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            P = pymupdf.open(pdf)[0]
            ls = sorted((round(l['bbox'][3], 1), l['spans'][0]['text']) for b in P.get_text('dict')['blocks'] for l in b.get('lines', [])
                        if l['spans'][0]['text'].startswith('L'))
            fs = sorted(round(l['bbox'][1], 1) for b in P.get_text('dict')['blocks'] for l in b.get('lines', [])
                        if l['spans'][0]['text'].startswith('FOOTER'))
            print(f'{name:10} lines {len(ls)} last {ls[-1]} footer_tops {fs}', flush=True)
    finally:
        app.Quit()
