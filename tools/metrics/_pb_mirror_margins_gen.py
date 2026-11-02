# -*- coding: utf-8 -*-
"""Where does the text sit on even pages under w:mirrorMargins?

en/technical__019d92d9 (mirrorMargins, left 1080 right 720): Word's page 2 text
starts at 36pt, Oxi keeps 54 on every page. Arms: mirror on/off x gutter 0/360 x
gutterAtTop. Readout: the left x of each page's first line and of a full-width
rule-free paragraph's right end, pages 1-3.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/mirror_margins'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
ARMS = {'plain': ('', 0), 'mirror': ('<w:mirrorMargins/>', 0), 'mirror_g360': ('<w:mirrorMargins/>', 360),
        'plain_g360': ('', 360), 'mirror_g360_top': ('<w:mirrorMargins/><w:gutterAtTop/>', 360)}
LINE = 'Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut labore et dolore magna aliqua'


def document(gutter):
    paras = ''.join(f'<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t>P{i:03d} {LINE}</w:t></w:r></w:p>' for i in range(140))
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{paras}'
            '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="720" w:bottom="1440" w:left="1080" '
            f'w:header="720" w:footer="720" w:gutter="{gutter}"/></w:sectPr></w:body></w:document>')


if __name__ == '__main__':
    import win32com.client, pymupdf
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for name, (st, g) in ARMS.items():
            at = os.path.abspath(os.path.join(OUT, f'{name}.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/_rels/document.xml.rels', DRELS)
                z.writestr('word/settings.xml', f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>{st}</w:settings>')
                z.writestr('word/document.xml', document(g))
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            P = pymupdf.open(pdf)
            out = []
            for i in range(min(3, len(P))):
                ws = P[i].get_text('words')
                out.append(f'p{i+1} L {min(w[0] for w in ws):.2f} R {max(w[2] for w in ws):.2f} T {min(w[1] for w in ws):.2f}')
            print(f'{name:16} ' + '  '.join(out), flush=True)
    finally:
        app.Quit()
