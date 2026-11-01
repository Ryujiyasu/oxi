# -*- coding: utf-8 -*-
"""Does the fontTable w:sig Unicode-range set pick a missing face's substitute?

en/creative__04374e00: body in Helvetica Neue (absent), fontTable charset 00,
family auto, panose 02000503000000020004, sig usb0 E50002FF (bit 26 = Georgian).
Word draws it in Sylfaen; the S1444 law (sig present, family auto -> Cambria) and
Oxi say otherwise. Hypothesis: the substitute must cover the scripts the signature
declares, and Sylfaen is the Georgian/Armenian face.

Arms: an invented face «Zqxv Neue» with the document's declaration, then usb0 with
single bits cleared / set. One Word process per arm (Word caches missing-face
substitutes per process). Readout: the PDF span font of a Latin line.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/missing_sig_script'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/></Relationships>')
FACE = 'Zqxv Neue'
# usb0 bit: 9 Cyrillic, 10 Armenian, 24 Thai, 26 Georgian, 29 Latin Ext Additional, 30 Greek Ext, 31 Gen Punct
ARMS = {
    'doc_E50002FF': 'E50002FF',
    'no26_E10002FF': 'E10002FF',
    'no9_E50000FF': 'E50000FF',
    'only26_040000FF': '040000FF',
    'armenian_000004FF': '000004FF',
    'thai_010000FF': '010000FF',
    'latin_000000FF': '000000FF',
}


def font_table(usb0, name=FACE):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts {W}>'
            f'<w:font w:name="{name}"><w:panose1 w:val="02000503000000020004"/><w:charset w:val="00"/><w:family w:val="auto"/>'
            f'<w:pitch w:val="variable"/><w:sig w:usb0="{usb0}" w:usb1="500079DB" w:usb2="00000010" w:usb3="00000000" w:csb0="00000001" w:csb1="00000000"/></w:font>'
            '</w:fonts>')


def document(name=FACE):
    rpr = f'<w:rPr><w:rFonts w:ascii="{name}" w:hAnsi="{name}"/><w:sz w:val="24"/></w:rPr>'
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
            f'<w:p><w:r>{rpr}<w:t>Quick brown fox jumps over the lazy dog</w:t></w:r></w:p>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


def run_arm(name, usb0, face=FACE):
    import win32com.client, pymupdf
    at = os.path.abspath(os.path.join(OUT, name + '.docx'))
    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
        z.writestr('word/_rels/document.xml.rels', DRELS)
        z.writestr('word/fontTable.xml', font_table(usb0, face)); z.writestr('word/document.xml', document(face))
    pdf = at[:-5] + '.pdf'
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        d = app.Documents.Open(at, ReadOnly=True)
        try:
            d.ExportAsFixedFormat(pdf, 17)
        finally:
            d.Close(False)
    finally:
        app.Quit()
    P = pymupdf.open(pdf)[0]
    fonts = {s['font'] for b in P.get_text('dict')['blocks'] for l in b.get('lines', []) for s in l['spans'] if s['text'].strip()}
    return fonts


if __name__ == '__main__':
    for name, usb0 in ARMS.items():
        print(f'{name:20} {run_arm(name, usb0)}', flush=True)
    print(f'{"realname_doc":20} {run_arm("realname_doc", "E50002FF", "Helvetica Neue")}', flush=True)
