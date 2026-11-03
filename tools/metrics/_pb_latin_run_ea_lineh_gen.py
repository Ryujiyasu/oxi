# -*- coding: utf-8 -*-
"""Does a Latin-only run's eastAsia face set its line's height?

ja/administrative__0ba21b0c: a text box heading «～個人懇談会～» (UD デジタル 教科書体
NP-B 10) + «Reunião de pais e mestres» (Arial 14, rFonts eastAsia UD デジタル 教科書体
NP-B): Word's next line sits 3.9pt lower than Oxi's (Oxi line 16.1 = Arial 14).

Arms (body, compat 15, single, no grid, ＭＳ 明朝 10.5 base): three paragraphs
A / X / B, X = «Reunião de pais» Arial 14 with eastAsia ＭＳ 明朝 / Arial / none,
alone or after a 10pt ＭＳ 明朝 run «～個人～». Readout: baselines of A, X, B.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/latin_run_ea_lineh'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
SP = ('<w:pPr><w:spacing w:after="0" w:line="0" w:lineRule="atLeast"/></w:pPr>' if os.environ.get('PB_ATL')
      else '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>')
BASE = '<w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'


def p(runs):
    return f'<w:p>{SP}{runs}</w:p>'


def r(t, rf, sz):
    return f'<w:r><w:rPr>{rf}<w:sz w:val="{sz}"/></w:rPr><w:t xml:space="preserve">{t}</w:t></w:r>'


def document(ea, lead):
    rf = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"' + (f' w:eastAsia="{ea}"' if ea else '') + '/>'
    x = (r('～個人～', BASE, 20) if lead else '') + r('Reunião de pais', rf, 28)
    body = p(r('ああああ', BASE, 21)) + p(x) + p(r('いいいい', BASE, 21))
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')


ARMS = [(ea, lead) for ea in ('ＭＳ 明朝', 'Arial', None) for lead in (False, True)]
if os.environ.get('PB_UD'):
    BASE = '<w:rFonts w:ascii="UD デジタル 教科書体 NP-B" w:eastAsia="UD デジタル 教科書体 NP-B" w:hAnsi="UD デジタル 教科書体 NP-B"/>'
    OUT = OUT + '_ud' + ('_atl' if os.environ.get('PB_ATL') else ''); os.makedirs(OUT, exist_ok=True)
    ARMS = [('Arial', False), ('Arial', True)]

if __name__ == '__main__':
    import win32com.client, fitz
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for ea, lead in ARMS:
            name = f'ea_{ {"ＭＳ 明朝": "mincho", "Arial": "arial", None: "none"}[ea]}_{"lead" if lead else "alone"}'
            at = os.path.abspath(os.path.join(OUT, f'{name}.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/document.xml', document(ea, lead))
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            P = fitz.open(pdf)
            ys = sorted({round(l['spans'][0]['origin'][1], 2) for b in P[0].get_text('dict')['blocks'] for l in b.get('lines', [])})
            print(f'{name:18} baselines {ys}  gaps {[round(b - a, 2) for a, b in zip(ys, ys[1:])]}', flush=True)
    finally:
        app.Quit()
