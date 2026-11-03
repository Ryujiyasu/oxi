# -*- coding: utf-8 -*-
"""Does a leading ideographic space (U+3000, drawn in the run's eastAsia face)
set the height of a Latin line? ja/administrative__0ba21b0c: «　Estamos
planejando…» (Arial 11, eastAsia UD デジタル 教科書体 NP-B, atLeast 0): Word's
first-to-second baseline gap 12.5, Oxi 11.6.
Arms: paragraph A (one line), X (two lines: «[lead]Lorem…» wrapping), B, with
lead U+3000 / none / ASCII space, eastAsia UD デジタル 教科書体 NP-B / ＭＳ 明朝, rule
auto / atLeast 0. Readout: baselines."""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/ideo_lead_lineh'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
TXT = 'Estamos planejando a reunião de pais e mestres por três dias, de 14 de julho a 16 de julho, para conversar.'
def document(lead, ea, rule):
    sp = '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>' if rule == 'auto' else '<w:spacing w:after="0" w:line="0" w:lineRule="atLeast"/>'
    rf = f'<w:rFonts w:ascii="Arial" w:eastAsia="{ea}" w:hAnsi="Arial"/><w:sz w:val="22"/>'
    run = lambda t: f'<w:r><w:rPr>{rf}</w:rPr><w:t xml:space="preserve">{t}</w:t></w:r>'
    p = lambda body: f'<w:p><w:pPr>{sp}</w:pPr>{body}</w:p>'
    body = p(run('Alpha line')) + p(run({'ideo': '　', 'none': '', 'ascii': ' '}[lead]) + run(TXT)) + p(run('Omega line'))
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="5000" w:bottom="1440" w:left="1080" '
            'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
ARMS = [(l, e, r) for l in ('ideo', 'none', 'ascii') for e in ('UD デジタル 教科書体 NP-B', 'ＭＳ 明朝') for r in ('auto', 'atl')]
if __name__ == '__main__':
    import win32com.client, fitz
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for l, e, r in ARMS:
            name = f'{l}_{"ud" if e.startswith("UD") else "mincho"}_{r}'
            at = os.path.abspath(os.path.join(OUT, f'{name}.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/document.xml', document(l, e, r))
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            ys = sorted({round(ln['spans'][0]['origin'][1], 2) for b in fitz.open(pdf)[0].get_text('dict')['blocks'] for ln in b.get('lines', [])})
            print(f'{name:18} gaps {[round(b - a, 2) for a, b in zip(ys, ys[1:])]}', flush=True)
    finally:
        app.Quit()
