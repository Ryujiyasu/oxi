# -*- coding: utf-8 -*-
"""Is «・「» compressed even on a SLACK line (compat 15, compressPunctuation, jc both)?

ja/creative__8048e30f: Word sets «・「働きやすい職場」づくり　等　…　・「問題社員」への
対応相談　　等» (ＭＳ ゴシック 12 bold) with ・ advancing 4.8 (0.4em) and 」 9.6, 19.2pt in
all -- 1.6 cells, beyond the S1490 capacity (k=4 0.775, `_pb_compcap_unit_gen.py`).
Arms: short (slack) paragraphs with ・「 / ・あ / 、「 / 」「, jc both / left, compress /
doNotCompress. Readout: PDF advances of the first three characters.
"""
import os, sys, zipfile, itertools
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/pair_dot_bracket'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
FONT = 'ＭＳ ゴシック'
RPR = f'<w:rPr><w:rFonts w:ascii="{FONT}" w:eastAsia="{FONT}" w:hAnsi="{FONT}"/><w:b/><w:sz w:val="24"/><w:lang w:eastAsia="ja-JP"/></w:rPr>'
TEXTS = ['・「働きやすい」', '・あ働きやすい」', '、「働きやすい」', '」「働きやすい」', 'あ「働きやすい」']


def settings(ctrl):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:characterSpacingControl w:val="{ctrl}"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')


def document(jc):
    body = ''.join(f'<w:p><w:pPr><w:jc w:val="{jc}"/>{RPR}</w:pPr><w:r>{RPR}<w:t>{t}</w:t></w:r></w:p>' for t in TEXTS)
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            'w:header="851" w:footer="992" w:gutter="0"/></w:sectPr></w:body></w:document>')


if __name__ == '__main__':
    import win32com.client, pymupdf
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for jc, ctrl in itertools.product(('both', 'left'), ('compressPunctuation', 'doNotCompress')):
            name = f'{jc}_{ctrl}'
            at = os.path.abspath(os.path.join(OUT, name + '.docx'))
            with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/settings.xml', settings(ctrl))
                z.writestr('word/document.xml', document(jc))
            pdf = at[:-5] + '.pdf'
            d = app.Documents.Open(at, ReadOnly=True)
            try:
                d.ExportAsFixedFormat(pdf, 17)
            finally:
                d.Close(False)
            P = pymupdf.open(pdf)[0]
            for b in P.get_text('rawdict')['blocks']:
                for l in b.get('lines', []):
                    ch = [c for s in l['spans'] for c in s['chars']]
                    xs = [c['bbox'][0] for c in ch]
                    adv = [round(xs[i + 1] - xs[i], 2) for i in range(min(3, len(xs) - 1))]
                    print(f'{name:28} {"".join(c["c"] for c in ch)[:9]:12} adv {adv}', flush=True)
    finally:
        app.Quit()
