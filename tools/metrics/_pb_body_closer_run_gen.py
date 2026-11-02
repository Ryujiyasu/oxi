# -*- coding: utf-8 -*-
"""Does a body line hang a RUN of closing marks («。）») past the column end?

ja/reports__67de0cc0 p5 (compat 14, linesAndChars grid, 2 columns, ＭＳ ゴシック
10.5, compressPunctuation): «…御相談ください。）» -- Word ends the line at «くださ»
and carries «い。）» down; Oxi hangs both marks and keeps 22 characters.

Arms: a full line of N kanji (exactly the column) then «い» + tail + «あいう…»,
tail = 。 / 。） / ） / 、」 / 」。 ; compat 14 / 15; grid linesAndChars vs none.
Readout: Word PDF characters on line 1 and the first character of line 2.
"""
import os, sys, zipfile
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
OUT = 'tests/fixtures/body_closer_run'; os.makedirs(OUT, exist_ok=True)
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RR = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
TAILS = {'p': '。', 'pk': '。）', 'k': '）', 'ck': '、」', 'kp': '」。'}
BODY = '住まいの地域の福祉事務所に御相談くださ'  # 19 chars


def settings(c):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
            '<w:characterSpacingControl w:val="compressPunctuation"/><w:balanceSingleByteDoubleByteWidth/><w:compat>'
            f'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{c}"/></w:compat></w:settings>')


def document(tail, grid):
    rpr = '<w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック"/><w:sz w:val="21"/></w:rPr>'
    text = BODY + 'い' + tail + 'あいうえおかきくけこ'
    g = '<w:docGrid w:type="linesAndChars" w:linePitch="411" w:charSpace="2048"/>' if grid else '<w:docGrid w:linePitch="360"/>'
    # text width 20 em of 10.5 = 210pt = 4200tw (grid: the cell pitch decides)
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
            f'<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r>{rpr}<w:t>{text}</w:t></w:r></w:p>'
            f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="{11906 - 1440 - 4200}" w:bottom="1440" w:left="1440" '
            f'w:header="720" w:footer="720" w:gutter="0"/>{g}</w:sectPr></w:body></w:document>')


if __name__ == '__main__':
    import win32com.client, pymupdf
    app = win32com.client.DispatchEx('Word.Application'); app.Visible = False; app.DisplayAlerts = 0
    try:
        for c in (14, 15):
            for grid in (False, True):
                for tk, tail in TAILS.items():
                    name = f'c{c}_{"grid" if grid else "nogrid"}_{tk}'
                    at = os.path.abspath(os.path.join(OUT, name + '.docx'))
                    with zipfile.ZipFile(at, 'w', zipfile.ZIP_DEFLATED) as z:
                        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RR)
                        z.writestr('word/_rels/document.xml.rels', DRELS); z.writestr('word/settings.xml', settings(c))
                        z.writestr('word/document.xml', document(tail, grid))
                    pdf = at[:-5] + '.pdf'
                    d = app.Documents.Open(at, ReadOnly=True)
                    try:
                        d.ExportAsFixedFormat(pdf, 17)
                    finally:
                        d.Close(False)
                    P = pymupdf.open(pdf)[0]
                    lines = {}
                    for b in P.get_text('rawdict')['blocks']:
                        for l in b.get('lines', []):
                            for s in l['spans']:
                                for ch in s['chars']:
                                    lines.setdefault(round(s['origin'][1], 1), []).append(ch['c'])
                    ys = sorted(lines)
                    l1 = ''.join(lines[ys[0]]); l2 = ''.join(lines[ys[1]]) if len(ys) > 1 else ''
                    print(f'{name:16} L1 n{len(l1):2} ...{l1[-5:]}  L2 {l2[:4]}', flush=True)
    finally:
        app.Quit()
