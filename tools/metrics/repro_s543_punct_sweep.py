# -*- coding: utf-8 -*-
"""S543 punct-type sweep: which yakumono get Word's light -0.75 oikomi
compression on jc=left lines? One para per punct P: 国*30 + P + 国*12 + 国
(44 chars), w:ind right tuned so the 44th char overflows ~0.5pt and fits
ONLY if P compresses >= 0.5. Detect: Word L1 char count (44 = compresses,
43 = does not). kern on/off variants. type=lines, MS Mincho sz=21, jc=left.
ASCII-safe output.
"""
import os
import zipfile

import win32com.client as w32

OUT = os.path.abspath('tools/golden-test/repros/s543_sweep')
os.makedirs(OUT, exist_ok=True)

CT = ('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RELS = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
DRELS = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
SETTINGS = ('<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:characterSpacingControl w:val="compressPunctuation"/>'
            '<w:compat><w:balanceSingleByteDoubleByteWidth/><w:useFELayout/>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>'
            '<w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
            '</w:compat></w:settings>')

# Parametric: S543_SZ half-points (21 = 10.5pt default; 24 = 12pt for the
# cap-scaling question). n chars chosen so the LAST char overflows by
# S543_OVER pt (default 0.5): measure = n*fs - over; ind_right = content - measure.
PUNCTS = ['、', '。', '，', '．', '（', '）', '「', '」', '『', '』', '［', '］', '・', '：', '；', '！', '？']
SZ = int(os.environ.get('S543_SZ', '21'))
FS = SZ / 2.0
OVER = float(os.environ.get('S543_OVER', '0.5'))
CONTENT = (11906 - 1304 * 2) / 20.0  # 464.9pt
N = int(CONTENT // FS)               # max chars at natural
IND_RIGHT_TW = int(round((CONTENT - (N * FS - OVER)) * 20))
P_IDX = N // 2 - 1
print('FS=%.1f N=%d ind_right=%dtw P_IDX=%d' % (FS, N, IND_RIGHT_TW, P_IDX))


def build(docx, kern):
    styles = ('<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
              '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/>'
              + ('<w:kern w:val="2"/>' if kern else '') +
              '<w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style></w:styles>')
    body = ''
    for p in PUNCTS:
        text = '国' * P_IDX + p + '国' * (N - P_IDX - 1)
        body += ('<w:p><w:pPr><w:ind w:right="%d"/><w:jc w:val="left"/>'
                 '<w:rPr><w:sz w:val="%d"/></w:rPr></w:pPr>'
                 '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="%d"/></w:rPr>'
                 '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (IND_RIGHT_TW, SZ, SZ, text))
    doc = ('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body>%s<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="1134" w:right="1304" w:bottom="1134" w:left="1304"/>'
           '<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr></w:body></w:document>') % body
    with zipfile.ZipFile(docx, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DRELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', styles)
        z.writestr('word/settings.xml', SETTINGS)


word = w32.DispatchEx('Word.Application')
word.Visible = False
try:
    for kern in (True, False):
        docx = os.path.join(OUT, 's543_sweep_%s.docx' % ('kern' if kern else 'nokern'))
        build(docx, kern)
        wdoc = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        try:
            print('==== %s ====' % os.path.basename(docx))
            for pi, p in enumerate(wdoc.Paragraphs):
                if pi >= len(PUNCTS):
                    break
                rng = p.Range
                start = rng.Start
                # chars on first line: count chars sharing the y of char 0
                y0 = wdoc.Range(start, start).Information(6)
                n_l1 = 0
                p_adv = None
                prev_x = None
                for i in range(N):
                    r = wdoc.Range(start + i, start + i)
                    if r.Information(6) != y0:
                        break
                    x = r.Information(5)
                    if i == P_IDX + 1:  # advance OF the punct = x(P_IDX+1)-x(P_IDX)
                        p_adv = round(x - prev_x, 2)
                    prev_x = x
                    n_l1 += 1
                print('  P=U+%04X L1=%d %s p_adv=%s' % (
                    ord(PUNCTS[pi]), n_l1,
                    'COMPRESSES(%d)' % N if n_l1 == N else ('no(%d)' % (N - 1) if n_l1 == N - 1 else '??'),
                    p_adv))
        finally:
            wdoc.Close(False)
finally:
    word.Quit()
