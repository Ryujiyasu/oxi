# -*- coding: utf-8 -*-
"""S532 repro: is the yakumono PAIR (。」 / ）」) compressed unconditionally by Word
(independent of justify demand), and does Oxi match?

Builds a docx (settings: characterSpacingControl=compressPunctuation, compat=15,
MS Gothic 12pt) with three paragraphs containing the pair mid-text:
  P1 jc=center (loose line -> no justify demand)
  P2 jc=both, SHORT line (loose -> no demand)
  P3 jc=both, line wraps (full -> demand exists)
Measures the 。advance (next_x - x) around the 。」 pair in Word (PDF text extraction)
and Oxi (--dump-glyphs). cp932-safe ASCII out."""
import os, sys, io, zipfile, subprocess, json
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT = os.path.join('c:/tmp', 's532_pair.docx')
DWRITE = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')

SHORT = 'やすく統一的なものとする。」とされた'
LONG = ('在り方について「国が著作権者である著作物については広く二次利用を認める形で表示する。'
        '当該表示については、できるだけ分かりやすく統一的なものとする。」とされたことを踏まえ、'
        '各府省ウェブサイトの利用ルールの見直しを行う。')

CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''
SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:characterSpacingControl w:val="compressPunctuation"/>
<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
</w:settings>'''


def para(text, jc):
    return ('<w:p><w:pPr><w:jc w:val="%s"/><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック"/><w:sz w:val="24"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hint="eastAsia"/><w:sz w:val="24"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (jc, text))


P4 = '規約（第１版）」の解説である'      # ）」 close+close
P5 = 'これを基本とする。「次の章」へ続く'  # 。「 close+open AND 」へ
P6 = '「公共データ利用規約（第1.0版）」の解説'  # the real d77a title text


def build():
    body = (para(SHORT, 'center') + para(SHORT, 'both') + para(LONG, 'both')
            + para(P4, 'center') + para(P5, 'center') + para(P6, 'center')
            + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
              '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body>%s</w:body></w:document>' % body)
    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/settings.xml', SETTINGS)
    print('built', OUT)


WORD_PDF = r'''
import sys, os
import win32com.client, pythoncom
pythoncom.CoInitialize()
word = win32com.client.Dispatch("Word.Application"); word.Visible=False; word.DisplayAlerts=False
doc = word.Documents.Open(sys.argv[1], ReadOnly=True, AddToRecentFiles=False)
doc.ExportAsFixedFormat(OutputFileName=sys.argv[2], ExportFormat=17, OpenAfterExport=False, OptimizeFor=0)
doc.Close(SaveChanges=False); word.Quit(); pythoncom.CoUninitialize()
'''


def word_pairs():
    pdf = 'c:/tmp/s532_pair.pdf'
    subprocess.run([sys.executable, '-c', WORD_PDF, OUT, pdf], capture_output=True, timeout=60)
    import fitz
    d = fitz.open(pdf)
    page = d[0]
    out = []
    for blk in page.get_text('rawdict')['blocks']:
        for ln in blk.get('lines', []):
            chars = []
            for sp in ln['spans']:
                for c in sp['chars']:
                    chars.append((c['c'], c['origin'][0]))
            PUNCT = '。、）」「（．，'
            for i in range(len(chars) - 1):
                if chars[i][0] in PUNCT and chars[i+1][0] in PUNCT:
                    adv = chars[i+1][1] - chars[i][1]
                    adv2 = (chars[i+2][1] - chars[i+1][1]) if i + 2 < len(chars) else -1
                    pair = chars[i][0] + chars[i+1][0]
                    out.append((round(ln['bbox'][1], 1), pair, adv, adv2, ''))
    return out


def oxi_pairs():
    subprocess.run([DWRITE, OUT, 'c:/tmp/s532_pair', '150', '--dump-glyphs=c:/tmp/s532_glyphs.json'],
                   capture_output=True)
    d = json.load(io.open('c:/tmp/s532_glyphs.json', encoding='utf-8'))
    gl = d['pages'][0]['glyphs']
    from collections import defaultdict
    lines = defaultdict(list)
    for g in gl:
        lines[round(g['top'], 1)].append(g)
    out = []
    PUNCT = '。、）」「（．，'
    for top in sorted(lines):
        L = sorted(lines[top], key=lambda g: g['x'])
        for i in range(len(L) - 1):
            if L[i]['char'] in PUNCT and L[i+1]['char'] in PUNCT:
                adv = L[i+1]['x'] - L[i]['x']
                adv2 = (L[i+2]['x'] - L[i+1]['x']) if i + 2 < len(L) else -1
                pair = L[i]['char'] + L[i+1]['char']
                out.append((top, pair, round(adv, 2), round(adv2, 2)))
    return out


if __name__ == '__main__':
    build()
    op = oxi_pairs()
    wp = word_pairs()
    print('P1=center short, P2=both short, P3=both wrap, P4=)] pair, P5=close+open, P6=title. fs=12.')
    print('pair (first advance, second advance):')
    for t, c, a, a2 in op:
        print('  OXI  top=%6.1f pair=%s first=%6.2f second=%6.2f' % (t, c, a, a2))
    for t, c, a, a2, _ in wp:
        print('  WORD top=%6.1f pair=%s first=%6.2f second=%6.2f' % (t, c, a, a2))
