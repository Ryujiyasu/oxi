# -*- coding: utf-8 -*-
"""S531 minimal repro: isolate the single-cell cellMar wrap-budget spec.

Builds a docx with TWO single-cell bordered tables holding the SAME long full-width
hiragana paragraph (MS Mincho 10.5pt, justified):
  - Table A: tblStyle=TS1 (style-inherited tblCellMar left/right=108tw), NO tblCellMar in tblPr
  - Table B: NO style, NO cellMar (full cell width)
If Word reserves the inherited cellMar as padding, Table A wraps ~1 char EARLIER per line than B.
Measures chars-on-line-1 for each via Word COM (wdFirstCharacterLineNumber) and via the Oxi
glyph dump, and reports whether Oxi matches Word. cp932-safe (ASCII out)."""
import os, sys, io, zipfile, subprocess, glob, json
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT_DOCX = os.path.join('c:/tmp', 's531_repro.docx')
DWRITE = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')

# 40 full-width hiragana (all same advance) -> clean wrap counting
TEXT = ('あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろ'
        'わをんがぎぐげござじずぜぞだぢづでどばびぶべぼぱぴぷぺぽ')  # ~70 chars

CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="MS Mincho" w:hAnsi="Century"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="table" w:customStyle="1" w:styleId="TS1"><w:name w:val="TS1"/><w:basedOn w:val="TableNormal"/><w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>'''


def cell_para():
    runs = '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t xml:space="preserve">%s</w:t></w:r>' % TEXT
    return '<w:p><w:pPr><w:jc w:val="both"/></w:pPr>%s</w:p>' % runs


def table(style_id):
    style = '<w:tblStyle w:val="%s"/>' % style_id if style_id else ''
    borders = ('<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
               '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
               '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
               '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
               '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
               '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders>')
    cellmar = '' if style_id else '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>'
    return ('<w:tbl><w:tblPr>%s<w:tblW w:w="6000" w:type="dxa"/><w:tblInd w:w="0" w:type="dxa"/>%s'
            '<w:tblLayout w:type="fixed"/>%s</w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>'
            '<w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>%s</w:tc></w:tr></w:tbl>'
            % (style, borders, cellmar, cell_para()))


def build():
    body = (table('TS1') + '<w:p/>' + table(None) + '<w:p/>'
            + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
              '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body>%s</w:body></w:document>' % body)
    with zipfile.ZipFile(OUT_DOCX, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/document.xml', doc)
        z.writestr('word/styles.xml', STYLES)
    print('built', OUT_DOCX)


WORD_MEASURE = r'''
import sys
import win32com.client, pythoncom
pythoncom.CoInitialize()
word = win32com.client.Dispatch("Word.Application"); word.Visible=False; word.DisplayAlerts=False
doc = word.Documents.Open(sys.argv[1], ReadOnly=True, AddToRecentFiles=False)
res=[]
try:
    for ti in range(1, doc.Tables.Count+1):
        rng = doc.Tables(ti).Cell(1,1).Range
        # line number of first char
        base = doc.Range(rng.Start, rng.Start).Information(10)  # wdFirstCharacterLineNumber
        n=0
        # count chars on line 1: char k occupies [k,k+1); its line = Information(10) of Range(k,k+1)
        for k in range(rng.Start, rng.End):
            ln = doc.Range(k, k+1).Information(10)
            if ln is not None and ln != base:
                break
            n+=1
        res.append(n)
    print("WORD_LINE1 " + " ".join(str(x) for x in res))
finally:
    doc.Close(SaveChanges=False); word.Quit(); pythoncom.CoUninitialize()
'''


def measure_word():
    r = subprocess.run([sys.executable, '-c', WORD_MEASURE, OUT_DOCX],
                       capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=60)
    for line in (r.stdout or '').splitlines():
        if line.startswith('WORD_LINE1'):
            return [int(x) for x in line.split()[1:]]
    print('word stderr:', (r.stderr or '')[:300])
    return None


def measure_oxi():
    subprocess.run([DWRITE, OUT_DOCX, 'c:/tmp/s531_repro', '150', '--dump-glyphs=c:/tmp/s531_repro_glyphs.json'],
                   capture_output=True, text=True)
    d = json.load(io.open('c:/tmp/s531_repro_glyphs.json', encoding='utf-8'))
    gl = d['pages'][0]['glyphs']
    from collections import defaultdict
    lines = defaultdict(list)
    for g in gl:
        lines[round(g['top'], 1)].append(g)
    # cell lines: minx > page margin (56.7). group by table via x-start and top order.
    tops = sorted(lines)
    # report first-line char counts per distinct table block (a block = contiguous lines, gap>20pt separates)
    blocks = []
    cur = []
    prev = None
    for t in tops:
        if prev is not None and t - prev > 20:
            if cur: blocks.append(cur)
            cur = []
        cur.append(t); prev = t
    if cur: blocks.append(cur)
    out = []
    for b in blocks:
        out.append(len(lines[b[0]]))  # chars on first line of the block
    return out, blocks, lines


if __name__ == '__main__':
    build()
    woxi, blocks, lines = measure_oxi()
    wword = measure_word()
    print('TEXT len = %d full-width chars' % len(TEXT))
    print('Oxi line-1 char counts per block:', woxi)
    print('Word line-1 char counts per table:', wword)
    print('(Table A = style cellMar 108/108; Table B = no cellMar. A should be ~1 LESS than B.)')
