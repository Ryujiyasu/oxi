# -*- coding: utf-8 -*-
"""S502 cellpos WRAP analysis: build a wrapping center+firstLine variant and report
line-0 composition (char count, first-x, last-x, width) for Word vs Oxi. If Oxi puts
a different #chars on line 0 than Word, the firstLine first-line-wrap width is the bug
(centering exposes it as a first-char x shift). cp932-safe: ASCII out to a file."""
import os, sys, json, subprocess, tempfile, zipfile
import fitz
import win32com.client, pythoncom

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DW = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
OUT = os.path.join(ROOT, 'tools', 'golden-test', 'repros', 'cellpos')
DPI = 150
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CJK = 'あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめも'  # wraps

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
WRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings %s><w:compat><w:adjustLineHeightInTable/>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>' % NS)
RPR = '<w:rPr><w:rFonts w:eastAsia="ＭＳ 明朝" w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'


def doc_xml(jc, firstline):
    ind = '<w:ind w:firstLineChars="100" w:firstLine="247"/>' if firstline else ''
    jct = '<w:jc w:val="%s"/>' % jc
    para = '<w:p><w:pPr>%s%s</w:pPr><w:r>%s<w:t>%s</w:t></w:r></w:p>' % (jct, ind, RPR, CJK)
    c0 = '<w:tc><w:tcPr><w:tcW w:w="2110" w:type="dxa"/></w:tcPr><w:p/></w:tc>'
    c1 = '<w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>%s</w:tc>' % para
    row = '<w:tr>%s%s</w:tr>' % (c0, c1)
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
           '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>'
           '<w:tblCellMar><w:left w:w="12" w:type="dxa"/><w:right w:w="12" w:type="dxa"/></w:tblCellMar></w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="2110"/><w:gridCol w:w="6000"/></w:tblGrid>%s</w:tbl>' % row)
    ref = '<w:p><w:r><w:t>REF</w:t></w:r></w:p>'
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>'
            '<w:docGrid w:type="linesAndChars" w:linePitch="292" w:charSpace="1453"/></w:sectPr>')
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document %s><w:body>%s%s%s</w:body></w:document>' % (NS, ref, tbl, sect)


def build(name, jc, fl):
    p = os.path.join(OUT, name)
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', WRELS); z.writestr('word/settings.xml', SETTINGS)
        z.writestr('word/document.xml', doc_xml(jc, fl))
    return p


def cluster(glyphs, getx, gety):
    gs = sorted(glyphs, key=lambda g: (round(gety(g), 0), getx(g)))
    lines = []
    for g in gs:
        if lines and abs(gety(g) - gety(lines[-1][0])) < 4.0:
            lines[-1].append(g)
        else:
            lines.append([g])
    return lines


def word_lines(docx):
    docx = os.path.abspath(docx)
    pdf = os.path.splitext(docx)[0] + '_rt.pdf'
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx('Word.Application')
    word.Visible = False
    try:
        doc = word.Documents.Open(docx, ReadOnly=True)
        doc.ExportAsFixedFormat(pdf, 17); doc.Close(False)
    finally:
        word.Quit()
    d = fitz.open(pdf)
    gl = []
    for blk in d[0].get_text('rawdict').get('blocks', []):
        for line in blk.get('lines', []):
            for span in line.get('spans', []):
                for ch in span.get('chars', []):
                    c = ch['c']
                    if c.strip() and c != 'R' and c != 'E' and c != 'F':
                        gl.append({'c': c, 'x': ch['origin'][0], 'y': ch['origin'][1]})
    return cluster(gl, lambda g: g['x'], lambda g: g['y'])


def oxi_lines(docx):
    fd, jp = tempfile.mkstemp(suffix='.json', dir='c:/tmp'); os.close(fd)
    subprocess.run([DW, os.path.abspath(docx), tempfile.mktemp(dir='c:/tmp'), str(DPI),
                    '--dump-glyphs=' + jp], capture_output=True, timeout=300)
    d = json.load(open(jp, encoding='utf-8')); os.unlink(jp)
    gl = []
    for page in d['pages']:
        for g in page['glyphs']:
            c = g['char']
            if c.strip() and c not in ('R', 'E', 'F'):
                gl.append({'c': c, 'x': g['x'], 'y': g.get('baseline', g.get('top', 0))})
    return cluster(gl, lambda g: g['x'], lambda g: g['y'])


def desc(line):
    xs = [g['x'] for g in line]
    return 'n=%2d  first_x=%.2f  last_x=%.2f  width=%.2f  text=%s' % (
        len(line), min(xs), max(xs), max(xs) - min(xs), ''.join(g['c'] for g in line))


def main():
    out = 'c:/tmp/_s502_cellpos_wrap_out.txt'
    L = ['S502 cellpos WRAP: line-0 composition Word vs Oxi (center+firstLine bug)']
    for name, jc, fl in [('wrap_center_fl.docx', 'center', True),
                         ('wrap_left_fl.docx', 'left', True)]:
        p = build(name, jc, fl)
        wl = word_lines(p)
        ol = oxi_lines(p)
        L.append('\n=== %s ===' % name)
        L.append('WORD line0: %s' % (desc(wl[0]) if wl else 'NONE'))
        L.append('OXI  line0: %s' % (desc(ol[0]) if ol else 'NONE'))
        if len(wl) > 1:
            L.append('WORD line1: %s' % desc(wl[1]))
        if len(ol) > 1:
            L.append('OXI  line1: %s' % desc(ol[1]))
    with open(out, 'w', encoding='utf-8') as f:
        f.write('\n'.join(L) + '\n')
    print('wrote', out)


if __name__ == '__main__':
    main()
