# -*- coding: utf-8 -*-
"""S502 center-vs-right repro: a MEDIUM grid line (real alignment slack) in a wide cell,
jc=center vs jc=right, charSpace=+1453 (expansion). Measure first-char x Word/ON/OFF to
confirm: center -> grid-EXPANDED width correct (ON); right -> NATURAL width correct (OFF).
cp932-safe: UTF-8 file, ASCII out."""
import os, sys, json, subprocess, tempfile, zipfile
import fitz
import win32com.client, pythoncom

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DW = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
OUT = os.path.join(ROOT, 'tools', 'golden-test', 'repros', 'cellpos')
DPI = 150
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CJK = 'あいうえおかきくけこ'  # 10 chars medium, real slack
NEEDLE = 'あ'

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
RPR = '<w:rPr><w:rFonts w:eastAsia="ＭＳ 明朝" w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>'  # fs=10.5 like b35


def doc_xml(jc):
    para = '<w:p><w:pPr><w:jc w:val="%s"/></w:pPr><w:r>%s<w:t>%s</w:t></w:r></w:p>' % (jc, RPR, CJK)
    c0 = '<w:tc><w:tcPr><w:tcW w:w="2110" w:type="dxa"/></w:tcPr><w:p/></w:tc>'
    c1 = '<w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>%s</w:tc>' % para
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
           '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
           '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>'
           '<w:tblCellMar><w:left w:w="12" w:type="dxa"/><w:right w:w="12" w:type="dxa"/></w:tblCellMar></w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="2110"/><w:gridCol w:w="6000"/></w:tblGrid>'
           '<w:tr>%s%s</w:tr></w:tbl>' % (c0, c1))
    ref = '<w:p><w:r><w:t>REF</w:t></w:r></w:p>'
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>'
            '<w:docGrid w:type="linesAndChars" w:linePitch="350" w:charSpace="-2714"/></w:sectPr>')  # b35 grid (compression)
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document %s><w:body>%s%s%s</w:body></w:document>' % (NS, ref, tbl, sect)


def build(name, jc):
    p = os.path.join(OUT, name)
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', WRELS); z.writestr('word/settings.xml', SETTINGS)
        z.writestr('word/document.xml', doc_xml(jc))
    return p


def word_x(docx):
    docx = os.path.abspath(docx); pdf = os.path.splitext(docx)[0] + '_rt.pdf'
    pythoncom.CoInitialize()
    w = win32com.client.DispatchEx('Word.Application'); w.Visible = False
    try:
        d = w.Documents.Open(docx, ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()
    for blk in fitz.open(pdf)[0].get_text('rawdict').get('blocks', []):
        for ln in blk.get('lines', []):
            for sp in ln.get('spans', []):
                for ch in sp.get('chars', []):
                    if ch['c'] == NEEDLE:
                        return round(ch['origin'][0], 2)
    return None


def oxi_x(docx, disable):
    e = dict(os.environ)
    if disable:
        e['OXI_S502_DISABLE'] = '1'
    fd, jp = tempfile.mkstemp(suffix='.json', dir='c:/tmp'); os.close(fd)
    subprocess.run([DW, os.path.abspath(docx), tempfile.mktemp(dir='c:/tmp'), str(DPI),
                    '--dump-glyphs=' + jp], capture_output=True, timeout=300, env=e)
    d = json.load(open(jp, encoding='utf-8')); os.unlink(jp)
    for page in d['pages']:
        for g in page['glyphs']:
            if g['char'] == NEEDLE:
                return round(g['x'], 2)
    return None


def main():
    L = ['S502 center-vs-right (10-char medium grid line, expansion charSpace=+1453)',
         '%-22s %9s %9s %9s %9s %9s' % ('variant', 'Word', 'ON', 'OFF', '|ON-W|', '|OFF-W|')]
    for name, jc in [('cr_center.docx', 'center'), ('cr_right.docx', 'right')]:
        p = build(name, jc)
        wx = word_x(p); on = oxi_x(p, False); off = oxi_x(p, True)
        L.append('%-22s %9.2f %9.2f %9.2f %9.2f %9.2f  -> %s' % (
            name, wx, on, off, abs(on - wx), abs(off - wx),
            'ON correct' if abs(on - wx) < abs(off - wx) else 'OFF correct'))
    with open('c:/tmp/_s502_cr_out.txt', 'w', encoding='utf-8') as f:
        f.write('\n'.join(L) + '\n')
    print('wrote c:/tmp/_s502_cr_out.txt')


if __name__ == '__main__':
    main()
