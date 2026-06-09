# -*- coding: utf-8 -*-
"""S521: verify the table ROW-OVERHEAD under-count (insideH border) on a controlled docGrid table.
Repro: docGrid linePitch=292 (linesAndChars, ALH), a 1-col table with 12 single-line rows, MS
Mincho 10.5pt, default single borders. Measure each row's content baseline (Word PDF vs Oxi dump).
Word row-to-row pitch should = grid pitch (14.6) + insideH border; if Oxi's pitch is the border
amount SHORTER, the under-count is confirmed. Variants toggle borders (none vs single vs thick).
cp932-safe: UTF-8 file, results to file, ASCII out."""
import os, io, zipfile, subprocess, json
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT = os.path.join(ROOT, 'tools', 'golden-test', 'repros', 'rowoverhead')
EXE = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
os.makedirs(OUT, exist_ok=True)
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
SETTINGS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:adjustLineHeightInTable/></w:settings>' % NS)
MIN = 'ＭＳ 明朝'

def rpr():
    return '<w:rPr><w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/><w:sz w:val="21"/></w:rPr>' % (MIN, MIN, MIN)

def borders(sz):
    # sz in eighths of a point; 0 = no borders
    if sz == 0:
        return ('<w:tblBorders><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/>'
                '<w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders>')
    e = ('<w:top w:val="single" w:sz="%d" w:space="0" w:color="000000"/>'
         '<w:left w:val="single" w:sz="%d" w:space="0" w:color="000000"/>'
         '<w:bottom w:val="single" w:sz="%d" w:space="0" w:color="000000"/>'
         '<w:right w:val="single" w:sz="%d" w:space="0" w:color="000000"/>'
         '<w:insideH w:val="single" w:sz="%d" w:space="0" w:color="000000"/>'
         '<w:insideV w:val="single" w:sz="%d" w:space="0" w:color="000000"/>') % (sz, sz, sz, sz, sz, sz)
    return '<w:tblBorders>%s</w:tblBorders>' % e

def build(name, bsz, nrows=12):
    rows = ''
    for i in range(nrows):
        rows += ('<w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>'
                 '<w:p><w:r>%s<w:t>あ%d</w:t></w:r></w:p></w:tc></w:tr>' % (rpr(), i))
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/>%s<w:tblLayout w:type="fixed"/></w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>%s</w:tbl>' % (borders(bsz), rows))
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="851" w:footer="397"/>'
            '<w:docGrid w:type="linesAndChars" w:linePitch="292" w:charSpace="1453"/></w:sectPr>')
    doc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document %s><w:body>%s%s</w:body></w:document>' % (NS, tbl, sect)
    p = os.path.join(OUT, name)
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RELS)
        z.writestr('word/document.xml', doc); z.writestr('word/settings.xml', SETTINGS)
    return p

def word_bl(docx):
    import win32com.client, pythoncom, fitz
    pdf = os.path.splitext(docx)[0] + '.pdf'
    pythoncom.CoInitialize(); w = win32com.client.DispatchEx('Word.Application'); w.Visible = False
    try:
        d = w.Documents.Open(os.path.abspath(docx), ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()
    bl = []
    for blk in fitz.open(pdf)[0].get_text('rawdict').get('blocks', []):
        for ln in blk.get('lines', []):
            for sp in ln.get('spans', []):
                ch = sp.get('chars', [])
                if ch and ch[0]['c'] == 'あ':
                    bl.append(ch[0]['origin'][1])
    return sorted(bl)

def oxi_bl(docx):
    pre = os.path.join('c:/tmp', os.path.splitext(os.path.basename(docx))[0])
    gj = pre + '_g.json'
    subprocess.run([EXE, os.path.abspath(docx), pre, '72', '--dump-glyphs=' + gj], capture_output=True, text=True)
    return sorted(x['baseline'] for x in json.load(open(gj, encoding='utf-8'))['pages'][0]['glyphs'] if x['char'] == 'あ')

def main():
    L = ['S521 table ROW-OVERHEAD: docGrid 292 single-line-row table, Word vs Oxi row pitch (insideH border toggle)']
    for name, bsz, lab in [('ro_noborder.docx', 0, 'no borders'),
                           ('ro_b4.docx', 4, 'single 0.5pt (sz=4)'),
                           ('ro_b8.docx', 8, 'single 1.0pt (sz=8)'),
                           ('ro_b16.docx', 16, 'single 2.0pt (sz=16)')]:
        dx = build(name, bsz)
        wb = word_bl(dx); ob = oxi_bl(dx)
        n = min(len(wb), len(ob))
        if n < 3:
            L.append('%s: too few (w=%d o=%d)' % (lab, len(wb), len(ob))); continue
        wp = [round(wb[i+1]-wb[i], 3) for i in range(n-1)]
        op = [round(ob[i+1]-ob[i], 3) for i in range(n-1)]
        wmean = sum(wp)/len(wp); omean = sum(op)/len(op)
        L.append('%-22s | W_rowpitch mean=%.3f set=%s | O mean=%.3f set=%s | O-W mean=%+.3f | accum(last-first dy)=%+.2f' % (
            lab, wmean, sorted(set(wp)), omean, sorted(set(op)), omean-wmean, (ob[n-1]-ob[0])-(wb[n-1]-wb[0])))
    txt = '\n'.join(L)
    io.open('c:/tmp/_s521_out.txt', 'w', encoding='utf-8').write(txt+'\n')
    print(txt)

if __name__ == '__main__':
    main()
