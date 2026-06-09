# -*- coding: utf-8 -*-
"""S520b: the real tokumei docGrid is linePitch=292 (14.6pt), NOT the 360(18.0) my S520 repro used.
14.6 is not 0.75-grid-aligned. Test docGrid linePitch in {292,300,357,360}: Word PDF pitch + mean
vs Oxi dump pitch + mean, and the CUMULATIVE Oxi-Word baseline drift over 30 lines. If Oxi's mean
pitch != Word's for 292, THAT is the tokumei cell-Y drift source. cp932-safe."""
import os, io, zipfile, subprocess
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT = os.path.join(ROOT, 'tools', 'golden-test', 'repros', 'gridpin')
EXE = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')

def build(name, pitch):
    MIN = 'ＭＳ 明朝'
    rpr = '<w:rPr><w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/><w:sz w:val="21"/></w:rPr>' % (MIN, MIN, MIN)
    paras = ''.join('<w:p><w:r>%s<w:t>あ%d</w:t></w:r></w:p>' % (rpr, i) for i in range(30))
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="851" w:footer="397"/>'
            '<w:docGrid w:type="linesAndChars" w:linePitch="%d" w:charSpace="1453"/></w:sectPr>' % pitch)
    doc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document %s><w:body>%s%s</w:body></w:document>' % (NS, paras, sect)
    p = os.path.join(OUT, name)
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RELS); z.writestr('word/document.xml', doc)
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
    return bl

def oxi_bl(docx):
    import json
    pre = os.path.join('c:/tmp', os.path.splitext(os.path.basename(docx))[0])
    gj = pre + '_g.json'
    subprocess.run([EXE, os.path.abspath(docx), pre, '72', '--dump-glyphs=' + gj], capture_output=True, text=True)
    return [x['baseline'] for x in json.load(open(gj, encoding='utf-8'))['pages'][0]['glyphs'] if x['char'] == 'あ']

def main():
    L = ['S520b docGrid linePitch sweep: Word vs Oxi pitch/mean/cumulative drift (MS Mincho 10.5pt)']
    for pitch in [292, 300, 357, 360]:
        dx = build('g%d.docx' % pitch, pitch)
        wb = word_bl(dx); ob = oxi_bl(dx)
        n = min(len(wb), len(ob))
        if n < 5:
            L.append('pitch=%d: too few lines (w=%d o=%d)' % (pitch, len(wb), len(ob))); continue
        wp = [wb[i+1]-wb[i] for i in range(n-1)]; op = [ob[i+1]-ob[i] for i in range(n-1)]
        wmean = sum(wp)/len(wp); omean = sum(op)/len(op)
        # cumulative drift: align first baseline, then last-line Oxi-Word
        drift0 = ob[0] - wb[0]; driftN = ob[n-1] - wb[n-1]
        L.append('pitch=%-4d(%.2fpt) | W_mean=%.4f O_mean=%.4f dMean=%+.4f | drift first=%+.2f last=%+.2f (accum=%+.2f over %d lines)' % (
            pitch, pitch/20.0, wmean, omean, omean-wmean, drift0, driftN, driftN-drift0, n))
        L.append('    W pitch set=%s  O pitch set=%s' % (sorted(set(round(p,2) for p in wp)), sorted(set(round(p,2) for p in op))))
    txt = '\n'.join(L)
    io.open('c:/tmp/_s520b_out.txt', 'w', encoding='utf-8').write(txt+'\n')
    print(txt)

if __name__ == '__main__':
    main()
