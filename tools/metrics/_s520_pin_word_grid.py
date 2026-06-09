# -*- coding: utf-8 -*-
"""S520 KEYSTONE: pin Word's exact vertical baseline rounding grid. Build a clean column of N
identical CJK lines (MS Mincho 10.5pt, single spacing, NO docGrid), export via Word PDF, read
each line's baseline to full precision, and analyze: (1) consecutive pitch values, (2) does each
absolute baseline fit a device grid g (search g in {1/96,1/120,1/150,1/96*..., 0.05,0.1,0.25,0.75pt
and twip}) — i.e. is (baseline_pt * dpi) near-integer? Report best-fitting grid. cp932-safe."""
import os, io, zipfile
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT = os.path.join(ROOT, 'tools', 'golden-test', 'repros', 'gridpin')
os.makedirs(OUT, exist_ok=True)
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')

def build(name, sz, n, line_tw=0, docgrid=False):
    MIN = 'ＭＳ 明朝'  # "ＭＳ 明朝" MS Mincho
    rpr = '<w:rPr><w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s"/><w:sz w:val="%d"/></w:rPr>' % (MIN, MIN, MIN, sz)
    ppr = ('<w:pPr><w:spacing w:line="%d" w:lineRule="exact"/></w:pPr>' % line_tw) if line_tw else ''
    paras = ''.join('<w:p>%s<w:r>%s<w:t>あ%d</w:t></w:r></w:p>' % (ppr, rpr, i) for i in range(n))
    grid = '<w:docGrid w:type="lines" w:linePitch="360"/>' if docgrid else ''
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="851" w:footer="397"/>%s</w:sectPr>' % grid)
    doc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document %s><w:body>%s%s</w:body></w:document>' % (NS, paras, sect)
    p = os.path.join(OUT, name)
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RELS); z.writestr('word/document.xml', doc)
    return p

def word_baselines(docx):
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

def analyze(bl, label, L):
    L.append('=== %s : %d lines' % (label, len(bl)))
    pitches = [round(bl[i+1]-bl[i], 4) for i in range(len(bl)-1)]
    L.append('  pitch set: %s' % sorted(set(round(p,3) for p in pitches)))
    L.append('  mean pitch: %.5f' % (sum(pitches)/len(pitches)) if pitches else '  (1 line)')
    # grid fit: for each candidate grid, residual of baseline*dpi from nearest integer
    cands = [('twip 1/20', 20.0), ('96dpi 0.75pt', 96/72.0), ('120dpi', 120/72.0), ('150dpi', 150/72.0),
             ('0.05pt', 20.0), ('0.1pt', 10.0), ('0.25pt', 4.0), ('300dpi', 300/72.0), ('600dpi', 600/72.0),
             ('0.01pt', 100.0)]
    for nm, units in cands:
        res = [abs((b*units) - round(b*units)) for b in bl]
        L.append('  grid %-14s maxres=%.4f meanres=%.4f' % (nm, max(res), sum(res)/len(res)))

def main():
    L = ['S520 pin Word vertical baseline grid (MS Mincho)']
    for name, sz, n, ltw, dg, lab in [
        ('col_105_single.docx', 21, 30, 0, False, 'MS Mincho 10.5pt single, no grid'),
        ('col_105_exact270.docx', 21, 30, 270, False, 'MS Mincho 10.5pt exact line=270(13.5pt)'),
        ('col_105_docgrid.docx', 21, 30, 0, True, 'MS Mincho 10.5pt docGrid lines 360'),
    ]:
        dx = build(name, sz, n, ltw, dg)
        bl = word_baselines(dx)
        if bl:
            analyze(bl, lab, L)
            L.append('  first 6 baselines: %s' % [round(b,4) for b in bl[:6]])
    txt = '\n'.join(L)
    io.open('c:/tmp/_s520_out.txt', 'w', encoding='utf-8').write(txt+'\n')
    print(txt)

if __name__ == '__main__':
    main()
