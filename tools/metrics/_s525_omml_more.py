# -*- coding: utf-8 -*-
"""S525 OMML coverage round 2: test MORE structures (delimiter/parentheses, matrix, accent, bar,
n-ary integral, function sin, box/groupChr) for clean STRUCTURAL bugs. Oxi vs Word ink bbox.
cp932-safe: UTF-8 file, ASCII out (no raw math chars to console)."""
import os, io, zipfile, subprocess
import numpy as np
from PIL import Image
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
OUT = os.path.join(ROOT, 'tools', 'golden-test', 'repros', 'omml')
EXE = os.path.join(ROOT, 'tools', 'oxi-dwrite-renderer', 'target', 'release', 'oxi-dwrite-renderer.exe')
os.makedirs(OUT, exist_ok=True)
WNS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
MNS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')

def r(t): return '<m:r><m:t>%s</m:t></m:r>' % t

# NB: use LITERAL operator/accent chars (∫ ∑ etc.), NOT XML numeric entities —
# Oxi's parser reads the m:chr/@val attribute raw (does not entity-decode), and
# real Word docs always write the literal char. With "&#8747;" the op parses as
# '&' and the integral-vs-sum limit logic misfires (S525 repro lesson).
STRUCT = {
    'delim':  '<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>%s</m:e></m:d>' % (
              '<m:f><m:num>%s</m:num><m:den>%s</m:den></m:f>' % (r('a'), r('b'))),
    'matrix': '<m:m><m:mr><m:e>%s</m:e><m:e>%s</m:e></m:mr><m:mr><m:e>%s</m:e><m:e>%s</m:e></m:mr></m:m>' % (r('1'), r('2'), r('3'), r('4')),
    'accent': '<m:acc><m:accPr><m:chr m:val="̂"/></m:accPr><m:e>%s</m:e></m:acc>' % r('x'),  # combining circumflex
    'bar':    '<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e>%s</m:e></m:bar>' % r('AB'),  # overline
    'integral':'<m:nary><m:naryPr><m:chr m:val="∫"/></m:naryPr><m:sub>%s</m:sub><m:sup>%s</m:sup><m:e>%s</m:e></m:nary>' % (r('0'), r('1'), '<m:f><m:num>%s</m:num><m:den>%s</m:den></m:f>'%(r('1'),r('x'))),
    'sum':    '<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub>%s</m:sub><m:sup>%s</m:sup><m:e>%s</m:e></m:nary>' % (r('i=1'), r('n'), r('i')),
    'func':   '<m:func><m:fName><m:r><m:t>sin</m:t></m:r></m:fName><m:e>%s</m:e></m:func>' % r('x'),
    'group':  '<m:groupChr><m:groupChrPr><m:chr m:val="⏟"/><m:pos m:val="bot"/></m:groupChrPr><m:e>%s</m:e></m:groupChr>' % r('xyz'),  # under-brace
}

def build(name, omml):
    para = '<w:p><m:oMathPara %s><m:oMath>%s</m:oMath></m:oMathPara></w:p>' % (MNS, omml)
    sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418"/></w:sectPr>'
    doc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document %s %s><w:body>%s%s</w:body></w:document>' % (WNS, MNS, para, sect)
    p = os.path.join(OUT, name)
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RELS); z.writestr('word/document.xml', doc)
    return p

def ink_bbox(png):
    if not os.path.exists(png): return None
    im = np.asarray(Image.open(png).convert('L'), dtype=np.float32)
    dark = im < 128
    if not dark.any(): return (0, 0, 0, 0, 0)
    ys, xs = np.where(dark)
    return (int(xs.min()), int(ys.min()), int(xs.max()), int(ys.max()), int(dark.sum()))

def word_png(docx):
    import win32com.client, pythoncom, fitz
    pdf = os.path.splitext(docx)[0] + '.pdf'; png = os.path.splitext(docx)[0] + '_w.png'
    pythoncom.CoInitialize(); w = win32com.client.DispatchEx('Word.Application'); w.Visible = False
    try:
        d = w.Documents.Open(os.path.abspath(docx), ReadOnly=True); d.ExportAsFixedFormat(pdf, 17); d.Close(False)
    finally:
        w.Quit()
    doc = fitz.open(pdf); doc[0].get_pixmap(matrix=fitz.Matrix(150/72, 150/72)).save(png); doc.close()
    return png

def main():
    L = ['S525 OMML round2: Oxi vs Word ink bbox (x0,y0,x1,y1,dark) @150dpi']
    for name, omml in STRUCT.items():
        dx = build(name + '.docx', omml)
        opng = os.path.join('c:/tmp', 's525_' + name)
        subprocess.run([EXE, os.path.abspath(dx), opng, '150'], capture_output=True, text=True)
        ob = ink_bbox(opng + '_p1.png'); wb = ink_bbox(word_png(dx))
        verdict = ''
        if ob is None or ob[4] < 5:
            verdict = ' <<< OXI EMPTY/MISSING'
        elif wb and wb[4] >= 5:
            ow, oh = ob[2]-ob[0], ob[3]-ob[1]; ww, wh = wb[2]-wb[0], wb[3]-wb[1]
            if ww > 0 and wh > 0:
                if abs(ow-ww) > max(8, ww*0.4) or abs(oh-wh) > max(8, wh*0.4):
                    verdict = ' <<< SIZE MISMATCH O %dx%d W %dx%d' % (ow, oh, ww, wh)
        L.append('%-9s Word=%s Oxi=%s%s' % (name, str(wb), str(ob), verdict))
    txt = '\n'.join(L)
    io.open('c:/tmp/_s525_out.txt', 'w', encoding='utf-8').write(txt + '\n')
    print(txt)

if __name__ == '__main__':
    main()
