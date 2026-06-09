# -*- coding: utf-8 -*-
"""S524 coverage: author minimal OMML (math) repros, render Oxi (dwrite PNG) vs Word (PDF->PNG),
and report per-structure whether Oxi produces output at all + a crude visual size/position compare.
OMML is greenfield-dormant in Oxi (0 corpus docs). Structures: fraction, superscript, subscript,
radical, n-ary sum, nested. cp932-safe: UTF-8 file, results to file, ASCII out."""
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

def r(t):
    return '<m:r><m:t>%s</m:t></m:r>' % t

STRUCT = {
    'frac':  '<m:f><m:num>%s</m:num><m:den>%s</m:den></m:f>' % (r('a'), r('b')),
    'sup':   '<m:sSup><m:e>%s</m:e><m:sup>%s</m:sup></m:sSup>' % (r('x'), r('2')),
    'sub':   '<m:sSub><m:e>%s</m:e><m:sub>%s</m:sub></m:sSub>' % (r('x'), r('i')),
    'rad':   '<m:rad><m:deg/><m:e>%s</m:e></m:rad>' % r('x'),
    'nary':  '<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub>%s</m:sub><m:sup>%s</m:sup><m:e>%s</m:e></m:nary>' % (r('i=1'), r('n'), r('i')),
    'nested':'<m:f><m:num><m:sSup><m:e>%s</m:e><m:sup>%s</m:sup></m:sSup></m:num><m:den>%s</m:den></m:f>' % (r('x'), r('2'), r('y')),
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
    if not os.path.exists(png):
        return None
    im = np.asarray(Image.open(png).convert('L'), dtype=np.float32)
    dark = im < 128
    if not dark.any():
        return (0, 0, 0, 0, 0)
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
    L = ['S524 OMML coverage: Oxi vs Word ink bbox (x0,y0,x1,y1,darkpx) @150dpi for each structure']
    for name, omml in STRUCT.items():
        dx = build(name + '.docx', omml)
        opng = os.path.join('c:/tmp', 's524_' + name)
        subprocess.run([EXE, os.path.abspath(dx), opng, '150'], capture_output=True, text=True)
        ob = ink_bbox(opng + '_p1.png')
        wb = ink_bbox(word_png(dx))
        L.append('--- %-7s' % name)
        L.append('   Word: %s' % str(wb))
        L.append('   Oxi : %s  %s' % (str(ob), 'EMPTY/MISSING' if (ob is None or ob[4] < 5) else ''))
        if ob and wb and ob[4] >= 5:
            ow, oh = ob[2]-ob[0], ob[3]-ob[1]; ww, wh = wb[2]-wb[0], wb[3]-wb[1]
            L.append('   size Oxi %dx%d vs Word %dx%d  (dark Oxi %d / Word %d ratio %.2f)' % (
                ow, oh, ww, wh, ob[4], wb[4], ob[4]/max(wb[4],1)))
    txt = '\n'.join(L)
    io.open('c:/tmp/_s524_out.txt', 'w', encoding='utf-8').write(txt + '\n')
    print(txt)

if __name__ == '__main__':
    main()
