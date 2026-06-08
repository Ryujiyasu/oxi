# -*- coding: utf-8 -*-
"""S505 b837 footnote width: dump docGrid + footnote 14's pPr (indent) + footnote default
style, to see why Oxi fits ~2 more fs11 chars/line than Word (1-line cascade). cp932-safe."""
import io, zipfile, re, glob, os
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
dx = glob.glob(os.path.join(ROOT, 'tools/golden-test/documents/docx', 'b837808d0555*.docx'))
dx = dx[0] if dx else glob.glob(os.path.join(ROOT, 'pipeline_data/docx', 'b837808d0555*.docx'))[0]
z = zipfile.ZipFile(dx)
L = ['S505 b837 footnote width  docx=%s' % os.path.basename(dx)]
doc = z.read('word/document.xml').decode('utf-8')
m = re.search(r'<w:docGrid [^>]*>', doc); L.append('docGrid: %s' % (m.group(0) if m else '?'))
m = re.search(r'<w:pgMar [^>]*>', doc); L.append('pgMar: %s' % (m.group(0) if m else '?'))
# footnotes.xml
try:
    fn = z.read('word/footnotes.xml').decode('utf-8')
    # find the footnote containing the over-fit text
    needle = '公開するデータに関して'
    idx = fn.find(needle)
    if idx >= 0:
        ps = max(fn.rfind('<w:p>', 0, idx), fn.rfind('<w:p ', 0, idx))
        pe = fn.find('</w:p>', idx) + 6
        para = fn[ps:pe]
        ppr = re.search(r'<w:pPr>(.*?)</w:pPr>', para, re.S)
        L.append('footnote14 pPr: %s' % (ppr.group(0) if ppr else '(none)'))
        ind = re.search(r'<w:ind [^>]*>', para); L.append('footnote14 ind: %s' % (ind.group(0) if ind else 'none'))
        sz = re.search(r'<w:sz w:val="(\d+)"', para); L.append('footnote14 sz: %s' % (sz.group(1) if sz else '?'))
    else:
        L.append('footnote needle not found; footnote count=%d' % fn.count('<w:footnote '))
except Exception as e:
    L.append('footnotes read err: %s' % e)
# footnote style indent from styles.xml
try:
    st = z.read('word/styles.xml').decode('utf-8')
    fm = re.search(r'<w:style[^>]*w:styleId="[^"]*[Ff]ootnote[^"]*"[^>]*>.*?</w:style>', st, re.S)
    if fm:
        ind = re.search(r'<w:ind [^>]*>', fm.group(0))
        L.append('footnote STYLE ind: %s' % (ind.group(0) if ind else 'none'))
except Exception as e:
    L.append('styles err: %s' % e)
with io.open('c:/tmp/_s505_b837_fn_out.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(L) + '\n')
print('wrote c:/tmp/_s505_b837_fn_out.txt')
