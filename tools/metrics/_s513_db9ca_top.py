# -*- coding: utf-8 -*-
"""S513 db9ca top paras: dump the first ~5 body paragraphs' pPr (sz, spacing, lineRule,
the empty paras) + docGrid + docDefaults, to find why the empty para between the title and
body is 20pt in Word vs 18pt in Oxi. cp932-safe."""
import io, zipfile, re, glob, os
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
dx = (glob.glob(os.path.join(ROOT, 'tools/golden-test/documents/docx', 'db9ca18368cd*.docx'))
      or glob.glob(os.path.join(ROOT, 'pipeline_data/docx', 'db9ca18368cd*.docx')))[0]
z = zipfile.ZipFile(dx)
xml = z.read('word/document.xml').decode('utf-8')
L = ['S513 db9ca top paras  docx=%s' % os.path.basename(dx)]
dg = re.search(r'<w:docGrid [^>]*>', xml); L.append('docGrid: %s' % (dg.group(0) if dg else 'none'))
m = re.search(r'<w:pgMar [^>]*>', xml); L.append('pgMar: %s' % (m.group(0) if m else '?'))
# docDefaults spacing + rPr sz
try:
    st = z.read('word/styles.xml').decode('utf-8')
    dd = re.search(r'<w:docDefaults>.*?</w:docDefaults>', st, re.S)
    if dd:
        sp = re.search(r'<w:spacing [^>]*>', dd.group(0)); sz = re.search(r'<w:sz w:val="(\d+)"', dd.group(0))
        L.append('docDefaults spacing: %s  sz=%s' % (sp.group(0) if sp else 'none', sz.group(1) if sz else '?'))
    # Normal style spacing
    nm = re.search(r'<w:style [^>]*w:styleId="a?"[^>]*>.*?</w:style>', st, re.S)
except Exception as e:
    L.append('styles err %s' % e)
# first ~6 paragraphs
paras = re.findall(r'<w:p\b[^>]*>.*?</w:p>', xml, re.S)[:6]
for i, p in enumerate(paras):
    body = re.search(r'<w:pPr>(.*?)</w:pPr>', p, re.S)
    ppr = body.group(1) if body else ''
    sz = re.search(r'<w:sz w:val="(\d+)"', p)
    spacing = re.search(r'<w:spacing [^>]*>', ppr)
    line = re.search(r'<w:spacing[^>]*w:line="(\d+)"[^>]*w:lineRule="([^"]+)"', ppr)
    sng = '<w:snapToGrid w:val="0"' in ppr
    pstyle = re.search(r'<w:pStyle w:val="([^"]+)"', ppr)
    txt = ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', p, re.S))
    L.append('--- para %d sz=%s pStyle=%s snapGrid0=%s spacing=%s text=%s' % (
        i, sz.group(1) if sz else '(inherit)', pstyle.group(1) if pstyle else '-',
        sng, spacing.group(0) if spacing else 'none', (txt[:24] if txt else '(EMPTY)')))
with io.open('c:/tmp/_s513_out.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(L) + '\n')
print('\n'.join(L))
