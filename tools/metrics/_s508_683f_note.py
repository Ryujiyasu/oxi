# -*- coding: utf-8 -*-
"""S508 683f note-para indent: find the (注) note paragraph (justify-compressed ~9pt by Oxi)
and print its pPr indent (left/right/firstLine/hanging) + jc + the enclosing tc width, to
find why Oxi justifies it ~9pt narrower than Word. cp932-safe: UTF-8 needle in-file."""
import io, zipfile, re, glob, os
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
dx = glob.glob(os.path.join(ROOT, 'tools/golden-test/documents/docx', '683ffcab86e2*.docx'))[0]
xml = zipfile.ZipFile(dx).read('word/document.xml').decode('utf-8')
NEEDLES = ['追加することで', '仕様書に従って', '本府省']
L = ['S508 683f note-para indent  docx=%s' % os.path.basename(dx)]
m = re.search(r'<w:sectPr.*?</w:sectPr>', xml, re.S)
pg = re.search(r'<w:pgMar [^>]*>', m.group(0)) if m else None
L.append('pgMar: %s' % (pg.group(0) if pg else '?'))
dg = re.search(r'<w:docGrid [^>]*>', xml)
L.append('docGrid: %s' % (dg.group(0) if dg else 'none'))
found = False
for nd in NEEDLES:
    idx = xml.find(nd)
    if idx < 0:
        L.append('needle NOT found: %s' % nd); continue
    ps = max(xml.rfind('<w:p>', 0, idx), xml.rfind('<w:p ', 0, idx))
    pe = xml.find('</w:p>', idx) + 6
    para = xml[ps:pe]
    ppr = re.search(r'<w:pPr>(.*?)</w:pPr>', para, re.S)
    ind = re.search(r'<w:ind [^>]*>', para)
    jc = re.search(r'<w:jc w:val="([^"]+)"', para)
    txt = ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', para, re.S))
    # enclosing tcW
    tcstart = xml.rfind('<w:tc>', 0, idx)
    tcw = re.search(r'<w:tcW w:w="(\d+)"', xml[tcstart:idx]) if tcstart >= 0 else None
    L.append('\n=== needle %s ===' % nd)
    L.append('  jc=%s  ind=%s' % (jc.group(1) if jc else '(default)', ind.group(0) if ind else 'NONE'))
    L.append('  tcW=%s' % (tcw.group(1) if tcw else '?'))
    L.append('  pPr=%s' % (ppr.group(0)[:200] if ppr else 'none'))
    L.append('  text=%s' % txt[:48])
    found = True
    break
if not found:
    L.append('NO needle matched')
with io.open('c:/tmp/_s508_out.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(L) + '\n')
print('wrote c:/tmp/_s508_out.txt')
