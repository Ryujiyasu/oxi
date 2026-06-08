# -*- coding: utf-8 -*-
"""Extract the <w:p> containing a needle from a document.xml, print its pPr. cp932-safe:
reads UTF-8 docx xml, writes ASCII-tagged result to a file. Also reports enclosing tc/tbl."""
import sys, io, zipfile, re

docx, needle, out = sys.argv[1], sys.argv[2], sys.argv[3]
z = zipfile.ZipFile(docx)
xml = z.read('word/document.xml').decode('utf-8')
idx = xml.find(needle)
L = []
if idx < 0:
    L.append('NEEDLE NOT FOUND: %r' % needle)
else:
    ps = xml.rfind('<w:p>', 0, idx)
    ps2 = xml.rfind('<w:p ', 0, idx)
    ps = max(ps, ps2)
    pe = xml.find('</w:p>', idx) + 6
    para = xml[ps:pe]
    # pPr
    m = re.search(r'<w:pPr>.*?</w:pPr>', para, re.S)
    L.append('=== pPr ===')
    L.append(m.group(0) if m else '(no pPr)')
    # is it inside a tc? find nearest <w:tc> before and </w:tc> after
    tc_before = xml.rfind('<w:tc>', 0, idx)
    tc_close_before = xml.rfind('</w:tc>', 0, idx)
    in_cell = tc_before > tc_close_before
    L.append('=== in_table_cell: %s ===' % in_cell)
    if in_cell:
        tcpr_end = xml.find('</w:tcPr>', tc_before)
        if tcpr_end > 0:
            L.append('=== tcPr ===')
            L.append(xml[tc_before:tcpr_end + 9])
    # full run rPr (first run)
    rm = re.search(r'<w:rPr>.*?</w:rPr>', para, re.S)
    L.append('=== run rPr ===')
    L.append(rm.group(0) if rm else '(no rPr)')
with io.open(out, 'w', encoding='utf-8') as f:
    f.write('\n'.join(L) + '\n')
print('wrote', out)
