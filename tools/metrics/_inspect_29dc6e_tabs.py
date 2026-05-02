# -*- coding: utf-8 -*-
"""Inspect 29dc6e's tab usage to characterize tab patterns."""
import sys, os, zipfile, re
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

DOCX = r'C:\Users\ryuji\oxi-4\tools\golden-test\documents\docx\29dc6e8943fe_order_01.docx'

with zipfile.ZipFile(DOCX) as z:
    doc = z.read('word/document.xml').decode('utf-8')

# Find all tab elements in run content (not pPr definitions)
print("=== Tab characters (<w:tab/>) in runs ===")
for i, m in enumerate(re.finditer(r'<w:tab/>', doc)):
    pos = m.start()
    # Find enclosing <w:p
    p_start = max(doc.rfind('<w:p ', 0, pos), doc.rfind('<w:p>', 0, pos))
    p_end = doc.find('</w:p>', pos) + len('</w:p>')
    para = doc[p_start:p_end]
    text = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', para))[:60]
    # tab definitions
    tabs_def = re.search(r'<w:tabs>(.*?)</w:tabs>', para, re.DOTALL)
    tabs_str = tabs_def.group(1)[:200] if tabs_def else 'none'
    # ind
    ind = re.search(r'<w:ind[^/>]*?/?>', para)
    ind_str = ind.group(0) if ind else 'no ind'
    # jc
    jc = re.search(r'<w:jc[^/>]*?w:val="([^"]+)"', para)
    print(f"\n[{i+1}] pos={pos}")
    print(f"  text: {text!r}")
    print(f"  tabs def: {tabs_str}")
    print(f"  ind: {ind_str}")
    print(f"  jc: {jc.group(1) if jc else 'default'}")

# Also: explicit tab definitions with position, alignment
print("\n=== Tab stop definitions (in pPr <w:tabs>) ===")
for m in re.finditer(r'<w:tabs>(.*?)</w:tabs>', doc, re.DOTALL):
    tabs_inner = m.group(1)
    # Each tab def
    tab_defs = re.findall(r'<w:tab\s+([^/>]+?)/>', tabs_inner)
    for td in tab_defs[:5]:
        print(f"  tab: {td}")
