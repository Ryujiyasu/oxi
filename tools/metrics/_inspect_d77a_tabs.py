# -*- coding: utf-8 -*-
"""Inspect d77a's actual <w:tab/> usages."""
import sys, os, zipfile, re, glob
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

paths = glob.glob(r'C:\Users\ryuji\oxi-4\tools\golden-test\documents\docx\d77a*')
DOCX = paths[0]
print(f"Inspecting: {os.path.basename(DOCX)}\n")

with zipfile.ZipFile(DOCX) as z:
    doc = z.read('word/document.xml').decode('utf-8')
    settings = z.read('word/settings.xml').decode('utf-8', errors='replace')

# Default tab stop
m = re.search(r'<w:defaultTabStop[^/>]*?w:val="(\d+)"', settings)
default_tab = int(m.group(1)) if m else 720
print(f"defaultTabStop: {default_tab} twips = {default_tab/20:.2f}pt")

# Each <w:tab/> in run content
print("\n=== <w:tab/> usages ===")
for i, m in enumerate(re.finditer(r'<w:tab/>', doc)):
    pos = m.start()
    p_start = max(doc.rfind('<w:p ', 0, pos), doc.rfind('<w:p>', 0, pos))
    p_end = doc.find('</w:p>', pos) + len('</w:p>')
    para = doc[p_start:p_end]
    text = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', para))
    # Pre-tab text (text before this tab)
    pre_pos = pos - p_start
    pre_para = para[:pre_pos]
    pre_text = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', pre_para))
    # Post-tab text
    post_pos = pos + len('<w:tab/>') - p_start
    post_para = para[post_pos:]
    post_text = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', post_para))
    # ind / jc
    ppr = re.search(r'<w:pPr>(.*?)</w:pPr>', para, re.DOTALL)
    pPr_str = ppr.group(0) if ppr else ''
    ind = re.search(r'<w:ind[^/>]*?/?>', pPr_str)
    jc = re.search(r'<w:jc[^/>]*?w:val="([^"]+)"', pPr_str)
    explicit_tabs = re.search(r'<w:tabs>.*?</w:tabs>', pPr_str, re.DOTALL)
    print(f"\n[{i+1}] pos={pos}")
    print(f"  full para text: {text[:80]!r}")
    print(f"  pre-tab text:   {pre_text!r}")
    print(f"  post-tab text:  {post_text[:50]!r}")
    print(f"  ind: {ind.group(0) if ind else 'no'}")
    print(f"  jc: {jc.group(1) if jc else 'default'}")
    print(f"  explicit tabs: {'YES' if explicit_tabs else 'NO'}")
