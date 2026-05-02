# -*- coding: utf-8 -*-
"""Inspect 4 vertical-writing docs to understand what features they use."""
import sys, os, zipfile, re
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

VERTICAL_DOCS = [
    '2ea81a8441cc_0025006-192.docx',
    '459f05f1e877_kyodokenkyuyoushiki01.docx',
    '7ead52b63f0e_000067058.docx',
    'ed025cbecffb_index-23.docx',
]

DOCX_DIR = r'C:\Users\ryuji\oxi-4\tools\golden-test\documents\docx'

for fname in VERTICAL_DOCS:
    path = os.path.join(DOCX_DIR, fname)
    print(f"\n========== {fname} ==========")
    with zipfile.ZipFile(path) as z:
        doc = z.read('word/document.xml').decode('utf-8', errors='replace')
    # Find textDirection elements
    td_matches = list(re.finditer(r'<w:textDirection[^/>]*?w:val="([^"]+)"', doc))
    print(f"textDirection occurrences: {len(td_matches)}")
    seen_vals = set()
    for m in td_matches:
        v = m.group(1)
        if v in seen_vals: continue
        seen_vals.add(v)
        # Find context — is it in <w:sectPr>, <w:pPr>, <w:tcPr>?
        before = doc[max(0, m.start()-200):m.start()]
        after = doc[m.start():m.start()+200]
        ctx = "?"
        if '<w:sectPr' in before[-150:]:
            ctx = 'sectPr'
        elif '<w:tcPr' in before[-150:]:
            ctx = 'tcPr'
        elif '<w:pPr' in before[-150:]:
            ctx = 'pPr'
        print(f"  val={v}, context={ctx}")
    # Show vertical layout features
    if 'eaVert' in doc:
        print(f"  eaVert occurrences: {doc.count('eaVert')}")
    # Section properties
    sectprs = re.findall(r'<w:sectPr[^>]*>.*?</w:sectPr>', doc, re.DOTALL)
    print(f"  Section breaks: {len(sectprs)}")
    for i, sp in enumerate(sectprs[:3]):
        pgsz = re.search(r'<w:pgSz[^/>]*/?>', sp)
        td = re.search(r'<w:textDirection[^/>]*?w:val="([^"]+)"', sp)
        if td:
            print(f"  sectPr[{i+1}] vert: {td.group(0)}, pgSz: {pgsz.group(0) if pgsz else 'none'}")
