# -*- coding: utf-8 -*-
"""Inspect 0e7af's <w:pBdr> usage."""
import sys, os, glob, zipfile, re
from collections import Counter
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

paths = glob.glob(r'C:\Users\ryuji\oxi-4\tools\golden-test\documents\docx\0e7af*')
DOCX = paths[0]
print(f"Inspecting: {os.path.basename(DOCX)}\n")

with zipfile.ZipFile(DOCX) as z:
    doc = z.read('word/document.xml').decode('utf-8')

# Find all <w:pBdr> blocks and what they contain
pBdr_matches = re.findall(r'<w:pBdr>(.*?)</w:pBdr>', doc, re.DOTALL)
print(f"Total <w:pBdr> blocks: {len(pBdr_matches)}")

# Categorize by which sides have borders
patterns = Counter()
configs = Counter()
for block in pBdr_matches:
    sides = []
    if '<w:top ' in block: sides.append('top')
    if '<w:bottom ' in block: sides.append('bottom')
    if '<w:left ' in block: sides.append('left')
    if '<w:right ' in block: sides.append('right')
    if '<w:between ' in block: sides.append('between')
    patterns['+'.join(sides) if sides else 'none'] += 1
    # Get full configs for first instance of each pattern
    sides_key = '+'.join(sides)
    if configs.get(sides_key, 0) < 2:  # show 2 examples per pattern
        configs[sides_key] += 1
        # Extract one border attrs
        bm = re.search(r'<w:(top|bottom|left|right|between)\s+([^/>]*?)/>', block)
        if bm:
            print(f"  pattern={sides_key}: example {bm.group(0)}")

print(f"\nBorder patterns:")
for p, c in patterns.most_common():
    print(f"  {p}: {c}")

# Check consecutive bordered paragraphs (shared borders likely)
# Find paragraphs with pBdr in order
print("\n=== Position of pBdr paragraphs ===")
for i, m in enumerate(re.finditer(r'<w:pBdr>', doc)):
    pos = m.start()
    p_start = max(doc.rfind('<w:p ', 0, pos), doc.rfind('<w:p>', 0, pos))
    p_end = doc.find('</w:p>', pos) + len('</w:p>')
    para = doc[p_start:p_end]
    text = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', para))[:40]
    if i < 10:
        print(f"  [{i+1}] pos={pos}: text={text!r}")
