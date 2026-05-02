# -*- coding: utf-8 -*-
import sys, zipfile, re

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
BOX = '□'

with zipfile.ZipFile('tools/golden-test/documents/docx/1ec1091177b1_006.docx') as z:
    doc = z.read('word/document.xml').decode('utf-8')
    styles = z.read('word/styles.xml').decode('utf-8')
    settings = z.read('word/settings.xml').decode('utf-8')
    numbering = ''
    try:
        numbering = z.read('word/numbering.xml').decode('utf-8')
    except KeyError:
        pass

# docDefaults
print("=== docDefaults pPrDefault ===")
m = re.search(r'<w:docDefaults>.*?</w:docDefaults>', styles, re.DOTALL)
if m:
    pprd = re.search(r'<w:pPrDefault>(.*?)</w:pPrDefault>', m.group(0), re.DOTALL)
    if pprd:
        print(pprd.group(0)[:800])
    rprd = re.search(r'<w:rPrDefault>(.*?)</w:rPrDefault>', m.group(0), re.DOTALL)
    if rprd:
        print(rprd.group(0)[:400])

# Normal style
print("\n=== Normal style ===")
m = re.search(r'<w:style[^>]*w:styleId="a"[^>]*>.*?</w:style>', styles, re.DOTALL)
if m:
    print(m.group(0)[:800])

# Settings: defaultTabStop, autoSpaceDE/DN, compatibility
print("\n=== settings.xml flags ===")
for tag in ['w:defaultTabStop', 'w:characterSpacingControl', 'w:adjustRightInd', 'w:doNotShadeFormData', 'w:compat']:
    for mm in re.finditer(r'<' + re.escape(tag) + r'[^/>]*?/?>', settings):
        print(mm.group(0))

# numbering check
if numbering:
    print(f"\n=== numbering.xml exists ({len(numbering)} bytes)")
else:
    print("\n=== numbering.xml MISSING")

# Look for first BOX paragraph context — pos 27204 was □１ in body
pos = 27204
p_start = doc.rfind('<w:p ', 0, pos)
if p_start < 0:
    p_start = doc.rfind('<w:p>', 0, pos)
p_end = doc.find('</w:p>', pos) + len('</w:p>')
print(f"\n=== First BOX paragraph (pos={pos}) full XML ===")
print(doc[p_start:p_end])
