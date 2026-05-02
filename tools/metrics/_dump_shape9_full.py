# -*- coding: utf-8 -*-
"""Dump Shape 9 full txbxContent: all paragraphs in order, properties."""
import sys, zipfile, re
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

with zipfile.ZipFile('tools/golden-test/documents/docx/1ec1091177b1_006.docx') as z:
    doc = z.read('word/document.xml').decode('utf-8')

# Find Shape 9's mc:AlternateContent block (the second instance with id=9; the first matters for BOX[5])
# Actually we want the AlternateContent that contains BOX[5] = pos 84340
BOX5 = 84340
ac_starts = [m.start() for m in re.finditer(r'<mc:AlternateContent>', doc)]
ac_ends = [m.end() for m in re.finditer(r'</mc:AlternateContent>', doc)]
ac_start = None
ac_end = None
for s in reversed(ac_starts):
    if s < BOX5:
        for e in ac_ends:
            if e > BOX5 and e > s:
                ac_start = s; ac_end = e
                break
        break
print(f"Shape AC range: [{ac_start}, {ac_end}]")
shape_xml = doc[ac_start:ac_end]

# Find txbxContent inside
txbx_m = re.search(r'<w:txbxContent>(.*?)</w:txbxContent>', shape_xml, re.DOTALL)
txbx = txbx_m.group(1) if txbx_m else ''
print(f"txbxContent length: {len(txbx)}")

# Find all <w:p> at top level
paras = re.findall(r'<w:p\s[^>]*>.*?</w:p>|<w:p>.*?</w:p>', txbx, re.DOTALL)
print(f"Paragraphs in txbxContent: {len(paras)}")

for i, p in enumerate(paras):
    text = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', p))[:40]
    ppr_m = re.search(r'<w:pPr>(.*?)</w:pPr>', p, re.DOTALL)
    pPr = ppr_m.group(1) if ppr_m else ''
    # Extract relevant pPr items
    snap = re.search(r'<w:snapToGrid[^/>]*/?>', pPr)
    spc = re.search(r'<w:spacing[^/>]*/?>', pPr)
    ind = re.search(r'<w:ind[^/>]*/?>', pPr)
    jc = re.search(r'<w:jc[^/>]*/?>', pPr)
    pStyle = re.search(r'<w:pStyle[^/>]*?w:val="([^"]+)"', pPr)
    print(f"\n[P{i+1}] text={text!r}")
    print(f"  pStyle={pStyle.group(1) if pStyle else None}")
    print(f"  snap={snap.group(0) if snap else None}")
    print(f"  spacing={spc.group(0) if spc else None}")
    print(f"  ind={ind.group(0) if ind else None}")
    print(f"  jc={jc.group(0) if jc else None}")

# Also dump bodyPr in full
bodypr = re.search(r'<wps:bodyPr[^/>]*?/?>', shape_xml)
print(f"\nbodyPr: {bodypr.group(0) if bodypr else None}")
