# -*- coding: utf-8 -*-
"""Side-by-side diff of V_L8 docx Shape vs 1ec1's Shape 9 block."""
import sys, zipfile, re
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# Read 1ec1 Shape 9 AC block
with zipfile.ZipFile('tools/golden-test/documents/docx/1ec1091177b1_006.docx') as z:
    doc1 = z.read('word/document.xml').decode('utf-8')

# Find AC containing BOX[5] = pos 84340
ac_starts = [m.start() for m in re.finditer(r'<mc:AlternateContent>', doc1)]
ac_ends = [m.end() for m in re.finditer(r'</mc:AlternateContent>', doc1)]
ac_1ec1 = None
for s in reversed(ac_starts):
    if s < 84340:
        for e in ac_ends:
            if e > 84340 and e > s:
                ac_1ec1 = doc1[s:e]
                break
        break

# Read V_L8 docx
with zipfile.ZipFile('pipeline_data/1ec1_shape35/V_L8_shape9_with_ind_105.docx') as z:
    doc2 = z.read('word/document.xml').decode('utf-8')
ac_v_l8 = None
for s in [m.start() for m in re.finditer(r'<mc:AlternateContent>', doc2)]:
    e = doc2.find('</mc:AlternateContent>', s) + len('</mc:AlternateContent>')
    if e > s:
        ac_v_l8 = doc2[s:e]
        break

print("=== 1ec1 Shape 9 AC block ===")
print(ac_1ec1[:3000])
print("\n=== V_L8 AC block ===")
print(ac_v_l8[:3000])
