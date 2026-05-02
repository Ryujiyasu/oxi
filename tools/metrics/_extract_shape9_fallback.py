# -*- coding: utf-8 -*-
"""Extract 1ec1's Shape 9 mc:Fallback block to file for embedding in V_R."""
import sys, os, zipfile, re
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

with zipfile.ZipFile('tools/golden-test/documents/docx/1ec1091177b1_006.docx') as z:
    doc = z.read('word/document.xml').decode('utf-8')
BOX5 = 84340
ac_starts = [m.start() for m in re.finditer(r'<mc:AlternateContent>', doc)]
ac_ends = [m.end() for m in re.finditer(r'</mc:AlternateContent>', doc)]
for s in reversed(ac_starts):
    if s < BOX5:
        for e in ac_ends:
            if e > BOX5 and e > s:
                ac = doc[s:e]
                fb = re.search(r'<mc:Fallback>(.*?)</mc:Fallback>', ac, re.DOTALL)
                if fb:
                    out_path = 'pipeline_data/1ec1_shape9_fallback.xml'
                    with open(out_path, 'w', encoding='utf-8') as f:
                        f.write(fb.group(0))
                    print(f"Saved {len(fb.group(0))} chars to {out_path}")
                    print("First 500 chars:")
                    print(fb.group(0)[:500])
                else:
                    print("No mc:Fallback found")
                break
        break
