# -*- coding: utf-8 -*-
import sys,glob,zipfile,re,os
sys.stdout.reconfigure(encoding='utf-8')
hits=[]
for f in glob.glob('tools/golden-test/documents/docx/*.docx'):
    try:
        z=zipfile.ZipFile(f); x=z.read('word/document.xml').decode('utf-8','replace')
    except Exception: continue
    nums=re.findall(r'<w:cols[^>]*?w:num="(\d+)"',x)
    multi=[n for n in nums if int(n)>=2]
    if multi: hits.append((os.path.basename(f), multi))
print('docs with multi-column sections (w:cols num>=2):')
for h in hits: print('  ',h[0], h[1])
print('total multi-col docs:', len(hits), '/ 269')
