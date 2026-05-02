# -*- coding: utf-8 -*-
import sys, zipfile, re
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
with zipfile.ZipFile('tools/golden-test/documents/docx/1ec1091177b1_006.docx') as z:
    th = z.read('word/theme/theme1.xml').decode('utf-8')
# Extract majorFont and minorFont sections
m = re.search(r'<a:majorFont>.*?</a:majorFont>', th, re.DOTALL)
print('=== majorFont ===')
if m: print(m.group(0))
m = re.search(r'<a:minorFont>.*?</a:minorFont>', th, re.DOTALL)
print('\n=== minorFont ===')
if m: print(m.group(0))
