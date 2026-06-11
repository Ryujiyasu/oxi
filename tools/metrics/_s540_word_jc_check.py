# -*- coding: utf-8 -*-
"""S540: for regression-doc paragraphs with EXPLICIT pPr jc=left in the XML,
ask Word COM for the RESOLVED ParagraphFormat.Alignment (0=left 3=justify).
Decides whether the has_explicit_jc fix direction matches Word per doc.
Writes c:/tmp/s540_jc_check.txt.
"""
import glob, io, os, re, subprocess, sys, zipfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DOCS = os.path.join(ROOT, 'tools', 'golden-test', 'documents', 'docx')
STEMS = ['7f272a', 'fded6', '34140b', 'de6e32', '04b88e']

out = io.open('c:/tmp/s540_jc_check.txt', 'w', encoding='utf-8')

WORD = r'''
# -*- coding: utf-8 -*-
import sys, io, json
import win32com.client, pythoncom
pythoncom.CoInitialize()
word = win32com.client.Dispatch('Word.Application'); word.Visible=False; word.DisplayAlerts=False
word.AutomationSecurity = 3
doc = word.Documents.Open(sys.argv[1], ReadOnly=True, AddToRecentFiles=False)
prefixes = json.load(io.open(sys.argv[2], encoding='utf-8'))
out = io.open(sys.argv[3], 'w', encoding='utf-8')
try:
    found = {}
    for para in doc.Paragraphs:
        t = (para.Range.Text or '').strip()
        for p in prefixes:
            if p and t.startswith(p) and p not in found:
                found[p] = para.Format.Alignment
    for p in prefixes:
        out.write('%s|%s\n' % (found.get(p, 'NOTFOUND'), p))
finally:
    out.close()
    doc.Close(SaveChanges=False); word.Quit(); pythoncom.CoUninitialize()
print('done')
'''

env = dict(os.environ); env['PYTHONIOENCODING'] = 'utf-8'
for stem in STEMS:
    paths = sorted(glob.glob(os.path.join(DOCS, stem + '*.docx')))
    if not paths:
        out.write('%s: NO DOCX\n' % stem); continue
    path = paths[0]
    z = zipfile.ZipFile(path)
    x = z.read('word/document.xml').decode('utf-8')
    starts = [m.start() for m in re.finditer(r'<w:p [^>]*>|<w:p>', x)] + [len(x)]
    prefixes = []
    n_left = 0
    for a, b in zip(starts, starts[1:]):
        chunk = x[a:b]
        m = re.search(r'<w:pPr>.*?</w:pPr>', chunk, re.S)
        if not m:
            continue
        ppr = m.group(0)
        if '<w:jc w:val="left"' not in ppr:
            continue
        n_left += 1
        txt = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', chunk)).strip()
        if len(txt) >= 6:
            prefixes.append(txt[:10])
    prefixes = prefixes[:25]
    out.write('=== %s (%s): %d explicit jc=left paras, checking %d ===\n'
              % (stem, os.path.basename(path)[:40], n_left, len(prefixes)))
    if not prefixes:
        continue
    pj = 'c:/tmp/_s540_prefixes.json'
    import json as _json
    io.open(pj, 'w', encoding='utf-8').write(_json.dumps(prefixes, ensure_ascii=False))
    rf = 'c:/tmp/_s540_wordjc_%s.txt' % stem
    r = subprocess.run([sys.executable, '-c', WORD, path, pj, rf],
                       capture_output=True, text=True, encoding='utf-8', errors='replace',
                       timeout=300, env=env)
    if r.returncode != 0:
        out.write('COM FAIL: %s\n' % (r.stderr or '')[:300]); continue
    for line in io.open(rf, encoding='utf-8'):
        out.write('  ' + line)
out.close()
print('ok')
