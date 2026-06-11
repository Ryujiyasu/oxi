# -*- coding: utf-8 -*-
"""S540: per-char X positions of the 3a4f maru-1 paragraph (① 労働者の意向…)
via Word COM collapsed ranges, vs Oxi dump elements. Goal: find why Word
breaks L2 after 場合に (37 chars) while Oxi fits 場合には (38).
Writes c:/tmp/s540_word_chars.txt. Run: python _s540_maru1_charx.py
"""
import io, json, os, subprocess, sys

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DOCX = r'c:\tmp\_3a4f_copy.docx'
GDI = os.path.join(ROOT, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')

WORD = r'''
# -*- coding: utf-8 -*-
import sys, io
import win32com.client, pythoncom
pythoncom.CoInitialize()
word = win32com.client.Dispatch('Word.Application'); word.Visible=False; word.DisplayAlerts=False
word.AutomationSecurity = 3
doc = word.Documents.Open(sys.argv[1], ReadOnly=True, AddToRecentFiles=False)
out = io.open(r'c:/tmp/s540_word_chars.txt', 'w', encoding='utf-8')
try:
    # jump to page 61, scan paragraphs for the target prefix
    start = doc.GoTo(What=1, Which=1, Count=59).Start
    rng = doc.Range(start, doc.Content.End)
    target = None
    for para in rng.Paragraphs:
        t = (para.Range.Text or '')
        if t.startswith(u'①'):  # ①
            if u'労働者の意向を踏まえた' in t:
                target = para
                break
    if target is None:
        out.write('TARGET NOT FOUND\n')
    else:
        pr = target.Range
        out.write('paralen=%d\n' % (pr.End - pr.Start))
        n = min(pr.End - pr.Start, 130)
        for i in range(n):
            c = doc.Range(pr.Start + i, pr.Start + i)
            ch = doc.Range(pr.Start + i, pr.Start + i + 1).Text or ''
            x = c.Information(5)   # wdHorizontalPositionRelativeToPage
            y = c.Information(6)
            pg = c.Information(3)
            out.write('%d|pg%s|y=%.2f|x=%.2f|%s\n' % (i, pg, y, x, ch.replace('\r','<CR>').replace('\n','')))
finally:
    out.close()
    doc.Close(SaveChanges=False); word.Quit(); pythoncom.CoUninitialize()
print('word done')
'''

env = dict(os.environ)
env['PYTHONIOENCODING'] = 'utf-8'
r = subprocess.run([sys.executable, '-c', WORD, DOCX], capture_output=True, text=True,
                   encoding='utf-8', errors='replace', timeout=300, env=env)
print('word rc=%s out=%s err=%s' % (r.returncode, (r.stdout or '').strip()[:200], (r.stderr or '').strip()[:300]))

# Oxi side: dump para 556 elements per line with x and w
subprocess.run([GDI, DOCX, r'c:\tmp\s540', '--dump-layout=c:/tmp/s540_dump.json'],
               capture_output=True, env=env)
d = json.load(io.open('c:/tmp/s540_dump.json', encoding='utf-8'))
out = io.open('c:/tmp/s540_oxi_chars.txt', 'w', encoding='utf-8')
for pi, page in enumerate(d['pages']):
    if pi < 58 or pi > 62:
        continue
    els = [e for e in page.get('elements', []) if e.get('type') == 'text' and e.get('para_idx') == 556]
    lines = {}
    for e in els:
        lines.setdefault(round(e['y'], 1), []).append(e)
    for y in sorted(lines):
        row = sorted(lines[y], key=lambda e: e.get('x', 0))
        out.write('p%d y=%.2f\n' % (pi + 1, y))
        for e in row:
            out.write('  x=%8.2f w=%7.2f |%s|\n' % (e.get('x', 0), e.get('w', 0), e.get('text', '')))
out.close()
print('oxi done')
