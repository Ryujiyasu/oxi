# -*- coding: utf-8 -*-
"""S544: Word per-char advances for the ed025c ②③⑤ paras (the S543 over-pack
targets). Writes c:/tmp/s544_word.txt. ascii-safe console."""
import glob, io, os, subprocess, sys

DOCS = os.path.abspath('tools/golden-test/documents/docx')

WORD = r'''
# -*- coding: utf-8 -*-
import sys, io
import win32com.client, pythoncom
pythoncom.CoInitialize()
word = win32com.client.Dispatch('Word.Application'); word.Visible=False; word.DisplayAlerts=False
word.AutomationSecurity = 3
doc = word.Documents.Open(sys.argv[1], ReadOnly=True, AddToRecentFiles=False)
out = io.open(r'c:/tmp/s544_word.txt', 'w', encoding='utf-8')
try:
    start = doc.GoTo(What=1, Which=1, Count=2).Start
    rng = doc.Range(start, doc.Content.End)
    n_done = 0
    for para in rng.Paragraphs:
        t = (para.Range.Text or '').strip()
        if not (t[:1] in (u'②', u'③', u'⑤')):
            continue
        pr = para.Range
        out.write('PARA|%s\n' % t[:24].replace('\r',''))
        n = min(pr.End - pr.Start, 100)
        for i in range(n):
            c = doc.Range(pr.Start + i, pr.Start + i)
            ch = doc.Range(pr.Start + i, pr.Start + i + 1).Text or ''
            if ch in ('\r', '\x07', '\n'): continue
            out.write('C|%d|%.2f|%.2f|%s\n' % (i, c.Information(6), c.Information(5), ch))
        n_done += 1
        if n_done >= 4: break
finally:
    out.close()
    doc.Close(SaveChanges=False); word.Quit(); pythoncom.CoUninitialize()
print('done')
'''

docx = glob.glob(os.path.join(DOCS, 'ed025c*.docx'))[0]
env = dict(os.environ); env['PYTHONIOENCODING'] = 'utf-8'
r = subprocess.run([sys.executable, '-c', WORD, docx], capture_output=True, text=True,
                   encoding='utf-8', errors='replace', timeout=600, env=env)
print('rc=%s %s %s' % (r.returncode, (r.stdout or '').strip()[:100], (r.stderr or '').strip()[:200]))

# analyze: per line, char count + advances summary
s = io.open('c:/tmp/s544_word.txt', encoding='utf-8').read()
out = io.open('c:/tmp/s544_word_lines.txt', 'w', encoding='utf-8')
cur = None
paras = []
for ln in s.splitlines():
    if ln.startswith('PARA|'):
        cur = {'t': ln[5:], 'chars': []}
        paras.append(cur)
    elif ln.startswith('C|') and cur is not None:
        _, i, y, x, ch = ln.split('|', 4)
        cur['chars'].append((int(i), float(y), float(x), ch))
for p in paras:
    out.write('=== %s ===\n' % p['t'])
    lines = {}
    for i, y, x, ch in p['chars']:
        lines.setdefault(y, []).append((i, x, ch))
    for y in sorted(lines):
        chs = sorted(lines[y])
        advs = [round(chs[k+1][1]-chs[k][1], 2) for k in range(len(chs)-1)]
        nat = sum(1 for a in advs if abs(a-10.5) < 0.01)
        comp = [(chs[k][2], advs[k]) for k in range(len(advs)) if advs[k] < 10.4]
        out.write('L n=%d x0=%.2f xlast=%.2f nat10.5=%d non-natural=%s\n'
                  % (len(chs), chs[0][1], chs[-1][1], nat,
                     [(c, a) for c, a in comp][:12]))
        out.write('  txt=%s\n' % ''.join(c for _, _, c in chs)[:50])
out.close()
print('ok')
