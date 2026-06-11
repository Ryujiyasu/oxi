# -*- coding: utf-8 -*-
"""S541: generalized per-char comparison for one page of one doc.
Word COM per-char (collapsed-range Information(5/6), advances) vs Oxi dump
elements, paragraph by paragraph. Reports per-line char counts and the first
advance divergence per line. Usage:
    python _s541_page_charcmp.py <docx-prefix> <word_page> [max_paras]
Outputs c:/tmp/s541_<prefix>_p<page>.txt
"""
import glob, io, json, os, subprocess, sys

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DOCS = os.path.join(ROOT, 'tools', 'golden-test', 'documents', 'docx')
GDI = os.path.join(ROOT, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')

WORD = r'''
# -*- coding: utf-8 -*-
import sys, io
import win32com.client, pythoncom
pythoncom.CoInitialize()
word = win32com.client.Dispatch('Word.Application'); word.Visible=False; word.DisplayAlerts=False
word.AutomationSecurity = 3
doc = word.Documents.Open(sys.argv[1], ReadOnly=True, AddToRecentFiles=False)
target = int(sys.argv[2]); maxp = int(sys.argv[3])
out = io.open(sys.argv[4], 'w', encoding='utf-8')
try:
    start = doc.GoTo(What=1, Which=1, Count=target).Start
    rng = doc.Range(start, doc.Content.End)
    np = 0
    for para in rng.Paragraphs:
        pr = para.Range
        s = doc.Range(pr.Start, pr.Start)
        pg = s.Information(3)
        if pg < target: continue
        if pg > target or np >= maxp: break
        np += 1
        out.write('PARA|%d|%s\n' % (np, (pr.Text or '').strip()[:20].replace('\r','')))
        n = min(pr.End - pr.Start, 200)
        for i in range(n):
            c = doc.Range(pr.Start + i, pr.Start + i)
            ch = doc.Range(pr.Start + i, pr.Start + i + 1).Text or ''
            if ch in ('\r', '\x07', '\n'): continue
            out.write('C|%d|%.2f|%.2f|%s\n' % (i, c.Information(6), c.Information(5), ch))
finally:
    out.close()
    doc.Close(SaveChanges=False); word.Quit(); pythoncom.CoUninitialize()
print('done')
'''


def main():
    prefix, wpage = sys.argv[1], int(sys.argv[2])
    maxp = int(sys.argv[3]) if len(sys.argv) > 3 else 12
    docx = sorted(glob.glob(os.path.join(DOCS, prefix + '*.docx')))[0]
    env = dict(os.environ); env['PYTHONIOENCODING'] = 'utf-8'
    wf = 'c:/tmp/_s541_word_raw.txt'
    r = subprocess.run([sys.executable, '-c', WORD, docx, str(wpage), str(maxp), wf],
                       capture_output=True, text=True, encoding='utf-8', errors='replace',
                       timeout=600, env=env)
    if r.returncode != 0:
        print('COM FAIL', (r.stderr or '')[:400]); return 1

    # word: group chars into lines per para by y
    wparas = []  # [(text20, [ [(idx,x,ch),...] per line ])]
    cur = None
    for line in io.open(wf, encoding='utf-8'):
        line = line.rstrip('\n')
        if line.startswith('PARA|'):
            cur = {'text': line.split('|', 2)[2], 'lines': {}}
            wparas.append(cur)
        elif line.startswith('C|') and cur is not None:
            _, i, y, x, ch = line.split('|', 4)
            cur['lines'].setdefault(float(y), []).append((int(i), float(x), ch))

    # oxi dump
    subprocess.run([GDI, docx, r'c:\tmp\s541', '--dump-layout=c:/tmp/s541_dump.json'],
                   capture_output=True, env=env)
    d = json.load(io.open('c:/tmp/s541_dump.json', encoding='utf-8'))

    out = io.open('c:/tmp/s541_%s_p%d.txt' % (prefix, wpage), 'w', encoding='utf-8')
    for wp in wparas:
        head = wp['text'][:8]
        if not head:
            continue
        out.write('\n=== W:%s ===\n' % wp['text'])
        wlines = []
        for y in sorted(wp['lines']):
            chs = sorted(wp['lines'][y])
            txt = ''.join(c for _, _, c in chs)
            advs = [round(chs[k + 1][1] - chs[k][1], 2) for k in range(len(chs) - 1)]
            wlines.append((txt, advs))
            x0 = chs[0][1]
            xlast = chs[-1][1]
            out.write('W n=%2d x0=%7.2f xlast=%7.2f |%s| advs=%s\n'
                      % (len(chs), x0, xlast, txt[:42], advs[:20]))
        # find oxi para: match by concatenated prefix over pages wpage-2..wpage+1
        best = None
        for pi in range(max(0, wpage - 3), min(len(d['pages']), wpage + 1)):
            groups = {}
            for e in d['pages'][pi].get('elements', []):
                if e.get('type') != 'text': continue
                k = (e.get('para_idx'), e.get('cell_para_idx'), e.get('cell_row_idx'), e.get('cell_col_idx'))
                groups.setdefault(k, []).append(e)
            for k, els in groups.items():
                lines = {}
                for e in els:
                    lines.setdefault(round(e['y'], 1), []).append(e)
                first_y = sorted(lines)[0]
                row = sorted(lines[first_y], key=lambda e: e.get('x', 0))
                txt = ''.join(e.get('text', '') for e in row)
                if txt[:6] and head and (txt[:6].startswith(head[:6]) or head.startswith(txt[:6])):
                    best = (pi + 1, lines)
                    break
            if best: break
        if not best:
            out.write('O (no match)\n'); continue
        opg, lines = best
        for y in sorted(lines):
            row = sorted(lines[y], key=lambda e: e.get('x', 0))
            txt = ''.join(e.get('text', '') for e in row)
            # per-element widths as advances proxy
            ws = [round(e.get('w', 0), 2) for e in row]
            n = sum(len(e.get('text', '')) for e in row)
            xend = max(e.get('x', 0) + e.get('w', 0) for e in row)
            out.write('O p%d n=%2d x0=%7.2f xend=%7.2f |%s| el_ws=%s\n'
                      % (opg, n, row[0].get('x', 0), xend, txt[:42], ws[:20]))
    out.close()
    print('ok -> c:/tmp/s541_%s_p%d.txt' % (prefix, wpage))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
