# -*- coding: utf-8 -*-
"""S538: per-paragraph Y comparison for one Word page (Word COM Information(6) on
collapsed starts vs Oxi dump-layout text element tops, matched by text prefix).
Finds WHERE on the page the cumulative height gap accumulates.
Usage: python _s538_para_y_cmp.py <word_page>"""
import os, sys, io, json, subprocess
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DOCX = r'c:\tmp\_3a4f_copy.docx'
GDI = os.path.join(ROOT, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')

WORD = r'''
import sys
import win32com.client, pythoncom
pythoncom.CoInitialize()
word = win32com.client.Dispatch('Word.Application'); word.Visible=False; word.DisplayAlerts=False
word.AutomationSecurity = 3
doc = word.Documents.Open(sys.argv[1], ReadOnly=True, AddToRecentFiles=False)
target = int(sys.argv[2])
try:
    # jump straight to the target page (wdGoToPage=1, wdGoToAbsolute=1)
    start = doc.GoTo(What=1, Which=1, Count=target).Start
    rng = doc.Range(start, doc.Content.End)
    count = 0
    for para in rng.Paragraphs:
        pr = para.Range
        s = doc.Range(pr.Start, pr.Start)
        pg = s.Information(3)
        if pg < target:
            continue  # the first paragraph may straddle the page boundary
        if pg > target:
            break
        y = s.Information(6)
        t = (pr.Text or '').strip()[:16].replace('\r','').replace('\x07','')
        print('WP|%.2f|%s' % (y, t))
        count += 1
        if count > 80:
            break
finally:
    doc.Close(SaveChanges=False); word.Quit(); pythoncom.CoUninitialize()
'''


def main():
    wpage = int(sys.argv[1])
    env = dict(os.environ)
    env['PYTHONIOENCODING'] = 'utf-8'
    r = subprocess.run([sys.executable, '-c', WORD, DOCX, str(wpage)],
                       capture_output=True, text=True, encoding='utf-8', errors='replace',
                       timeout=180, env=env)
    wparas = []
    for l in (r.stdout or '').splitlines():
        if l.startswith('WP|'):
            _, y, t = l.split('|', 2)
            if t.strip():
                wparas.append((float(y), t.strip()))
    # Oxi: use the latest dump (regenerate to be safe)
    subprocess.run([GDI, DOCX, 'c:/tmp/s538', '150', '--dump-layout=c:/tmp/s538.json'], capture_output=True)
    d = json.load(io.open('c:/tmp/s538.json', encoding='utf-8'))
    # search oxi pages wpage-2..wpage for each word para's first match (leftmost element of a line)
    oxi_lines = []  # (page, y, text) per line: group text elements by (page, round y)
    from collections import defaultdict
    for pi in range(max(0, wpage - 3), min(len(d['pages']), wpage + 1)):
        groups = defaultdict(list)
        for e in d['pages'][pi].get('elements', []):
            if str(e.get('type','')) == 'text' and (e.get('text') or '').strip():
                groups[round(e.get('y',0) / 3)].append(e)
        for k, es in groups.items():
            es.sort(key=lambda e: e.get('x', 0))
            txt = ''.join(e.get('text','') for e in es)[:16]
            oxi_lines.append((pi + 1, es[0].get('y', 0), txt))
    L = []
    prev_w = prev_o = None
    for wy, wt in wparas:
        best = None
        for (opg, oy, ot) in oxi_lines:
            if ot[:6] and wt[:6] and (ot[:6] in wt or wt[:6] in ot):
                best = (opg, oy)
                break
        if best:
            opg, oy = best
            dw = ('%6.1f' % (wy - prev_w)) if prev_w is not None else '     -'
            do = ('%6.1f' % (oy - prev_o)) if prev_o is not None and opg == prev_opg else '     -'
            L.append('w_y=%7.2f (d%s) | o p%-3d y=%7.2f (d%s) | %s' % (wy, dw, opg, oy, do, wt))
            prev_w, prev_o, prev_opg = wy, oy, opg
        else:
            L.append('w_y=%7.2f          | (no oxi match)            | %s' % (wy, wt))
    txt = '\n'.join(L)
    io.open('c:/tmp/_s538_out.txt', 'w', encoding='utf-8').write(txt + '\n')
    print('written %d word paras' % len(wparas))


if __name__ == '__main__':
    main()
