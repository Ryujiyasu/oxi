"""S492 — investigate the 3a4f F1 regression. Is F1 correct on 3a4f's jc=left lines
(SSIM drop = pagination noise on a large doc) or does it over/under-pack them (real bug)?

Full 3a4f doc: Word page count + per-paragraph (resolved alignment, L1 char count) for
wrapping paras, vs Oxi GDI L1 OFF and ON. Matched by text prefix.
"""
import os, glob, subprocess, json, re
import win32com.client as w32

BIN = os.path.abspath('tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')
WD_VPOS = 6
ALIGN = {0: 'left', 1: 'center', 2: 'right', 3: 'both', 4: 'distribute'}
DOCX = os.path.abspath(glob.glob('tools/golden-test/documents/docx/3a4f*.docx')[0])


def oxi_l1(on):
    env = dict(os.environ); env.pop('OXI_S492_JCNATURAL', None)
    if on: env['OXI_S492_JCNATURAL'] = '1'
    out = 'c:/tmp/_3a4f_%d.json' % on
    subprocess.run([BIN, DOCX, 'c:/tmp/_3a4f_x', '--dump-layout=' + out],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=env)
    d = json.load(open(out, encoding='utf-8'))
    paras = {}
    for pg in d['pages']:
        for e in pg['elements']:
            if e['type'] != 'text' or e.get('para_idx') is None:
                continue
            paras.setdefault(e['para_idx'], []).append(e)
    by_pref = {}
    for pi, els in paras.items():
        y0 = sorted(set(round(e['y'], 1) for e in els))[0]
        l1 = [e for e in els if abs(e['y'] - y0) < 2]
        full = sorted(els, key=lambda e: (round(e['y'], 1), e['x']))
        pref = re.sub(r'\s', '', ''.join(e['text'] for e in full))[:12]
        by_pref[pref] = len(l1)
    return by_pref


off = oxi_l1(0)
on = oxi_l1(1)

word = w32.DispatchEx('Word.Application'); word.Visible = False
try:
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    try:
        pages = doc.ComputeStatistics(2)  # wdStatisticPages
        print("Word page count for 3a4f:", pages)
        rows = []
        for p in doc.Paragraphs:
            rng = p.Range; txt = rng.Text
            clean = txt.replace('\r', '').replace('\x07', '').replace('\n', '')
            if len(clean) < 20:
                continue
            start, end = rng.Start, rng.End
            y0 = doc.Range(start, start).Information(WD_VPOS)
            yN = doc.Range(max(start, end - 1), max(start, end - 1)).Information(WD_VPOS)
            if (yN - y0) <= 2:
                continue
            al = ALIGN.get(p.Alignment, str(p.Alignment))
            n = 0
            for i in range(len(txt)):
                if txt[i] in ('\r', '\n', '\x07'):
                    continue
                if doc.Range(start + i, start + i).Information(WD_VPOS) > y0 + 2:
                    break
                n += 1
            pref = re.sub(r'\s', '', clean)[:12]
            rows.append((al, n, off.get(pref), on.get(pref), clean[:14]))
    finally:
        doc.Close(False)
finally:
    word.Quit()

from collections import Counter
ac = Counter(r[0] for r in rows)
print("wrapping paras by alignment:", dict(ac))
print("\n=== jc=left wrapping lines: Word vs Oxi OFF vs ON (delta to Word) ===")
print("%-6s %5s %5s %5s  %8s %8s  %s" % ('align', 'Word', 'OFF', 'ON', 'OFFΔ', 'ONΔ', 'text'))
off_better = on_better = same = 0
for al, w, o, nn, t in rows:
    if al == 'both' or o is None or nn is None:
        continue
    od = o - w; nd = nn - w
    if abs(nd) < abs(od): on_better += 1
    elif abs(nd) > abs(od): off_better += 1
    else: same += 1
    tag = ''
    if abs(nd) < abs(od): tag = '  ON closer'
    elif abs(nd) > abs(od): tag = '  ON WORSE'
    print("%-6s %5d %5d %5d  %+8d %+8d  %s%s" % (al, w, o, nn, od, nd, t, tag))
print("\nON closer to Word: %d ; ON worse: %d ; same: %d" % (on_better, off_better, same))
