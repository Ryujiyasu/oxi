# -*- coding: utf-8 -*-
"""S492i — paragraph break dataset for b837 (fitting target for the lookahead break).
For every paragraph: Word per-line char counts (COM) vs Oxi NATURAL per-line counts
(OXI_S474_NATURAL, no compression) vs Oxi flat-K(ship). Aligned by text prefix.
Reveals empirically where/when Word fits MORE than natural (compression), to derive
the lookahead rule. cp932-safe: UTF-8 file, JSON out (utf-8!), ASCII summary.
"""
import os, glob, subprocess, json, re
import win32com.client as w32

BIN = os.path.abspath('tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')
DOCX = os.path.abspath(glob.glob('tools/golden-test/documents/docx/b837*.docx')[0])
WD_VPOS = 6
ALIGN = {0: 'left', 1: 'center', 2: 'right', 3: 'both', 4: 'distribute'}


def oxi_lines(envset):
    env = dict(os.environ)
    for k in ('OXI_S474_NATURAL', 'OXI_S492_JCNATURAL'):
        env.pop(k, None)
    env.update(envset)
    subprocess.run([BIN, DOCX, 'c:/tmp/_bds', '--dump-layout=c:/tmp/_bds.json'],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=env)
    d = json.load(open('c:/tmp/_bds.json', encoding='utf-8'))
    pm = {}
    for pgi, pg in enumerate(d['pages']):
        for e in pg['elements']:
            if e['type'] == 'text' and e.get('para_idx') is not None:
                pm.setdefault(e['para_idx'], []).append((pgi, e))
    out = {}
    for pi, els in pm.items():
        els.sort(key=lambda t: (t[0], round(t[1]['y'], 1), t[1]['x']))
        txt = ''.join(e['text'] for _, e in els)
        lines = {}
        for pgi, e in els:
            lines.setdefault((pgi, round(e['y'], 1)), []).append(e)
        counts = [sum(len(e['text']) for e in lines[k]) for k in sorted(lines)]
        out[re.sub(r'\s', '', txt)[:14]] = counts
    return out


natural = oxi_lines({'OXI_S474_NATURAL': '1'})
flatk = oxi_lines({})

# Word per-line counts
word = w32.DispatchEx('Word.Application'); word.Visible = False
wpar = []
try:
    wdoc = word.Documents.Open(DOCX, ReadOnly=True)
    try:
        for p in wdoc.Paragraphs:
            clean = p.Range.Text.replace('\r', '').replace('\x07', '').replace('\n', '')
            if len(clean) < 8:
                continue
            rng = p.Range; txt = rng.Text; start = rng.Start
            y0 = wdoc.Range(start, start).Information(WD_VPOS)
            counts = []; cur = 0; prev_y = y0
            for i in range(len(txt)):
                ch = txt[i]
                if ch in ('\r', '\n', '\x07'):
                    continue
                y = wdoc.Range(start + i, start + i).Information(WD_VPOS)
                if y > prev_y + 2:
                    counts.append(cur); cur = 0; prev_y = y
                cur += 1
            counts.append(cur)
            wpar.append({'key': re.sub(r'\s', '', clean)[:14], 'jc': ALIGN.get(p.Alignment, '?'),
                         'li': round(p.LeftIndent, 1), 'total': len(clean), 'word': counts})
    finally:
        wdoc.Close(False)
finally:
    word.Quit()

ds = []
for w in wpar:
    nat = natural.get(w['key']); fk = flatk.get(w['key'])
    ds.append({**w, 'natural': nat, 'flatk': fk})
json.dump(ds, open('c:/tmp/b837_breakdataset.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

# ASCII analysis: per-para, compare line counts; flag where Word L0 != natural L0
print("=== b837 break dataset (jc / total / Word-lines / natural-lines / flatk-lines) ===")
n_word_eq_nat = n_word_gt_nat = n_word_lt_nat = 0
for w in wpar:
    nat = natural.get(w['key']); fk = flatk.get(w['key'])
    if not nat:
        continue
    we = sum(w['word']); ne = sum(nat)
    tagline = ''
    # compare per-line where possible
    if len(w['word']) != len(nat):
        tagline = 'NLINES W%d N%d' % (len(w['word']), len(nat))
    else:
        diffs = [w['word'][j] - nat[j] for j in range(len(nat))]
        if any(diffs):
            tagline = 'perline ' + str(diffs)
    if tagline:
        print("%-5s li=%3.0f tot=%3d  W=%s  N=%s  FK=%s  %s" %
              (w['jc'], w['li'], w['total'], w['word'], nat, fk, tagline))
print("\n(full: c:/tmp/b837_breakdataset.json)")
