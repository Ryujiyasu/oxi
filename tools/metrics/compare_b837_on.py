# -*- coding: utf-8 -*-
"""S492f — compare Oxi OFF vs S492-ON (incl linesAndChars) against Word for b837.
Reuses Word data from c:/tmp/b837_align.json (the OFF alignment); renders Oxi with
OXI_S492_JCNATURAL=1 and matches paras by normalized text prefix. Reports whether the
jc=left over-pack/overflow is fixed and whether any para's page/line count regressed.
"""
import os, glob, subprocess, json, re

BIN = os.path.abspath('tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')
DOCX = os.path.abspath(glob.glob('tools/golden-test/documents/docx/b837*.docx')[0])
BOUNDARY = 524.4

env = dict(os.environ); env['OXI_S492_JCNATURAL'] = '1'
subprocess.run([BIN, DOCX, 'c:/tmp/_b837on', '--dump-layout=c:/tmp/_b837on.json'],
               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=env)
don = json.load(open('c:/tmp/_b837on.json', encoding='utf-8'))
opmap = {}
for pgi, pg in enumerate(don['pages']):
    for e in pg['elements']:
        if e['type'] == 'text' and e.get('para_idx') is not None:
            opmap.setdefault(e['para_idx'], []).append((pgi, e))
on_by = {}
for pi, els in opmap.items():
    els.sort(key=lambda t: (t[0], round(t[1]['y'], 1), t[1]['x']))
    txt = ''.join(e['text'] for _, e in els)
    lines = {}
    for pgi, e in els:
        lines.setdefault((pgi, round(e['y'], 1)), []).append(e)
    lk = sorted(lines)
    l1 = sum(len(e['text']) for e in lines[lk[0]])
    maxxend = max((e['x'] + e['w']) for _, e in els)
    on_by[re.sub(r'\s', '', txt)[:12]] = {'nlines': len(lines), 'l1': l1,
                                          'page1': lk[0][0] + 1, 'maxxend': round(maxxend, 1)}

rows = json.load(open('c:/tmp/b837_align.json', encoding='utf-8'))
total_pages_on = len(don['pages'])
print("Oxi b837 total pages: ON=%d (Word=7)" % total_pages_on)
print("\nidx jc    Wln Wl1 Wpg | OFF(ln/l1/xend/pg) | ON(ln/l1/xend/pg)  verdict")
fixed = regr = 0
for i, rec in enumerate(rows):
    w = rec['word']; o = rec['oxi']
    if not o:
        continue
    pref = w['norm'][:12]
    on = on_by.get(pref)
    if not on:
        continue
    # interesting = jc=left, or L1 diverged OFF, or overflow OFF
    interesting = (w['jc'] != 'both') or (o['l1'] != w['l1']) or (o['maxxend'] > BOUNDARY + 1)
    if not interesting:
        continue
    v = []
    if o['l1'] != w['l1'] and on['l1'] == w['l1']:
        v.append('L1-FIXED'); fixed += 1
    elif on['l1'] != w['l1'] and on['l1'] != o['l1']:
        v.append('L1 %+d->%+d' % (o['l1'] - w['l1'], on['l1'] - w['l1']))
    if o['maxxend'] > BOUNDARY + 1 and on['maxxend'] <= BOUNDARY + 1:
        v.append('OVF-FIXED')
    if on['nlines'] != o['nlines']:
        v.append('NLINES %d->%d' % (o['nlines'], on['nlines']))
    if on['page1'] != o['page1']:
        v.append('PAGE %d->%d' % (o['page1'], on['page1']))
    if on['l1'] != w['l1'] and o['l1'] == w['l1']:
        v.append('L1-BROKE'); regr += 1
    print("%3d %-5s %3d %3d %3d | %2d/%2d/%6.1f/%d | %2d/%2d/%6.1f/%d  %s" %
          (i, w['jc'], w['nlines'], w['l1'], w['page1'],
           o['nlines'], o['l1'], o['maxxend'], o['page1'],
           on['nlines'], on['l1'], on['maxxend'], on['page1'], ' '.join(v)))
print("\nL1-FIXED=%d  L1-BROKE=%d  (ON total pages=%d vs Word 7)" % (fixed, regr, total_pages_on))
