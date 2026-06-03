# -*- coding: utf-8 -*-
"""S492g — per-context capacity derivation. Measure Word's per-char advances on the
b837 paras where flat-K=3.0 diverges: idx 29 (Oxi UNDER-packs, needs MORE capacity)
and idx 66/71/73 (Oxi OVER-packs, needs LESS). For each, list punct types + Word's
rendered advance (12 - advance = compression Word applied). If 、/openers compress
~1.5 while 。/closers/pair-first compress ~6, a per-TYPE capacity beats flat-K.
cp932-safe: UTF-8 file, results to JSON, ASCII summary. Advances use Info(5) diffs
(line-box convention cancels in the diff).
"""
import os, glob, json, re
import win32com.client as w32

WD_VPOS, WD_HPOS = 6, 5
DOCX = os.path.abspath(glob.glob('tools/golden-test/documents/docx/b837*.docx')[0])
rows = json.load(open('c:/tmp/b837_align.json', encoding='cp932'))
targets = {29: 'UNDER', 66: 'OVER', 71: 'OVER', 73: 'OVER'}
prefixes = {}
for i, r in enumerate(rows):
    if i in targets and r['oxi']:
        prefixes[i] = r['word']['norm'][:14]

CLOSE = set('、。，．）」』〕】》〉｝］')
OPEN = set('（「『〔【《〈｛［')
PUNCT = CLOSE | OPEN | set('・：；')

word = w32.DispatchEx('Word.Application'); word.Visible = False
out = {}
try:
    wdoc = word.Documents.Open(DOCX, ReadOnly=True)
    try:
        prefix_to_idx = {v: k for k, v in prefixes.items()}
        for p in wdoc.Paragraphs:
            clean = p.Range.Text.replace('\r', '').replace('\x07', '').replace('\n', '')
            key = re.sub(r'\s', '', clean)[:14]
            if key not in prefix_to_idx:
                continue
            idx = prefix_to_idx[key]
            rng = p.Range; txt = rng.Text; start = rng.Start
            y0 = wdoc.Range(start, start).Information(WD_VPOS)
            # collect L1 chars with x
            seq = []
            for i in range(len(txt)):
                ch = txt[i]
                if ch in ('\r', '\n', '\x07'):
                    continue
                y = wdoc.Range(start + i, start + i).Information(WD_VPOS)
                if y > y0 + 2:
                    break
                x = wdoc.Range(start + i, start + i).Information(WD_HPOS)
                seq.append((ch, round(x, 2)))
            # advances
            adv = []
            for j in range(len(seq) - 1):
                adv.append((seq[j][0], round(seq[j + 1][1] - seq[j][1], 2)))
            # punct compression: 12 - advance for punct chars (fs=12)
            punct_comp = []
            for j, (ch, a) in enumerate(adv):
                if ch in PUNCT:
                    nxt = adv[j + 1][0] if j + 1 < len(adv) else ''
                    cls = ('OPEN' if ch in OPEN else 'CLOSE' if ch in CLOSE else 'OTHER')
                    pairfirst = (ch in CLOSE and nxt in (OPEN | CLOSE))
                    punct_comp.append({'ch_class': cls, 'pairfirst': pairfirst,
                                       'adv': a, 'compress': round(12.0 - a, 2)})
            out[idx] = {'kind': targets[idx], 'L1': len(seq),
                        'n_punct': len(punct_comp), 'punct': punct_comp}
    finally:
        wdoc.Close(False)
finally:
    word.Quit()

json.dump(out, open('c:/tmp/b837_divergent.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print("=== Word per-context punct compression on divergent b837 paras ===")
for idx in sorted(out):
    o = out[idx]
    print("\nidx %d (%s) Word L1=%d, %d punct:" % (idx, o['kind'], o['L1'], o['n_punct']))
    for pc in o['punct']:
        tag = pc['ch_class'] + ('+pairfirst' if pc['pairfirst'] else '')
        print("   %-18s adv=%.2f  compress=%.2f" % (tag, pc['adv'], pc['compress']))
# aggregate by class
from collections import defaultdict
agg = defaultdict(list)
for idx in out:
    for pc in out[idx]['punct']:
        k = 'pairfirst' if pc['pairfirst'] else pc['ch_class']
        agg[k].append(pc['compress'])
print("\n=== compression by class (mean) ===")
for k, v in agg.items():
    print("  %-12s n=%d mean_compress=%.2f  (range %.1f-%.1f)" % (k, len(v), sum(v) / len(v), min(v), max(v)))
