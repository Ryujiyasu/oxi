# -*- coding: utf-8 -*-
"""S492m — empirically find the b837 break model. Simulate candidate per-line break
models on Oxi NATURAL per-char widths and score each against Word's per-line counts
(b837_breakdataset.json). Models: natural-greedy, flatK, demand (greedy + compress a
line only to avoid an extra line), balance (min raggedness). Whichever best matches
Word's per-line counts is the rule to implement in Rust. cp932-safe (UTF-8, ASCII out).
"""
import os, glob, subprocess, json, re

BIN = os.path.abspath('tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')
DOCX = os.path.abspath(glob.glob('tools/golden-test/documents/docx/b837*.docx')[0])

# CJK punct + max break-compression (pt) per type (measured render table, S492g/S470)
CLOSE = set('、。，．）」』〕】》〉｝］')
OPEN = set('（「『〔【《〈｛［')
def max_comp(ch, nxt):
    if ch in CLOSE and (nxt in OPEN or nxt in CLOSE):
        return 6.0  # pair-first half-em collapse
    if ch in ('。', '．') or ch in CLOSE:
        return 6.0  # period / closing bracket: heavy
    if ch in ('、', '，') or ch in OPEN:
        return 1.5  # comma / opening: light
    return 0.0

# Oxi natural per-char widths per para
env = dict(os.environ); env['OXI_S474_NATURAL'] = '1'; env.pop('OXI_S492_JCNATURAL', None)
subprocess.run([BIN, DOCX, 'c:/tmp/_fm', '--dump-layout=c:/tmp/_fm.json'],
               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=env)
d = json.load(open('c:/tmp/_fm.json', encoding='utf-8'))
pm = {}
for pgi, pg in enumerate(d['pages']):
    for e in pg['elements']:
        if e['type'] == 'text' and e.get('para_idx') is not None:
            pm.setdefault(e['para_idx'], []).append((pgi, e))
oxi_chars = {}  # key -> [(char, width)]
for pi, els in pm.items():
    els.sort(key=lambda t: (t[0], round(t[1]['y'], 1), t[1]['x']))
    chars = []
    for _, e in els:
        t = e['text']
        if len(t) == 1:
            chars.append((t, e['w']))
        else:  # merged element: split width evenly
            for c in t:
                chars.append((c, e['w'] / len(t)))
    key = re.sub(r'\s', '', ''.join(c for c, _ in chars))[:14]
    oxi_chars[key] = chars

ds = json.load(open('c:/tmp/b837_breakdataset.json', encoding='utf-8'))

LINE_START_PROH = CLOSE | set('・：；？！ぁぃぅぇぉっゃゅょゎ')  # approx Word default (no ー)


def greedy(chars, avails, compress_fn):
    """Generic greedy: per line, fit chars; compress_fn(line_chars, avail) -> max chars."""
    lines = []
    i = 0
    n = len(chars)
    li = 0
    while i < n:
        avail = avails(li)
        cnt = compress_fn(chars, i, avail)
        cnt = max(1, cnt)
        # kinsoku oidashi: if next char (line start) would be prohibited, pull back 1
        if i + cnt < n and chars[i + cnt][0] in LINE_START_PROH and cnt > 1:
            cnt -= 1
        lines.append(cnt)
        i += cnt
        li += 1
    return lines


def fit_natural(chars, i, avail):
    w = 0.0; c = 0
    while i + c < len(chars):
        w += chars[i + c][1]
        if w > avail + 0.5:
            break
        c += 1
    return c


def fit_compressed(chars, i, avail, cap_mult):
    # fit chars allowing each punct to compress up to its max*cap_mult
    w = 0.0; comp = 0.0; c = 0
    while i + c < len(chars):
        ch, cw = chars[i + c]
        nxt = chars[i + c + 1][0] if i + c + 1 < len(chars) else ''
        w += cw
        comp += max_comp(ch, nxt) * cap_mult
        if w - comp > avail + 0.5:
            break
        c += 1
    return c


def score(model_lines_fn):
    ok = 0; tot = 0
    for rec in ds:
        key = rec['key']
        if key not in oxi_chars or not rec.get('word'):
            continue
        chars = oxi_chars[key]
        li_pt = rec['li']
        fli = 12.0 if rec['li'] >= 12 else 0.0  # approx firstLine (b837 paras have ~12)
        content = 453.5
        def avails(lineidx, li_pt=li_pt):
            base = content - li_pt
            return base - (12.0 if lineidx == 0 else 0.0)  # approx firstLine on L0
        lines = model_lines_fn(chars, avails)
        tot += 1
        if lines == rec['word']:
            ok += 1
    return ok, tot


models = {
    'natural': lambda chars, av: greedy(chars, av, fit_natural),
    'flatK(cap1.0 all=3)': lambda chars, av: greedy(chars, av, lambda c, i, a: _fitK(c, i, a)),
    'demand(maxcomp)': lambda chars, av: greedy(chars, av, lambda c, i, a: fit_compressed(c, i, a, 1.0)),
    'demand(0.5)': lambda chars, av: greedy(chars, av, lambda c, i, a: fit_compressed(c, i, a, 0.5)),
}


def _fitK(chars, i, avail):
    w = 0.0; np = 0; c = 0
    while i + c < len(chars):
        ch, cw = chars[i + c]
        w += cw
        if ch in CLOSE or ch in OPEN:
            np += 1
        if w - np * 3.0 > avail + 0.5:
            break
        c += 1
    return c


print("model                    exact-para-match / total")
for name, fn in models.items():
    ok, tot = score(fn)
    print("  %-22s %d / %d  (%.0f%%)" % (name, ok, tot, 100 * ok / max(1, tot)))
