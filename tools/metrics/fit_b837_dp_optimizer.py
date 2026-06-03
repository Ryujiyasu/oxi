# -*- coding: utf-8 -*-
"""S492o — DP (Knuth-Plass-style) paragraph break optimizer + weight search, scored vs
Word per-line counts on b837, with ACCURATE per-line avails derived from the dump's
actual line x0 (not approximate firstLine). HARD GATE: DP must beat the natural/greedy
score on the SAME clean set to justify the Rust build. cp932-safe (UTF-8, ASCII out).
"""
import os, glob, subprocess, json, re

BIN = os.path.abspath('tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')
DOCX = os.path.abspath(glob.glob('tools/golden-test/documents/docx/b837*.docx')[0])
BOUNDARY = 524.4  # left margin 70.9 + content 453.5
CLOSE = set('、。，．）」』〕】》〉｝］')
OPEN = set('（「『〔【《〈｛［')
LINE_START_PROH = CLOSE | set('・：；？！ぁぃぅぇぉっゃゅょゎ')
LINE_END_PROH = OPEN


def max_comp(ch, nxt):
    if ch in CLOSE and (nxt in OPEN or nxt in CLOSE):
        return 6.0
    if ch in ('。', '．') or ch in CLOSE:
        return 6.0
    if ch in ('、', '，') or ch in OPEN:
        return 1.5
    return 0.0


env = dict(os.environ); env['OXI_S474_NATURAL'] = '1'; env.pop('OXI_S492_JCNATURAL', None)
subprocess.run([BIN, DOCX, 'c:/tmp/_dp', '--dump-layout=c:/tmp/_dp.json'],
               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=env)
d = json.load(open('c:/tmp/_dp.json', encoding='utf-8'))
pm = {}
for pgi, pg in enumerate(d['pages']):
    for e in pg['elements']:
        if e['type'] == 'text' and e.get('para_idx') is not None:
            pm.setdefault(e['para_idx'], []).append((pgi, e))
oxi_chars = {}; oxi_avail = {}
for pi, els in pm.items():
    els.sort(key=lambda t: (t[0], round(t[1]['y'], 1), t[1]['x']))
    chars = []
    for _, e in els:
        t = e['text']
        if len(t) == 1:
            chars.append((t, e['w']))
        else:
            for c in t:
                chars.append((c, e['w'] / len(t)))
    key = re.sub(r'\s', '', ''.join(c for c, _ in chars))[:14]
    oxi_chars[key] = chars
    lines = {}
    for pgi, e in els:
        lines.setdefault((pgi, round(e['y'], 1)), []).append(e)
    lk = sorted(lines)
    x0_L0 = min(e['x'] for e in lines[lk[0]])
    x0_cont = min((min(e['x'] for e in lines[k]) for k in lk[1:]), default=x0_L0)
    oxi_avail[key] = (BOUNDARY - x0_L0, BOUNDARY - x0_cont)

ds = [r for r in json.load(open('c:/tmp/b837_breakdataset.json', encoding='utf-8'))
      if r.get('word') and r['key'] in oxi_chars and r['jc'] in ('both', 'distribute')
      and len(r['word']) > 1 and sum(r['word']) == len(oxi_chars[r['key']])]
print("clean justified multi-line paras:", len(ds))


def prefix(chars):
    N = len(chars)
    nat = [c[1] for c in chars]
    mc = [max_comp(chars[k][0], chars[k + 1][0] if k + 1 < N else '') for k in range(N)]
    pn = [0.0] * (N + 1); pc = [0.0] * (N + 1)
    for k in range(N):
        pn[k + 1] = pn[k] + nat[k]; pc[k + 1] = pc[k] + mc[k]
    return pn, pc


def dp(chars, av, w_slack, w_comp, w_line, tol=0.6):
    N = len(chars); pn, pc = prefix(chars)
    aL0, aC = av
    INF = float('inf'); best = [INF] * (N + 1); best[0] = 0.0; prev = [-1] * (N + 1)
    for j in range(1, N + 1):
        if j < N and chars[j][0] in LINE_START_PROH:
            continue
        if chars[j - 1][0] in LINE_END_PROH:
            continue
        for i in range(j):
            if best[i] == INF:
                continue
            avail = aL0 if i == 0 else aC
            natural = pn[j] - pn[i]; comp = pc[j] - pc[i]
            if natural - comp > avail + tol:
                continue
            if j == N:
                lc = 0.0
            else:
                slack = max(0.0, avail - natural); used = max(0.0, natural - avail)
                lc = w_slack * slack * slack + w_comp * used * used
            t = best[i] + lc + w_line
            if t < best[j]:
                best[j] = t; prev[j] = i
    if best[N] == INF:
        return None
    counts = []; j = N
    while j > 0:
        i = prev[j]; counts.append(j - i); j = i
    return counts[::-1]


def greedy_natural(chars, av):
    N = len(chars); aL0, aC = av; counts = []; i = 0; first = True
    while i < N:
        avail = aL0 if first else aC; w = 0.0; c = 0
        while i + c < N:
            w += chars[i + c][1]
            if w > avail + 0.6:
                break
            c += 1
        c = max(1, c)
        if i + c < N and chars[i + c][0] in LINE_START_PROH and c > 1:
            c -= 1
        counts.append(c); i += c; first = False
    return counts


gok = sum(1 for r in ds if greedy_natural(oxi_chars[r['key']], oxi_avail[r['key']]) == r['word'])
print("natural-greedy (accurate avail): %d/%d (%.0f%%)" % (gok, len(ds), 100 * gok / len(ds)))


def greedy_maxcomp(chars, av):
    """Greedy max-fill using REAL per-context max_comp (not flat-K): a char fits iff
    (cum_natural - cum_maxcomp) <= avail."""
    N = len(chars); aL0, aC = av; counts = []; i = 0; first = True
    while i < N:
        avail = aL0 if first else aC; w = 0.0; mc = 0.0; c = 0
        while i + c < N:
            ch = chars[i + c][0]; nxt = chars[i + c + 1][0] if i + c + 1 < N else ''
            w += chars[i + c][1]; mc += max_comp(ch, nxt)
            if w - mc > avail + 0.6:
                break
            c += 1
        c = max(1, c)
        if i + c < N and chars[i + c][0] in LINE_START_PROH and c > 1:
            c -= 1
        counts.append(c); i += c; first = False
    return counts


mok = sum(1 for r in ds if greedy_maxcomp(oxi_chars[r['key']], oxi_avail[r['key']]) == r['word'])
print("greedy + per-context max_comp: %d/%d (%.0f%%)" % (mok, len(ds), 100 * mok / len(ds)))

best = (0, None)
for ws in (0.0, 0.001, 0.01, 0.1, 1.0):
    for wc in (0.0, 0.01, 0.1, 1.0, 10.0):
        for wl in (0.0, 10.0, 50.0, 200.0, 1000.0, 5000.0):
            ok = sum(1 for r in ds if dp(oxi_chars[r['key']], oxi_avail[r['key']], ws, wc, wl) == r['word'])
            if ok > best[0]:
                best = (ok, (ws, wc, wl))
print("DP BEST: %d/%d (%.0f%%) at w_slack=%s w_comp=%s w_line=%s" %
      (best[0], len(ds), 100 * best[0] / len(ds), *best[1]))
ws, wc, wl = best[1]
print("\nper-para (mark | Word | DP | natGreedy) at best weights:")
for r in ds:
    dpc = dp(oxi_chars[r['key']], oxi_avail[r['key']], ws, wc, wl)
    g = greedy_natural(oxi_chars[r['key']], oxi_avail[r['key']])
    print("  %s av=(%.0f,%.0f) W=%s DP=%s G=%s" %
          ('OK' if dpc == r['word'] else 'X ', oxi_avail[r['key']][0], oxi_avail[r['key']][1], r['word'], dpc, g))
