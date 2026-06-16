# -*- coding: utf-8 -*-
"""tokyoshugyo #2 (char-budget OIDASHI wall) localizer.

Reliable per-line Word-vs-Oxi wrap comparison for the 賃金 chapter, built on a
CHAR-STREAM alignment (difflib) instead of text-prefix matching or fitz line
grouping — both of which the memory flagged as BLOCKED on this doc (repeated
regulation phrases collide; fitz over-counts table-dense pages).

Idea: Word PDF and Oxi dump emit the SAME character sequence in reading order.
Cluster each into visual lines (Word: PDF chars by Y baseline; Oxi: dump text
elements by distinct Y). Concatenate to a char stream tagging each char with its
visual line id. Align the two streams with difflib (handles footers/desync).
Then a char that STARTS a new Word line but is mid-line in Oxi = Word OIDASHI
(Oxi over-fits, keeps the char up). Tabulate the break chars + context.

Usage:
  python tools/metrics/_tks_oidashi.py            # uses cached PDF+dump
  python tools/metrics/_tks_oidashi.py --reexport # regen Word PDF
  python tools/metrics/_tks_oidashi.py --redump   # regen Oxi dump
"""
import os, sys, json, subprocess, tempfile, difflib
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DOCX = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx', 'tokyoshugyo_000599795.docx')
PDF  = os.path.join(tempfile.gettempdir(), 'tks_truth.pdf')
DUMP = os.environ.get('TKS_DUMP', os.path.join(tempfile.gettempdir(), 'tks_base.json'))
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')

# Word 賃金 chapter pages 46-64 (1-based); Oxi packs it into p47-64.
# Override with TKS_WP / TKS_OP = "lo:hi" for whole-doc analysis.
def _rng(env, default):
    v = os.environ.get(env)
    if v:
        lo, hi = v.split(':'); return range(int(lo), int(hi))
    return default
W_PAGES = _rng('TKS_WP', range(46, 65))
O_PAGES = _rng('TKS_OP', range(47, 65))

if '--reexport' in sys.argv or not os.path.exists(PDF):
    import win32com.client as win32
    w = win32.DispatchEx('Word.Application'); w.Visible = False
    try:
        d = w.Documents.Open(DOCX, ReadOnly=True)
        d.ExportAsFixedFormat(PDF, 17)
        d.Close(False)
    finally:
        w.Quit()
    print('exported PDF', file=sys.stderr)

if '--redump' in sys.argv or not os.path.exists(DUMP):
    subprocess.run([RENDERER, DOCX, os.path.join(tempfile.gettempdir(), 'tks_p_'),
                    '--dump-layout=' + DUMP], capture_output=True, timeout=300)
    print('dumped Oxi', file=sys.stderr)

import fitz

YAK = '、。，．・「」『』（）【】〔〕'
OPENERS = '「『（【〔《〈'
CLOSERS = '」』）】〕》〉'

# --demand: derive Word's per-約物 compression discriminator. For each Word 約物
# (mid-line, has a successor on the same line), classify compressed (advance <
# 0.9×em) vs natural, and tabulate by next-char class + line position. The em is
# the line's median non-約物 advance (robust to font size). Run with --reexport
# fresh PDF. This is the prescribed statistical derivation of the oikomi/oidashi
# per-line demand discriminator (the S573 wall).
if '--demand' in __import__('sys').argv:
    import sys as _sys, os as _os, tempfile as _tmp
    _sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    from collections import Counter as _C
    _PDF = _os.path.join(_tmp.gettempdir(), 'tks_truth.pdf')
    _doc = fitz.open(_PDF)
    def _cls(c):
        if c in '、，': return 'comma'
        if c in '。．': return 'period'
        if c == '・': return 'naka'
        if c in OPENERS: return 'opener'
        if c in CLOSERS: return 'closer'
        if c and ('0' <= c <= '9' or 'A' <= c <= 'Z' or 'a' <= c <= 'z'): return 'latin'
        return 'cjk'
    rows = []  # (yak_class, compressed, next_cls, posfrac, em)
    for pno in range(46, 65):  # 賃金 chapter
        rd = _doc[pno - 1].get_text('rawdict'); chars = []
        for blk in rd['blocks']:
            if blk.get('type', 0): continue
            for ln in blk.get('lines', []):
                for sp in ln['spans']:
                    for ch in sp['chars']:
                        b = ch['bbox']
                        if 60 < b[1] < 800: chars.append((round((b[1]+b[3])/2, 1), b[0], b[2], ch['c']))
        rws = {}
        for yc, x0, x1, c in chars: rws.setdefault(yc, []).append((x0, x1, c))
        for yc in sorted(rws):
            r = sorted(rws[yc])
            if len(r) < 4: continue
            advs = [r[i+1][0] - r[i][0] for i in range(len(r)-1)]
            # em = median advance of non-約物 chars
            nonyak = [advs[i] for i in range(len(r)-1) if r[i][2] not in YAK]
            if not nonyak: continue
            em = sorted(nonyak)[len(nonyak)//2]
            x_left = r[0][0]; x_right = r[-1][1]; span = x_right - x_left
            for i in range(len(r)-1):
                c = r[i][2]
                if c not in '、。，．・」』）】〕':  # compressible 約物
                    continue
                adv = advs[i]
                compressed = adv < 0.9 * em
                nxt = _cls(r[i+1][2])
                posfrac = (r[i][0] - x_left) / span if span > 0 else 0
                rows.append((_cls(c), compressed, nxt, posfrac, em, adv))
    print(f"\n===== Word per-約物 DEMAND discriminator ({len(rows)} compressible 約物, 賃金 chapter) =====")
    comp = [r for r in rows if r[1]]
    nat = [r for r in rows if not r[1]]
    print(f"  compressed (oikomi): {len(comp)} ({100*len(comp)/max(1,len(rows)):.1f}%)   natural: {len(nat)}")
    print("  --- by 約物 type: compressed / total ---")
    for yt in ['comma', 'period', 'naka', 'closer']:
        tot = [r for r in rows if r[0] == yt]; cc = [r for r in tot if r[1]]
        if tot: print(f"     {yt}: {len(cc)}/{len(tot)} compressed ({100*len(cc)/len(tot):.0f}%)")
    print("  --- compressed 約物 by NEXT-char class ---")
    for k, n in _C(r[2] for r in comp).most_common():
        tot = sum(1 for r in rows if r[2] == k)
        print(f"     next={k}: {n} compressed / {tot} total ({100*n/tot:.0f}% of next={k})")
    print("  --- natural 約物 by NEXT-char class (for contrast) ---")
    for k, n in _C(r[2] for r in nat).most_common():
        print(f"     next={k}: {n} natural")
    # position: do compressed 約物 cluster late in the line (line-fitting)?
    import statistics as _st
    if comp: print(f"  compressed posfrac: median={_st.median(r[3] for r in comp):.2f}")
    if nat: print(f"  natural   posfrac: median={_st.median(r[3] for r in nat):.2f}")
    _sys.exit()

def word_lines():
    """Word PDF chapter -> list of dicts {page,y,chars=[(c,x0,x1)]}, reading order."""
    doc = fitz.open(PDF)
    out = []
    for pno in W_PAGES:
        pg = doc[pno - 1]
        rd = pg.get_text('rawdict')
        chars = []
        for blk in rd['blocks']:
            if blk.get('type', 0) != 0: continue
            for ln in blk.get('lines', []):
                for sp in ln['spans']:
                    for ch in sp['chars']:
                        b = ch['bbox']
                        y0 = b[1]
                        if y0 < 60 or y0 > 800: continue  # strip header/footer
                        chars.append((ch['c'], b[0], b[2], (b[1] + b[3]) / 2))
        # cluster by Y center
        chars.sort(key=lambda t: (round(t[3] / 3.0), t[1]))
        rows = {}
        for c, x0, x1, yc in chars:
            key = round(yc / 3.0)
            rows.setdefault(key, []).append((c, x0, x1, yc))
        for key in sorted(rows):
            row = sorted(rows[key], key=lambda t: t[1])
            out.append({'page': pno, 'y': row[0][3],
                        'chars': [(c, x0, x1) for c, x0, x1, _ in row]})
    return out

def oxi_lines():
    """Oxi dump chapter -> list of dicts {page,y,text,char_offset,para}, reading order."""
    d = json.load(open(DUMP, encoding='utf-8'))
    out = []
    for pg in d['pages']:
        if pg['page'] not in O_PAGES: continue
        rows = {}
        for el in pg.get('elements', []):
            if el.get('type') != 'text' or not el.get('text'): continue
            key = round(el['y'], 0)
            rows.setdefault(key, []).append(el)
        for key in sorted(rows):
            row = sorted(rows[key], key=lambda e: e['x'])
            txt = ''.join(e.get('text', '') for e in row)
            out.append({'page': pg['page'], 'y': key, 'text': txt,
                        'char_offset': row[0].get('char_offset'),
                        'para': row[0].get('cell_para_idx'),
                        'para_idx': row[0].get('para_idx'),
                        'row': row[0].get('cell_row_idx'),
                        'col': row[0].get('cell_col_idx'),
                        'x0': row[0]['x'], 'x1': row[-1]['x'] + row[-1].get('w', 0)})
    return out

W = word_lines()
O = oxi_lines()
print(f"Word chapter visual lines: {len(W)}   Oxi chapter visual lines: {len(O)}", file=sys.stderr)

# Per Word page: content-right boundary = max line right-extent (proxy for the
# right margin / cell content edge that Word fills to).
wpage_right = {}
for ln in W:
    x1 = max((x1 for _, _, x1 in ln['chars']), default=0)
    wpage_right[ln['page']] = max(wpage_right.get(ln['page'], 0), x1)
print("Word per-page content-right (max line extent):", file=sys.stderr)
for p in sorted(wpage_right): print(f"   p{p}: {wpage_right[p]:.1f}", file=sys.stderr)

# Build char streams tagging each char with its line index. Strip spaces so the
# two streams align (Word PDF may not emit trailing spaces; grids differ).
def stream(lines, getchars):
    s = []          # list of (char, line_idx)
    for li, ln in enumerate(lines):
        for c in getchars(ln):
            if c in (' ', '　', '\t', '\r', '\n'): continue
            s.append((c, li))
    return s

ws = stream(W, lambda ln: [c for c, _, _ in ln['chars']])
os_ = stream(O, lambda ln: list(ln['text']))
wtext = ''.join(c for c, _ in ws)
otext = ''.join(c for c, _ in os_)
print(f"Word stream chars: {len(wtext)}   Oxi stream chars: {len(otext)}", file=sys.stderr)

sm = difflib.SequenceMatcher(None, wtext, otext, autojunk=False)
ratio = sm.ratio()
print(f"stream similarity ratio: {ratio:.4f}", file=sys.stderr)

# Walk matching blocks. For each matched char pair (wi, oi), we know its
# word_line and oxi_line. Detect OIDASHI: a char that is the FIRST char of its
# Word line (word_line increments vs the previous matched char) while in Oxi it
# stays on the SAME line as the previous matched char (oxi_line unchanged).
events = []
prev = None  # (w_line, o_line, wi, oi)
for blk in sm.get_matching_blocks():
    a, b, size = blk.a, blk.b, blk.size
    for k in range(size):
        wi, oi = a + k, b + k
        wl = ws[wi][1]; ol = os_[oi][1]
        c = wtext[wi]
        if prev is not None:
            pwl, pol, pwi, poi = prev
            word_broke = wl != pwl
            oxi_broke = ol != pol
            if word_broke and not oxi_broke:
                # Word started a new line here; Oxi kept this char on prev line.
                wprev = W[pwl]['chars']
                wthis = W[wl]['chars']
                w_end_x1 = max((x1 for _, _, x1 in wprev), default=0)
                pushed_w = (wthis[0][2] - wthis[0][1]) if wthis else 0  # first char width
                pright = wpage_right.get(W[pwl]['page'], 0)
                room = pright - w_end_x1
                oxi = O[ol]
                in_cell = oxi.get('para') is not None
                events.append({
                    'kind': 'OIDASHI',          # Word pushed down, Oxi over-fit
                    'wpage': W[wl]['page'], 'opage': O[ol]['page'],
                    'break_char': c, 'prev_char': wtext[pwi],
                    'w_prev_line': ''.join(x for x, _, _ in wprev),
                    'w_this_line': ''.join(x for x, _, _ in wthis),
                    'o_line': oxi['text'],
                    'w_end_x1': w_end_x1, 'pright': pright, 'room': room,
                    'pushed_w': pushed_w, 'in_cell': in_cell,
                    'ox1': oxi.get('x1', 0),
                })
            elif oxi_broke and not word_broke:
                events.append({
                    'kind': 'OIKOMI',           # Oxi broke earlier than Word
                    'wpage': W[wl]['page'], 'opage': O[ol]['page'],
                    'break_char': c, 'prev_char': wtext[pwi],
                    'w_this_line': ''.join(x for x, _, _ in W[wl]['chars']),
                    'o_line': O[ol]['text'],
                })
        prev = (wl, ol, wi, oi)

oid = [e for e in events if e['kind'] == 'OIDASHI']
oik = [e for e in events if e['kind'] == 'OIKOMI']
print(f"\n=== break divergences: {len(oid)} OIDASHI (Oxi over-fit), {len(oik)} OIKOMI (Oxi under-fit) ===\n")

def ctx(s, n=22): return s[-n:] if s else ''
# Decisive: did Word HAVE ROOM for the pushed char (demand oidashi) or not (width)?
had_room = [e for e in oid if e['room'] >= e['pushed_w'] - 0.5]
no_room  = [e for e in oid if e['room'] <  e['pushed_w'] - 0.5]
print(f"VERDICT split: {len(had_room)} HAD-ROOM (demand oidashi), {len(no_room)} NO-ROOM (width gap)")
hr_cell = sum(1 for e in had_room if e['in_cell']); hr_body = len(had_room) - hr_cell
print(f"  HAD-ROOM split: {hr_cell} cell (boundary artifact: pageRight≠cell-right), {hr_body} BODY (genuine non-greedy/kinsoku)")
ncell = sum(1 for e in oid if e['in_cell']); print(f"  cell lines: {ncell}/{len(oid)}, body lines: {len(oid)-ncell}/{len(oid)}\n")
print("--- OIDASHI detail (room = pageRight − WordLineEnd; pushed_w = pushed char width) ---")
for e in oid:
    v = 'HAD-ROOM(demand)' if e['room'] >= e['pushed_w'] - 0.5 else 'no-room(width)'
    loc = 'CELL' if e['in_cell'] else 'body'
    print(f"  Wp{e['wpage']} {loc} prev='{e['prev_char']}'→'{e['break_char']}' "
          f"Wend={e['w_end_x1']:.1f} pR={e['pright']:.1f} room={e['room']:.1f} pushW={e['pushed_w']:.1f} oxiX1={e['ox1']:.1f}  [{v}]")
    print(f"     W end: …{ctx(e['w_prev_line'])}  | next:{e['w_this_line'][:14]}")

# Feature tally: what char does Word push down, what precedes it
from collections import Counter
print("\n--- OIDASHI break-char (first char of Word's new line) frequency ---")
for c, n in Counter(e['break_char'] for e in oid).most_common():
    tag = ' [yakumono]' if c in YAK else (' [opener]' if c in OPENERS else '')
    print(f"   '{c}' ×{n}{tag}")
print("\n--- OIDASHI prev-char (last char Word keeps on the line) frequency ---")
for c, n in Counter(e['prev_char'] for e in oid).most_common():
    tag = ' [yakumono]' if c in YAK else ''
    print(f"   '{c}' ×{n}{tag}")

# ===== ROOTS: per-paragraph FIRST break divergence (the cascade roots). Within
# each Oxi paragraph, walk matched chars in order; Oxi breaks where its line
# increments, Word where its line increments. The FIRST char where they disagree
# = that paragraph's root divergence (the rest of the para cascades). =====
if '--roots' in sys.argv:
    from collections import Counter
    def pkey(ol): return (ol.get('para_idx'), ol.get('para'), ol.get('row'), ol.get('col'))
    # matched (word_line, oxi_line, char) in stream order
    matched = []
    for blk in sm.get_matching_blocks():
        for k in range(blk.size):
            wi = blk.a + k; oi = blk.b + k
            matched.append((ws[wi][1], os_[oi][1], wtext[wi]))
    # group by oxi paragraph (contiguous runs of same pkey)
    roots = []
    i = 0
    n = len(matched)
    while i < n:
        para = pkey(O[matched[i][1]])
        j = i
        seq = []
        while j < n and pkey(O[matched[j][1]]) == para:
            seq.append(matched[j]); j += 1
        # walk seq; find first break disagreement
        for k in range(1, len(seq)):
            pwl, pol, _ = seq[k-1]; wl, ol, c = seq[k]
            wb = wl != pwl; ob = ol != pol
            if wb != ob:
                kind = 'W-breaks-Oxi-keeps(over-fit)' if wb else 'Oxi-breaks-W-keeps(under-fit)'
                roots.append({'para': para, 'kind': kind, 'char': c,
                              'prev': seq[k-1][2],
                              'in_cell': para[1] is not None,
                              'ctx': ''.join(x[2] for x in seq[max(0,k-10):k+2])})
                break
        i = j
    over = [r for r in roots if 'over-fit' in r['kind']]
    under = [r for r in roots if 'under-fit' in r['kind']]
    print(f"\n===== per-paragraph CASCADE ROOTS: {len(roots)} paragraphs with a break divergence =====")
    nc = sum(1 for r in roots if r['in_cell'])
    print(f"  {len(over)} over-fit roots, {len(under)} under-fit roots | {nc} cell, {len(roots)-nc} body")
    print("  --- root break-char freq (the char at the first divergence) ---")
    for c, k in Counter(r['char'] for r in roots).most_common(15):
        t = ' [YAK]' if c in YAK else ''
        print(f"     '{c}' ×{k}{t}")
    # Classify body OVER-fit roots: autospace (digit/Latin in ctx) vs 約物 vs pure-CJK
    import re
    body_over = [r for r in roots if not r['in_cell'] and 'over-fit' in r['kind']]
    has_latin = [r for r in body_over if re.search(r'[0-9A-Za-zＡ-Ｚａ-ｚ０-９]', r['ctx'])]
    has_yak = [r for r in body_over if any(c in YAK for c in r['ctx']) and r not in has_latin]
    pure = [r for r in body_over if r not in has_latin and r not in has_yak]
    print(f"  --- body OVER-fit roots ({len(body_over)}): {len(has_latin)} have digit/Latin (autospace), "
          f"{len(has_yak)} 約物-only, {len(pure)} pure-CJK ---")
    print("  sample autospace (digit/Latin) body over-fit roots:")
    for r in has_latin[:10]:
        print(f"     '{r['prev']}'→'{r['char']}'  …{r['ctx']}")
    print("  sample pure-CJK body over-fit roots:")
    for r in pure[:10]:
        print(f"     '{r['prev']}'→'{r['char']}'  …{r['ctx']}")

# ===== PAGE-DELTA: collision-robust per-char page delta (the gate the official
# pagination_diff can't compute on this doc — char-stream aligned, no text-prefix
# collision). delta = oxi_page − word_page. Run whole-doc (TKS_WP/OP=1:91). =====
if '--pagedelta' in sys.argv:
    from collections import Counter
    dh = Counter()
    matched = 0
    for blk in sm.get_matching_blocks():
        for k in range(blk.size):
            wi = blk.a + k; oi = blk.b + k
            wp = W[ws[wi][1]]['page']; op = O[os_[oi][1]]['page']
            dh[op - wp] += 1
            matched += 1
    print(f"\n===== per-char PAGE-DELTA (oxi_page − word_page), {matched} matched chars =====")
    tot = sum(dh.values())
    for d in sorted(dh):
        bar = '#' * int(60 * dh[d] / tot)
        print(f"   Δ{d:+d}: {dh[d]:6d} ({100*dh[d]/tot:5.1f}%) {bar}")
    d0 = dh.get(0, 0)
    print(f"   delta=0 (correct page): {100*d0/tot:.1f}%")

# ===== ABSORPTION: how much 約物 compression Word applies per FULL line =====
# absorption = Σ(natural em widths) − actual line width. >0 = Word compressed
# 約物 (oikomi); ≈0 = natural; <0 = Word expanded (justify spread). The MAX
# positive absorption on full (justified) lines = Word's per-line oikomi budget.
if '--absorb' in sys.argv:
    import unicodedata
    FS = 10.5
    def is_fw(c):
        if c in ' 　\t': return False
        return unicodedata.east_asian_width(c) in ('W', 'F', 'A')
    COMPRESS = '、。，．・」』】〕》〉｝］）'
    rows = []
    for li, ln in enumerate(W):
        cs = ln['chars']
        if len(cs) < 3: continue
        x0 = cs[0][1]; x1 = cs[-1][2]
        actual = x1 - x0
        nat = sum(FS if is_fw(c) else FS / 2.0 for c, _, _ in cs)
        pright = wpage_right.get(ln['page'], 0)
        full = x1 >= pright - FS  # reaches the margin = justified non-last line
        nyak = sum(1 for c, _, _ in cs if c in COMPRESS)
        rows.append({'absorb': nat - actual, 'full': full, 'nyak': nyak,
                     'page': ln['page'], 'n': len(cs)})
    full = [r for r in rows if r['full']]
    print(f"\n===== ABSORPTION on {len(full)} full(justified) Word lines (of {len(rows)}) =====")
    import statistics
    av = [r['absorb'] for r in full]
    print(f"  absorption pt: min={min(av):.2f} p25={statistics.quantiles(av,n=4)[0]:.2f} "
          f"median={statistics.median(av):.2f} p75={statistics.quantiles(av,n=4)[2]:.2f} "
          f"p95={sorted(av)[int(len(av)*0.95)]:.2f} max={max(av):.2f}")
    # histogram in 1pt buckets
    from collections import Counter
    h = Counter(round(a) for a in av)
    print("  histogram (1pt buckets):", dict(sorted(h.items())))
    # per-約物 compression on lines that DID compress (absorb > 0.5)
    comp = [r for r in full if r['absorb'] > 0.5 and r['nyak'] > 0]
    print(f"\n  lines Word COMPRESSED (absorb>0.5, has 約物): {len(comp)}")
    per = [r['absorb'] / r['nyak'] for r in comp]
    if per:
        print(f"  per-約物 compression pt: median={statistics.median(per):.2f} "
              f"p95={sorted(per)[int(len(per)*0.95)]:.2f} max={max(per):.2f}")
    # lines that did NOT compress despite having 約物 (absorb<=0.5) = oidashi/expand
    nocomp = [r for r in full if r['absorb'] <= 0.5 and r['nyak'] > 0]
    print(f"  full lines NOT compressed despite 約物 (expand/natural): {len(nocomp)}")
    # the discriminator question: is there a clean absorb cap?
    big = [r for r in full if r['absorb'] > 3.0]
    print(f"  full lines with absorption > 3.0pt: {len(big)} "
          f"(absorb,nyak): {[(round(r['absorb'],1),r['nyak']) for r in sorted(big,key=lambda r:-r['absorb'])[:15]]}")
    # Print the TEXT of the high-absorb lines + flag 約物 PAIRS (default kerning,
    # not demand oikomi). Need the line object back: rows carry page+n but not li.
    print("  --- high-absorb line texts (does compression = 約物 PAIR kern or DEMAND?) ---")
    def has_pair(s):
        return any(s[i] in COMPRESS and s[i+1] in (COMPRESS + OPENERS) for i in range(len(s)-1))
    shown = 0
    for ln in W:
        cs = ln['chars']
        if len(cs) < 3: continue
        nat = sum(FS if is_fw(c) else FS/2.0 for c,_,_ in cs)
        ab = nat - (cs[-1][2]-cs[0][1])
        if ab > 3.0 and cs[-1][2] >= wpage_right.get(ln['page'],0) - FS:
            txt = ''.join(c for c,_,_ in cs)
            print(f"     absorb={ab:.1f} pair={has_pair(txt)} | {txt[:44]}")
            shown += 1
            if shown >= 14: break

# ===== OIKOMI side: what char does Word KEEP that Oxi wrapped? (kinsoku check) =====
print("\n--- OIKOMI break-char (char Word KEEPS on the line, Oxi wraps) frequency ---")
def is_prohib(c): return c in '、。，．）」』】〕》〉｝］・ーぁぃぅぇぉっゃゅょゎ々'
for c, n in Counter(e['break_char'] for e in oik).most_common(20):
    tag = ' [LINE-START-PROHIBITED]' if is_prohib(c) else (' [yak]' if c in YAK else '')
    print(f"   '{c}' ×{n}{tag}")
nprohib = sum(1 for e in oik if is_prohib(e['break_char']))
print(f"   => {nprohib}/{len(oik)} OIKOMI break-chars are line-start-prohibited")
noid_prohib = sum(1 for e in oid if is_prohib(e['break_char']))
print(f"   (for comparison: {noid_prohib}/{len(oid)} OIDASHI break-chars are prohibited)")

# ===== per-YAKUMONO advance comparison Word vs Oxi (the break-budget evidence) =====
if '--yak' in sys.argv:
    # Flat per-char lists with x + line id + advance-to-next-on-same-line.
    def flat_word():
        f = []
        for li, ln in enumerate(W):
            cs = ln['chars']  # (c,x0,x1)
            for k, (c, x0, x1) in enumerate(cs):
                adv = (cs[k+1][1] - x0) if k+1 < len(cs) else None
                f.append((c, x0, li, adv))
        return f
    def flat_oxi():
        f = []
        d2 = json.load(open(DUMP, encoding='utf-8'))
        for li, ln in enumerate(O):
            pass
        # rebuild per-char from dump elements grouped per O line via (page,y)
        # Simpler: re-walk O lines, splitting element text into chars at x + w/len
        for li, ln in enumerate(O):
            # we only kept line text + x0/x1; need per-char x. Re-extract from dump.
            pass
        return f
    # Per-char Oxi x: re-extract from the dump elements (text may be multi-char).
    d2 = json.load(open(DUMP, encoding='utf-8'))
    ochars = []  # (c, x, line_id)
    lineidx = {}
    for li, ln in enumerate(O):
        lineidx[(ln['page'], ln['y'])] = li
    for pg in d2['pages']:
        if pg['page'] not in O_PAGES: continue
        rows = {}
        for el in pg.get('elements', []):
            if el.get('type') != 'text' or not el.get('text'): continue
            rows.setdefault(round(el['y'], 0), []).append(el)
        for y, els in rows.items():
            li = lineidx.get((pg['page'], y))
            if li is None: continue
            for el in sorted(els, key=lambda e: e['x']):
                t = el.get('text', '')
                x = el['x']; w = el.get('w', 0)
                # element may hold multiple chars; distribute width evenly (CJK)
                n = len(t)
                cw = (w / n) if n else 0
                for j, c in enumerate(t):
                    ochars.append((c, x + j * cw, li, cw))
    # build advance for oxi per line
    ofch = []
    for k in range(len(ochars)):
        c, x, li, cw = ochars[k]
        adv = None
        if k+1 < len(ochars) and ochars[k+1][2] == li:
            adv = ochars[k+1][1] - x
        ofch.append((c, x, li, adv))
    wf = flat_word()
    wstr = ''.join(c for c, *_ in wf if c not in ' 　\t\r\n')
    ostr = ''.join(c for c, *_ in ofch if c not in ' 　\t\r\n')
    wmap = [i for i, (c, *_ ) in enumerate(wf) if c not in ' 　\t\r\n']
    omap = [i for i, (c, *_ ) in enumerate(ofch) if c not in ' 　\t\r\n']
    sm2 = difflib.SequenceMatcher(None, wstr, ostr, autojunk=False)
    rows = []  # (yak, next_char, word_adv, oxi_adv, page)
    for blk in sm2.get_matching_blocks():
        for k in range(blk.size):
            wi = wmap[blk.a + k]; oi = omap[blk.b + k]
            c = wf[wi][0]
            if c not in '、。，．・「」『』（）':
                continue
            wadv = wf[wi][3]; oadv = ofch[oi][3]
            if wadv is None or oadv is None:
                continue
            # next char (for context) from word stream
            nxt = wstr[blk.a + k + 1] if blk.a + k + 1 < len(wstr) else ''
            rows.append((c, nxt, wadv, oadv, W[wf[wi][2]]['page']))
    print(f"\n===== per-yakumono advance Word vs Oxi ({len(rows)} mid-line matched 約物) =====")
    # classify
    def cls(adv, fs=10.5): return 'NAT' if adv >= fs*0.95 else ('½' if adv >= fs*0.6 else 'HVY')
    from collections import Counter as C2
    joint = C2()
    for c, nxt, wa, oa, pg in rows:
        joint[(cls(wa), cls(oa))] += 1
    print("joint (Word-class, Oxi-class) counts:")
    for k, n in sorted(joint.items(), key=lambda x: -x[1]):
        print(f"   Word {k[0]:3} / Oxi {k[1]:3}: {n}")
    # The over-fit signature: Word NAT but Oxi compressed
    sig = [r for r in rows if cls(r[2]) == 'NAT' and cls(r[3]) != 'NAT']
    # Per-page-bucket: where is the over-fit (Word-NAT/Oxi-comp) vs Word-compress?
    print("\n--- per-page: Word-NAT-Oxi-comp (over-fit)  vs  Word-comp (agree/under) ---")
    pgbuck = {}
    for c, nxt, wa, oa, pg in rows:
        b = pgbuck.setdefault(pg, [0, 0, 0])
        wc = cls(wa) == 'NAT'
        oc = cls(oa) != 'NAT'
        if wc and oc: b[0] += 1          # over-fit signature
        elif not wc: b[1] += 1           # Word compresses
        b[2] += 1
    for pg in sorted(pgbuck):
        b = pgbuck[pg]
        print(f"   p{pg}: overfit(WNAT/Ocomp)={b[0]:2d}  Word-compresses={b[1]:2d}  total約物={b[2]}")
    print(f"\n*** Word-NAT but Oxi-compressed (the break over-fit) : {len(sig)} ***")
    nb = C2(r[1] for r in sig)
    print("   next-char after the over-compressed 約物:")
    for ch, n in nb.most_common(12):
        t = ' [opener]' if ch in OPENERS else (' [yak]' if ch in YAK else ' [kanji/kana]')
        print(f"      '{ch}' ×{n}{t}")
    for r in sig[:12]:
        print(f"      '{r[0]}'→'{r[1]}' Word adv={r[2]:.2f} Oxi adv={r[3]:.2f} (p{r[4]})")
    # Also: does Word EVER compress mid-、 (+kanji) in this chapter?
    wc = [r for r in rows if cls(r[2]) != 'NAT' and r[1] not in OPENERS and r[1] not in YAK]
    print(f"\n*** Word compresses 約物 followed by KANJI/KANA (true demand oikomi): {len(wc)} ***")
    for r in wc[:20]:
        print(f"      '{r[0]}'→'{r[1]}' Word adv={r[2]:.2f} Oxi adv={r[3]:.2f} (p{r[4]})")
