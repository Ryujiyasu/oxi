# -*- coding: utf-8 -*-
"""ikujidetail wi355/wi440 char-stream-aligned line analysis.

The official pagination_diff uses text-PREFIX matching, which the memory flags
as UNRELIABLE on ikujidetail: the regulation repeats legal phrases verbatim
(産前・産後休業、出生時育児休業… appears dozens of times) so per-paragraph
prefix matches collide and drift measurements mis-attribute paragraphs.

This tool instead CHAR-STREAM aligns the Word PDF render-truth against the Oxi
dump (difflib over the full character sequence, footers stripped), then labels
each matched char with:
  - Word page + Word visual-line id (from PDF glyph Y clusters)
  - Oxi  page + Oxi  visual-line id + para_idx + char_offset (authoritative)

From that single alignment it derives:
  --pagedelta : collision-robust per-char page delta (oxi_page - word_page),
                + per-Word-page dominant delta = where the +1 cascade onsets.
  --region LO:HI : side-by-side Word-line vs Oxi-line table for Word pages
                LO..HI, with the page each line sits on, so the exact line
                where Oxi's page fills one line earlier than Word is visible.
  --roots     : per-Oxi-paragraph FIRST break divergence (OIDASHI=Oxi over-fits
                keeps a char Word pushes down; OIKOMI=Oxi wraps a char Word keeps).
  --boundary  : zoom on the two known boundaries (wi355 Wp13, wi440 Wp16): the
                last ~6 Word lines of the page before the spill + the spilling
                paragraph, char-aligned to Oxi, to pin the gained line.

Usage:
  python tools/metrics/_ikuji_charalign.py --pagedelta
  python tools/metrics/_ikuji_charalign.py --region 11:14
  python tools/metrics/_ikuji_charalign.py --roots
  python tools/metrics/_ikuji_charalign.py --boundary
  add --reexport to regen the Word PDF, --redump to regen the Oxi dump.
"""
import os, sys, json, subprocess, tempfile, difflib
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DOCX = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx', 'ikujidetail_002197815.docx')
PDF  = os.path.join(tempfile.gettempdir(), 'ikd_truth.pdf')
DUMP = os.environ.get('IKD_DUMP', os.path.join(tempfile.gettempdir(), 'ikuji_dump.json'))
RENDERER = os.path.join(REPO, 'tools', 'oxi-gdi-renderer', 'target', 'release', 'oxi-gdi-renderer.exe')

YAK = '、。，．・「」『』（）【】〔〕'
OPENERS = '「『（【〔《〈'
CLOSERS = '」』）】〕》〉'
PROHIB_START = '、。，．）」』】〕》〉｝］・ーぁぃぅぇぉっゃゅょゎ々'

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
    subprocess.run([RENDERER, DOCX, os.path.join(tempfile.gettempdir(), 'ikd_p_'),
                    '--dump-layout=' + DUMP], capture_output=True, timeout=300)
    print('dumped Oxi', file=sys.stderr)

import fitz


def word_lines():
    """Word PDF -> reading-order list of {page, y, chars=[(c,x0,x1)]}."""
    doc = fitz.open(PDF)
    out = []
    for pno in range(1, len(doc) + 1):
        pg = doc[pno - 1]
        rd = pg.get_text('rawdict')
        chars = []
        for blk in rd['blocks']:
            if blk.get('type', 0) != 0:
                continue
            for ln in blk.get('lines', []):
                for sp in ln['spans']:
                    for ch in sp['chars']:
                        b = ch['bbox']
                        y0 = b[1]
                        if y0 < 65 or y0 > 800:   # strip header/footer
                            continue
                        chars.append((ch['c'], b[0], b[2], (b[1] + b[3]) / 2))
        # cluster by Y center (line spacing ~14.3pt; group when gap > 7pt)
        chars.sort(key=lambda t: (t[3], t[1]))
        rows = []
        cur = []
        cur_y = None
        for c, x0, x1, yc in chars:
            if cur_y is None or abs(yc - cur_y) <= 7.0:
                cur.append((c, x0, x1, yc))
                cur_y = yc if cur_y is None else (cur_y * (len(cur) - 1) + yc) / len(cur)
            else:
                rows.append(cur); cur = [(c, x0, x1, yc)]; cur_y = yc
        if cur:
            rows.append(cur)
        for row in rows:
            row.sort(key=lambda t: t[1])
            out.append({'page': pno, 'y': round(row[0][3], 1),
                        'chars': [(c, x0, x1) for c, x0, x1, _ in row]})
    return out


def oxi_lines():
    """Oxi dump -> reading-order list of {page, y, text, para_idx, char_offset, x0, x1}."""
    d = json.load(open(DUMP, encoding='utf-8'))
    out = []
    for pg in d['pages']:
        rows = {}
        for el in pg.get('elements', []):
            if el.get('type') != 'text' or not el.get('text'):
                continue
            rows.setdefault(round(el['y'], 1), []).append(el)
        for key in sorted(rows):
            row = sorted(rows[key], key=lambda e: e['x'])
            txt = ''.join(e.get('text', '') for e in row)
            out.append({'page': pg['page'], 'y': key, 'text': txt,
                        'para_idx': row[0].get('para_idx'),
                        'char_offset': row[0].get('char_offset'),
                        'x0': row[0]['x'], 'x1': row[-1]['x'] + row[-1].get('w', 0)})
    return out


W = word_lines()
O = oxi_lines()
print(f"Word visual lines: {len(W)}   Oxi visual lines: {len(O)}", file=sys.stderr)


def stream(lines, getchars):
    s = []
    for li, ln in enumerate(lines):
        for c in getchars(ln):
            if c in (' ', '　', '\t', '\r', '\n'):
                continue
            s.append((c, li))
    return s


ws = stream(W, lambda ln: [c for c, _, _ in ln['chars']])
os_ = stream(O, lambda ln: list(ln['text']))
wtext = ''.join(c for c, _ in ws)
otext = ''.join(c for c, _ in os_)
print(f"Word stream chars: {len(wtext)}   Oxi stream chars: {len(otext)}", file=sys.stderr)

sm = difflib.SequenceMatcher(None, wtext, otext, autojunk=False)
print(f"stream similarity ratio: {sm.ratio():.4f}", file=sys.stderr)

# matched triples in stream order: (word_line_idx, oxi_line_idx, char)
MATCHED = []
for blk in sm.get_matching_blocks():
    for k in range(blk.size):
        wi, oi = blk.a + k, blk.b + k
        MATCHED.append((ws[wi][1], os_[oi][1], wtext[wi]))


def do_pagedelta():
    from collections import Counter
    dh = Counter()
    perpage = {}
    for wl, ol, c in MATCHED:
        wp = W[wl]['page']; op = O[ol]['page']
        dh[op - wp] += 1
        perpage.setdefault(wp, Counter())[op - wp] += 1
    tot = sum(dh.values())
    print(f"\n===== per-char PAGE-DELTA (oxi_page - word_page), {tot} matched chars =====")
    for d in sorted(dh):
        bar = '#' * int(60 * dh[d] / tot)
        print(f"   d{d:+d}: {dh[d]:6d} ({100*dh[d]/tot:5.1f}%) {bar}")
    print(f"   delta=0: {100*dh.get(0,0)/tot:.2f}%")
    print("\n   --- per Word-page dominant delta (cascade onset) ---")
    prev = 0
    for wp in sorted(perpage):
        dom = perpage[wp].most_common(1)[0][0]
        flag = '  <<< delta onset' if dom != prev else ''
        print(f"     Wp{wp:2d}: dom d={dom:+d}  {dict(sorted(perpage[wp].items()))}{flag}")
        prev = dom


def do_region(lo, hi):
    """Side-by-side Word vs Oxi visual lines for Word pages lo..hi.
    Each Word line -> the set of Oxi (page,line) its matched chars fell on."""
    # map each (word_line) -> list of oxi_line ids of its matched chars
    from collections import defaultdict, Counter
    wl2ol = defaultdict(Counter)
    for wl, ol, c in MATCHED:
        wl2ol[wl][ol] += 1
    print(f"\n===== REGION Word pages {lo}..{hi}: Word line -> dominant Oxi line/page =====")
    print("  Wp.Wl  Wy    | text (Word)                                  -> Op.Oln Oy   para_idx")
    for wl, ln in enumerate(W):
        if not (lo <= ln['page'] <= hi):
            continue
        wtxt = ''.join(c for c, _, _ in ln['chars'])
        if wl2ol[wl]:
            ol = wl2ol[wl].most_common(1)[0][0]
            o = O[ol]
            opg = f"O{o['page']}.{ol}"
            oy = o['y']
            pa = o['para_idx']; co = o['char_offset']
            note = f"-> {opg} y={oy:.0f} para={pa} off={co}"
        else:
            note = "-> (unmatched)"
        print(f"  W{ln['page']:2d}    {ln['y']:6.1f} | {wtxt[:42]:<42} {note}")


def do_roots():
    """Per Oxi paragraph: first char where Word and Oxi disagree on a line break."""
    from collections import Counter
    def pidx(ol): return O[ol]['para_idx']
    roots = []
    i, n = 0, len(MATCHED)
    while i < n:
        pa = pidx(MATCHED[i][1])
        j = i
        seq = []
        while j < n and pidx(MATCHED[j][1]) == pa:
            seq.append(MATCHED[j]); j += 1
        for k in range(1, len(seq)):
            pwl, pol, _ = seq[k-1]; wl, ol, c = seq[k]
            wb = wl != pwl; ob = ol != pol
            if wb != ob:
                kind = 'OIDASHI(Oxi-overfit)' if wb else 'OIKOMI(Oxi-underfit)'
                roots.append({'para': pa, 'kind': kind, 'char': c, 'prev': seq[k-1][2],
                              'wpage': W[wl]['page'], 'opage': O[ol]['page'],
                              'ctx': ''.join(x[2] for x in seq[max(0,k-12):k+2])})
                break
        i = j
    over = [r for r in roots if 'OIDASHI' in r['kind']]
    under = [r for r in roots if 'OIKOMI' in r['kind']]
    print(f"\n===== per-paragraph CASCADE ROOTS: {len(roots)} paras with a break divergence =====")
    print(f"  {len(over)} OIDASHI (Oxi over-fits), {len(under)} OIKOMI (Oxi under-fits)")
    print("  --- break-char freq ---")
    for c, k in Counter(r['char'] for r in roots).most_common(15):
        t = ' [YAK]' if c in YAK else (' [prohib-start]' if c in PROHIB_START else '')
        print(f"     '{c}' x{k}{t}")
    print("\n  --- OIDASHI roots (Oxi keeps a char Word pushes down = Oxi over-fit) ---")
    for r in over:
        print(f"     para={r['para']:4d} Wp{r['wpage']}/Op{r['opage']} '{r['prev']}'->'{r['char']}'  ...{r['ctx']}")
    print("\n  --- OIKOMI roots (Oxi wraps a char Word keeps = Oxi under-fit) ---")
    for r in under:
        print(f"     para={r['para']:4d} Wp{r['wpage']}/Op{r['opage']} '{r['prev']}'->'{r['char']}'  ...{r['ctx']}")


def do_boundary():
    """Zoom on the two known spill boundaries. For each, walk the Word lines on
    the page-before-spill and the spilling paragraph, char-aligned to Oxi, to
    pin the line where Oxi's page filled one line earlier."""
    from collections import Counter, defaultdict
    wl2ol = defaultdict(Counter)
    for wl, ol, c in MATCHED:
        wl2ol[wl][ol] += 1
    for wpage, label in [(13, 'wi355'), (16, 'wi440')]:
        print(f"\n========== BOUNDARY {label}: Word page {wpage} -> {wpage+1} ==========")
        # Word lines on wpage and wpage+1 (first few)
        print("  Word-line                                       Wpg | Oxi-page.line  Oy     para/off")
        for wl, ln in enumerate(W):
            if ln['page'] not in (wpage, wpage + 1):
                continue
            if ln['page'] == wpage + 1 and ln['y'] > W[0]['y'] + 6 * 14.3:
                # only first ~6 lines of the next page
                pass
            wtxt = ''.join(c for c, _, _ in ln['chars'])
            if wl2ol[wl]:
                ol = wl2ol[wl].most_common(1)[0][0]
                o = O[ol]
                note = f"O{o['page']}.{ol:4d} y={o['y']:6.1f} p{o['para_idx']}/{o['char_offset']}"
                # flag when Word and Oxi pages differ
                diff = '  <<<' if o['page'] != ln['page'] else ''
            else:
                note = "(unmatched)"; diff = ''
            # limit next-page rows
            if ln['page'] == wpage + 1:
                # show only until we pass the spilling para's resync
                pass
            print(f"  {wtxt[:44]:<44}  W{ln['page']:2d} | {note}{diff}")
        # Oxi-side: count lines on Oxi page wpage vs Word page wpage
        wlines_on_p = sum(1 for ln in W if ln['page'] == wpage)
        olines_on_p = sum(1 for ln in O if ln['page'] == wpage)
        print(f"  --> Word page {wpage}: {wlines_on_p} lines  |  Oxi page {wpage}: {olines_on_p} lines")


def do_parastart():
    """Per Oxi paragraph (para_idx), delta = oxi_start_page - word_start_page,
    via char-stream alignment (collision-robust). The gate measures para starts;
    this reproduces it without text-prefix collisions. A para's start = its first
    matched char (char_offset 0 region). Word page from the aligned Word line."""
    from collections import Counter, defaultdict
    # para_idx -> ordered list of (word_line, oxi_line) for its matched chars
    # We need each Oxi line's para_idx; rebuild oxi line -> para_idx
    oline_para = [ln['para_idx'] for ln in O]
    # first matched char per para (in stream order)
    seen = {}
    starts = []  # (para_idx, oxi_page, word_page)
    for wl, ol, c in MATCHED:
        pa = oline_para[ol]
        if pa is None or pa in seen:
            continue
        seen[pa] = True
        starts.append((pa, O[ol]['page'], W[wl]['page']))
    dh = Counter()
    errs = []
    for pa, op, wp in starts:
        d = op - wp
        dh[d] += 1
        if d != 0:
            errs.append((pa, wp, op, d))
    print(f"\n===== PARA-START page delta (oxi - word), {len(starts)} paras matched =====")
    print(f"   histogram: {dict(sorted(dh.items()))}")
    print(f"   n_nonzero: {sum(v for k,v in dh.items() if k!=0)}")
    for pa, wp, op, d in errs:
        # get the para's leading text
        txt = next((O[ol]['text'] for ol in range(len(O)) if O[ol]['para_idx'] == pa), '')
        print(f"     para={pa:4d} Wp{wp:2d}->Op{op:2d} d={d:+d} | {txt[:34]}")


if __name__ == '__main__':
    if '--parastart' in sys.argv:
        do_parastart()
    if '--pagedelta' in sys.argv:
        do_pagedelta()
    if '--region' in sys.argv:
        i = sys.argv.index('--region')
        lo, hi = sys.argv[i+1].split(':')
        do_region(int(lo), int(hi))
    if '--roots' in sys.argv:
        do_roots()
    if '--boundary' in sys.argv:
        do_boundary()
    if not any(a in sys.argv for a in ('--pagedelta', '--region', '--roots', '--boundary')):
        do_pagedelta()
        do_boundary()
