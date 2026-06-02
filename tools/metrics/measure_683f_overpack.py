"""S492 decisive value test — does Oxi over-pack the jc=left wrapping paragraphs
of bottom-N docs (683f, 0e7af) vs Word? If yes, F1 (gate punct-compression break
to justified paras) is a real BOTTOM-N lever, not just a correctness fix.

For each doc: Word per-paragraph (alignment, full text, L1 char count) for wrapping
paras; Oxi per-paragraph L1 char count from --dump-layout; matched by text prefix.
"""
import os, glob, re, subprocess, json
import win32com.client as w32

WD_VPOS, WD_HPOS = 6, 5
ALIGN = {0: 'left', 1: 'center', 2: 'right', 3: 'both', 4: 'distribute'}
BIN = os.path.abspath('tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')

DOCS = {}
for stem in ['683ffcab86e2', '0e7af1ae8f21', 'd77a58485f16', 'b837808d0555']:
    g = glob.glob('pipeline_data/golden_per_page/%s*_p1.docx' % stem)
    if g:
        DOCS[stem] = g[0]


def norm(s):
    return re.sub(r'\s', '', s)[:12]


def oxi_para_l1(docx):
    out = 'c:/tmp/_ovp_%s.json' % os.path.basename(docx).split('_')[0]
    subprocess.run([BIN, docx, 'c:/tmp/_ovp_x', '--dump-layout=' + out],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    d = json.load(open(out, encoding='utf-8'))
    paras = {}
    for pg in d['pages']:
        for e in pg['elements']:
            if e['type'] != 'text':
                continue
            pi = e.get('para_idx')
            if pi is None:
                continue
            paras.setdefault(pi, []).append(e)
    res = {}  # para_idx -> (l1_count, full_prefix)
    for pi, els in paras.items():
        ys = sorted(set(round(e['y'], 1) for e in els))
        y0 = ys[0]
        l1 = [e for e in els if abs(e['y'] - y0) < 2]
        l1.sort(key=lambda e: e['x'])
        full = sorted(els, key=lambda e: (round(e['y'], 1), e['x']))
        l1n = len(''.join(e['text'] for e in l1))
        pref = norm(''.join(e['text'] for e in full))
        res[pi] = (l1n, pref)
    return res


word = w32.DispatchEx('Word.Application'); word.Visible = False
try:
    for stem, docx in DOCS.items():
        abspath = os.path.abspath(docx)
        oxi = oxi_para_l1(abspath)
        oxi_by_pref = {pref: (l1, pi) for pi, (l1, pref) in oxi.items()}
        doc = word.Documents.Open(abspath, ReadOnly=True)
        rows = []
        try:
            for p in doc.Paragraphs:
                rng = p.Range; txt = rng.Text
                clean = txt.replace('\r', '').replace('\x07', '').replace('\n', '')
                if len(clean) < 20:
                    continue
                start, end = rng.Start, rng.End
                y0 = doc.Range(start, start).Information(WD_VPOS)
                yN = doc.Range(max(start, end - 1), max(start, end - 1)).Information(WD_VPOS)
                if (yN - y0) <= 2:
                    continue  # not wrapping
                al = ALIGN.get(p.Alignment, str(p.Alignment))
                n = 0
                for i in range(len(txt)):
                    ch = txt[i]
                    if ch in ('\r', '\n', '\x07'):
                        continue
                    if doc.Range(start + i, start + i).Information(WD_VPOS) > y0 + 2:
                        break
                    n += 1
                pref = norm(clean)
                ox = oxi_by_pref.get(pref)
                oxn = ox[0] if ox else None
                rows.append((al, n, oxn, clean[:16]))
        finally:
            doc.Close(False)
        print("\n=== %s ===" % stem)
        print("%-6s %5s %5s %6s  %s" % ('align', 'Word', 'Oxi', 'delta', 'text'))
        for al, n, oxn, t in rows:
            if oxn is None:
                print("%-6s %5d %5s %6s  %s  (no oxi match)" % (al, n, '-', '-', t))
            else:
                d = oxn - n
                flag = '  <-- OVER-PACK' if (al != 'both' and d > 0) else ('' if d == 0 else '  d=%+d' % d)
                print("%-6s %5d %5d %+6d  %s%s" % (al, n, oxn, d, t, flag))
finally:
    word.Quit()
