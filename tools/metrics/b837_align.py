# -*- coding: utf-8 -*-
"""S492f — clean text-based Word<->Oxi paragraph alignment for b837 (stop conflating
paragraphs; cp932-safe: UTF-8 file, results written to JSON + ASCII summary, no
Japanese to console). For each aligned paragraph: Word jc + indent + line count +
L1 chars + first page, vs Oxi line count + L1 chars + first page + max line xend
(boundary 524.4). Flags paras where Oxi diverges (line count, page, or overflow).
"""
import os, glob, subprocess, json, re
import win32com.client as w32

BIN = os.path.abspath('tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')
WD_VPOS, WD_HPOS = 6, 5
ALIGN = {0: 'left', 1: 'center', 2: 'right', 3: 'both', 4: 'distribute'}
DOCX = os.path.abspath(glob.glob('tools/golden-test/documents/docx/b837*.docx')[0])
BOUNDARY = 524.4


def norm(s):
    return re.sub(r'\s', '', s)


# --- Oxi paragraphs ---
subprocess.run([BIN, DOCX, 'c:/tmp/_b837align', '--dump-layout=c:/tmp/_b837align.json'],
               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
d = json.load(open('c:/tmp/_b837align.json', encoding='utf-8'))
opmap = {}
for pgi, pg in enumerate(d['pages']):
    for e in pg['elements']:
        if e['type'] == 'text' and e.get('para_idx') is not None:
            opmap.setdefault(e['para_idx'], []).append((pgi, e))
oxi = []
for pi in sorted(opmap):
    els = sorted(opmap[pi], key=lambda t: (t[0], round(t[1]['y'], 1), t[1]['x']))
    txt = ''.join(e['text'] for _, e in els)
    lines = {}
    for pgi, e in els:
        lines.setdefault((pgi, round(e['y'], 1)), []).append(e)
    lk = sorted(lines)
    l1 = lines[lk[0]]
    maxxend = max((e['x'] + e['w']) for _, e in els)
    oxi.append({'text': txt, 'norm': norm(txt), 'nlines': len(lines),
                'l1': sum(len(e['text']) for e in l1), 'page1': lk[0][0] + 1,
                'maxxend': round(maxxend, 1)})

# --- Word paragraphs ---
word = w32.DispatchEx('Word.Application'); word.Visible = False
wlist = []
try:
    wdoc = word.Documents.Open(DOCX, ReadOnly=True)
    try:
        for p in wdoc.Paragraphs:
            rng = p.Range; txt = rng.Text
            clean = txt.replace('\r', '').replace('\x07', '').replace('\n', '')
            if len(clean) < 4:
                continue
            start = rng.Start
            y0 = wdoc.Range(start, start).Information(WD_VPOS)
            pg1 = wdoc.Range(start, start).Information(3)  # wdActiveEndPageNumber
            nlines = 1; l1 = 0; prev_y = y0; counted_l1 = False
            for i in range(len(txt)):
                if txt[i] in ('\r', '\n', '\x07'):
                    continue
                y = wdoc.Range(start + i, start + i).Information(WD_VPOS)
                if y > prev_y + 2:
                    nlines += 1; prev_y = y; counted_l1 = True
                if not counted_l1:
                    l1 += 1
            wlist.append({'norm': norm(clean), 'text': clean[:20],
                          'jc': ALIGN.get(p.Alignment, str(p.Alignment)),
                          'li': round(p.LeftIndent, 1), 'fli': round(p.FirstLineIndent, 1),
                          'nlines': nlines, 'l1': l1, 'page1': int(pg1)})
    finally:
        wdoc.Close(False)
finally:
    word.Quit()

# --- align by norm-text prefix (12) ---
def key(s):
    return s[:12]
oidx = {key(o['norm']): o for o in oxi}
rows = []
for w in wlist:
    o = oidx.get(key(w['norm']))
    rows.append((w, o))

out = []
for w, o in rows:
    rec = {'word': w, 'oxi': ({'nlines': o['nlines'], 'l1': o['l1'], 'page1': o['page1'],
                              'maxxend': o['maxxend']} if o else None)}
    out.append(rec)
json.dump(out, open('c:/tmp/b837_align.json', 'w'), ensure_ascii=False, indent=1)

# --- ASCII summary: divergences ---
print("=== b837 Word<->Oxi paragraph divergences (matched=%d/%d Word paras) ===" %
      (sum(1 for _, o in rows if o), len(rows)))
print("idx jc     li/fli  Wln Oln  Wl1 Ol1  Wpg Opg  oxi_xend  flags")
for i, (w, o) in enumerate(rows):
    if not o:
        print("%3d %-6s %5.0f/%-3.0f  unmatched" % (i, w['jc'], w['li'], w['fli']))
        continue
    flags = []
    if w['nlines'] != o['nlines']:
        flags.append('NLINES %+d' % (o['nlines'] - w['nlines']))
    if w['page1'] != o['page1']:
        flags.append('PAGE %+d' % (o['page1'] - w['page1']))
    if o['maxxend'] > BOUNDARY + 1:
        flags.append('OVERFLOW %.1f' % (o['maxxend'] - BOUNDARY))
    if w['l1'] != o['l1']:
        flags.append('L1 %+d' % (o['l1'] - w['l1']))
    if flags:
        print("%3d %-6s %5.0f/%-3.0f  %3d %3d  %3d %3d  %3d %3d  %8.1f  %s" %
              (i, w['jc'], w['li'], w['fli'], w['nlines'], o['nlines'], w['l1'], o['l1'],
               w['page1'], o['page1'], o['maxxend'], '; '.join(flags)))
print("\n(full data: c:/tmp/b837_align.json)")
