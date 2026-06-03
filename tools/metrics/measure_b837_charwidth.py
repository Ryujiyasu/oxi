# -*- coding: utf-8 -*-
"""S492j — char-width check on a b837 over-pack body para (Oxi L0 = Word L0 + 1, same
line count => not compression, candidate char-width). Picks such a para from the
dataset, then compares Oxi per-char advance (dump width) vs Word per-char advance
(Info5 diff) to find which chars Oxi renders NARROWER than Word (the +1 source).
cp932-safe: UTF-8 file, ASCII output (codepoints, not glyphs).
"""
import os, glob, subprocess, json, re
import win32com.client as w32

BIN = os.path.abspath('tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')
DOCX = os.path.abspath(glob.glob('tools/golden-test/documents/docx/b837*.docx')[0])
WD_VPOS, WD_HPOS = 6, 5

ds = json.load(open('c:/tmp/b837_breakdataset.json', encoding='utf-8'))
# pick a 2-line over-pack para: word==2 lines, natural==2 lines, word[0]==natural[0]-1
target = None
for w in ds:
    nat = w.get('natural')
    if nat and len(w['word']) == 2 and len(nat) == 2 and nat[0] == w['word'][0] + 1:
        target = w; break
if not target:
    print("no clean 2-line over-pack para found"); raise SystemExit
key = target['key']
print("target para key prefix (codepoints):", [hex(ord(c)) for c in key[:8]])
print("Word lines=%s natural=%s flatk=%s" % (target['word'], target['natural'], target['flatk']))

# Oxi per-char widths (default/flatK render)
subprocess.run([BIN, DOCX, 'c:/tmp/_cw', '--dump-layout=c:/tmp/_cw.json'],
               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
d = json.load(open('c:/tmp/_cw.json', encoding='utf-8'))
oxi_chars = None
for pi, pg in enumerate(d['pages']):
    pm = {}
for pgi, pg in enumerate(d['pages']):
    pass
# collect by para
pmap = {}
for pgi, pg in enumerate(d['pages']):
    for e in pg['elements']:
        if e['type'] == 'text' and e.get('para_idx') is not None:
            pmap.setdefault(e['para_idx'], []).append((pgi, e))
for pi, els in pmap.items():
    els.sort(key=lambda t: (t[0], round(t[1]['y'], 1), t[1]['x']))
    txt = ''.join(e['text'] for _, e in els)
    if re.sub(r'\s', '', txt)[:14] == key:
        # L0 chars + widths
        y0 = sorted(set((p, round(e['y'], 1)) for p, e in els))[0]
        l0 = [e for p, e in els if (p, round(e['y'], 1)) == y0]
        l0.sort(key=lambda e: e['x'])
        oxi_chars = [(e['text'], round(e['w'], 2)) for e in l0]
        break

# Word per-char advances on L0
word = w32.DispatchEx('Word.Application'); word.Visible = False
word_adv = []
try:
    wdoc = word.Documents.Open(DOCX, ReadOnly=True)
    try:
        for p in wdoc.Paragraphs:
            clean = p.Range.Text.replace('\r', '').replace('\x07', '').replace('\n', '')
            if re.sub(r'\s', '', clean)[:14] != key:
                continue
            rng = p.Range; txt = rng.Text; start = rng.Start
            y0 = wdoc.Range(start, start).Information(WD_VPOS)
            seq = []
            for i in range(len(txt)):
                ch = txt[i]
                if ch in ('\r', '\n', '\x07'):
                    continue
                y = wdoc.Range(start + i, start + i).Information(WD_VPOS)
                if y > y0 + 2:
                    break
                x = wdoc.Range(start + i, start + i).Information(WD_HPOS)
                seq.append((ch, x))
            for j in range(len(seq) - 1):
                word_adv.append((seq[j][0], round(seq[j + 1][1] - seq[j][1], 2)))
            break
    finally:
        wdoc.Close(False)
finally:
    word.Quit()

print("\nL0 per-char: idx | codept | Oxi_w | Word_adv | diff(Oxi-Word)")
n = min(len(oxi_chars or []), len(word_adv))
sum_oxi = sum_word = 0.0
for i in range(n):
    oc, ow = oxi_chars[i]
    wc, wa = word_adv[i]
    sum_oxi += ow; sum_word += wa
    flag = '  <-- Oxi NARROWER' if ow < wa - 0.3 else ('  Oxi wider' if ow > wa + 0.3 else '')
    cp = '+'.join(hex(ord(c)) for c in oc)
    print("  %2d  %-14s  %5.2f  %6.2f  %+5.2f  %s" % (i, cp, ow, wa, ow - wa, flag))
print("\nL0 cumulative: Oxi=%.2f  Word=%.2f  (Oxi narrower by %.2f over %d chars)" %
      (sum_oxi, sum_word, sum_word - sum_oxi, n))
