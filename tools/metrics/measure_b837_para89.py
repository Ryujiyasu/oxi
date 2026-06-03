"""S492c — b837 para 89 (p6->p7 +18pt) char-width vs break-policy decomposition.
Oxi fits para 89 in 4 lines, Word in 3 (Oxi ~1 char/line fewer). Decisive test:
compare per-char cumulative X on line 1 (Word Information(5) vs Oxi dump). If Oxi's
X grows faster and reaches the margin 1 char earlier -> CHAR-WIDTH cause (font tables).
If Oxi breaks while still short of where Word breaks -> BREAK-POLICY (S476 capacity).
No code change.
"""
import os, glob, subprocess, json, re
import win32com.client as w32

BIN = os.path.abspath('tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe')
WD_VPOS, WD_HPOS = 6, 5
DOCX = os.path.abspath(glob.glob('tools/golden-test/documents/docx/b837*.docx')[0])

# --- Oxi: para 89, line-by-line chars + x + w ---
subprocess.run([BIN, DOCX, 'c:/tmp/_b837_x', '--dump-layout=c:/tmp/_b837.json'],
               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
d = json.load(open('c:/tmp/_b837.json', encoding='utf-8'))
els89 = [e for pg in d['pages'] for e in pg['elements']
         if e['type'] == 'text' and e.get('para_idx') == 89]
els89.sort(key=lambda e: (round(e['y'], 1), e['x']))
oxi_text = ''.join(e['text'] for e in els89)
lines = {}
for e in els89:
    lines.setdefault(round(e['y'], 1), []).append(e)
oxi_lines = [sorted(lines[y], key=lambda e: e['x']) for y in sorted(lines)]
oxi_counts = [sum(len(e['text']) for e in ln) for ln in oxi_lines]
print("Oxi para89: %d lines, counts=%s" % (len(oxi_lines), oxi_counts))
# line1 cumulative x (left edge of each char-element)
l1 = oxi_lines[0]
oxi_x0 = l1[0]['x']
print("Oxi L1: %d els, x0=%.2f, last char right=%.2f" % (len(l1), oxi_x0, l1[-1]['x'] + l1[-1]['w']))
prefix = re.sub(r'\s', '', oxi_text)[:8]
print("para89 prefix:", repr(oxi_text[:24]))

# --- Word: find matching para, line-1 per-char X ---
word = w32.DispatchEx('Word.Application'); word.Visible = False
try:
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    try:
        target = None
        for p in doc.Paragraphs:
            clean = p.Range.Text.replace('\r', '').replace('\x07', '').replace('\n', '')
            if re.sub(r'\s', '', clean)[:8] == prefix:
                target = p; break
        if target is None:
            print("Word: para89 not matched by prefix", repr(prefix))
        else:
            rng = target.Range; txt = rng.Text; start = rng.Start
            y0 = doc.Range(start, start).Information(WD_VPOS)
            wx = []  # (char, x) for L1
            wcounts = []
            cur = 0; prev_y = y0
            for i in range(len(txt)):
                ch = txt[i]
                if ch in ('\r', '\n', '\x07'):
                    continue
                y = doc.Range(start + i, start + i).Information(WD_VPOS)
                x = doc.Range(start + i, start + i).Information(WD_HPOS)
                if y > prev_y + 2:
                    wcounts.append(cur); cur = 0; prev_y = y
                if not wcounts:  # still on L1
                    wx.append((ch, round(x, 2)))
                cur += 1
            wcounts.append(cur)
            print("\nWord para89: %d lines, counts=%s" % (len(wcounts), wcounts))
            print("Word L1: %d chars, x0=%.2f, last char x=%.2f" % (len(wx), wx[0][1], wx[-1][1]))
            print("\n=== LINE COUNT ===  Word %d lines vs Oxi %d lines (Oxi %+d)" %
                  (len(wcounts), len(oxi_lines), len(oxi_lines) - len(wcounts)))
            print("L1 chars: Word %d vs Oxi %d" % (wcounts[0], oxi_counts[0]))
            # cumulative X comparison at matched indices (relative to x0)
            print("\nidx  Word_dx   Oxi_dx   (cumulative from line start; char)")
            for i in range(min(len(wx), len(l1), 40)):
                wdx = wx[i][1] - wx[0][1]
                # oxi element i left edge minus x0
                odx = l1[i]['x'] - oxi_x0 if i < len(l1) else None
                mark = ''
                if odx is not None and abs(odx - wdx) > 1.5:
                    mark = '  <-- diverge %.2f' % (odx - wdx)
                print("%3d  %7.2f  %7s   %s%s" % (i, wdx, ('%.2f' % odx) if odx is not None else '-', wx[i][0], mark))
    finally:
        doc.Close(False)
finally:
    word.Quit()
