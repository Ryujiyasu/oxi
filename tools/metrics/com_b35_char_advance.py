# -*- coding: utf-8 -*-
"""S492w-h — COM-measure Word's HORIZONTAL char advance in b35 cells (Information(5)=
wdHorizontalPositionRelativeToPage). Oxi compresses 10.5pt cell chars to 9.84pt (charSpace
=-2714 grid). Decisive: does Word use ~10.5+ (1em+, NOT compressed -> Oxi over-compresses =
the horizontal char-grid bug) or ~9.84 (compressed, matching Oxi)? Sample consecutive char X
in multi-char in-table 10.5pt paragraphs. Writes ASCII results. cp932-safe."""
import json
import win32com.client as win32

DOCX = r'c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\b35123fe8efc_tokumei_08_01.docx'
OUT = r'c:\tmp\b35_char_adv_word.json'
wdHorizPos = 5
wdVertPos = 6

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0
out = []
try:
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    paras = doc.Paragraphs
    for i in range(1, paras.Count + 1):
        p = paras(i)
        rng = p.Range
        try:
            if rng.Tables.Count == 0:
                continue
            fs = float(rng.Font.Size)
        except Exception:
            continue
        if abs(fs - 10.5) > 0.1:
            continue
        start, end = rng.Start, rng.End - 1
        if end - start < 6:
            continue
        # sample first ~14 chars of the FIRST line: collect (x,y); advance where y constant
        pts = []
        for j in range(start, min(start + 20, end)):
            cr = doc.Range(j, j)
            try:
                x = float(cr.Information(wdHorizPos))
                y = float(cr.Information(wdVertPos))
                pts.append((j, x, y))
            except Exception:
                pass
        if len(pts) < 4:
            continue
        # advances within the same line (same y within 2pt), positive dx
        y0 = pts[0][2]
        advs = []
        for a, b in zip(pts, pts[1:]):
            if abs(a[2] - y0) < 2 and abs(b[2] - y0) < 2:
                dx = b[1] - a[1]
                if 0 < dx < 2 * fs:
                    advs.append(round(dx, 2))
        if advs:
            import statistics
            out.append({'i': i, 'fs': fs, 'n_adv': len(advs),
                        'median_adv': round(statistics.median(advs), 2),
                        'advs': advs[:10]})
    doc.Close(False)
finally:
    word.Quit()

with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, indent=1)

import statistics
allmed = [r['median_adv'] for r in out]
print("Word b35 10.5pt CELL horizontal char advance (Information(5)); 1em=10.5pt; Oxi=9.84pt")
for r in out[:20]:
    print("  para i=%d  median_adv=%.2fpt  n=%d  advs=%s" % (r['i'], r['median_adv'], r['n_adv'], r['advs']))
if allmed:
    print("\nWORD overall median cell char advance (10.5pt) = %.2fpt" % statistics.median(allmed))
    print("OXI = 9.84pt.  em = 10.5pt.")
    wm = statistics.median(allmed)
    print("VERDICT: Word %.2f vs Oxi 9.84 -> Oxi %s by %.2fpt/char (%.1f%%)" % (
        wm, 'OVER-COMPRESSES' if wm > 9.84 else 'matches/expands', wm - 9.84, (wm - 9.84) / wm * 100))
print("wrote", OUT)
