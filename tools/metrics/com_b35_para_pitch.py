# -*- coding: utf-8 -*-
"""S492v (robust) — COM-measure b35 per-PARAGRAPH intra-line pitch by sampling each
paragraph's char Y (Information(6)). Captures cell paragraphs via Range.Tables.Count>0
without fragile Cell() indexing. Decisive: for in-table paras (snapToGrid=0,
adjustLineHeightInTable=TRUE) does Word use NATURAL (~14.0/10.5pt) or GRID (17.5)?
Writes ASCII results to file."""
import json
import win32com.client as win32

DOCX = r'c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\b35123fe8efc_tokumei_08_01.docx'
OUT = r'c:\tmp\b35_para_pitch_word.json'
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
        start, end = rng.Start, rng.End - 1
        if end - start < 2:
            continue
        try:
            intable = rng.Tables.Count > 0
        except Exception:
            intable = None
        try:
            fs = float(rng.Font.Size)
        except Exception:
            fs = None
        nchars = end - start
        ys = []
        j = start
        # sample up to ~80 chars to keep runtime reasonable
        step = max(1, nchars // 80)
        while j < end:
            cr = doc.Range(j, j)
            try:
                ys.append(float(cr.Information(wdVertPos)))
            except Exception:
                pass
            j += step
        if len(ys) < 2:
            continue
        levels = sorted(set(round(y, 1) for y in ys))
        merged = []
        for y in levels:
            if merged and y - merged[-1] < 3:
                continue
            merged.append(y)
        if len(merged) < 2:
            continue
        deltas = [round(merged[k] - merged[k - 1], 2) for k in range(1, len(merged))]
        sd = sorted(deltas)
        med = sd[len(sd) // 2]
        out.append({'i': i, 'in_table': bool(intable) if intable is not None else None,
                    'font_size': fs, 'nchars': nchars, 'n_lines': len(merged),
                    'pitch': med, 'deltas': deltas})
    doc.Close(False)
finally:
    word.Quit()

with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, indent=1)

nat = {9.0: 12.0, 10.0: 13.0, 10.5: 14.0, 11.0: 14.0, 12.0: 16.0}
print("b35 Word per-paragraph intra-line pitch (>=2 wrapped lines):")
print("  i  in_tbl  fs    nchars nlines  Word_pitch  natural  verdict")
ncell_nat = ncell_grid = 0
for r in out:
    fs = round(r['font_size'], 1) if r['font_size'] else 0
    p = r['pitch']
    n = nat.get(fs)
    vn = (n is not None and abs(p - n) < 1.2)
    vg = abs(p - 17.5) < 1.0
    verdict = 'NATURAL' if vn else ('GRID17.5' if vg else '?')
    if r['in_table']:
        if vn:
            ncell_nat += 1
        elif vg:
            ncell_grid += 1
    print("  %2d   %s   %4.1f   %4d   %3d     %5.2f      %s     %s  %s" % (
        r['i'], 'T' if r['in_table'] else '.', fs, r['nchars'], r['n_lines'], p,
        ('%.1f' % n) if n else '-', verdict, sorted(set(r['deltas']))[:5]))
print("\nIN-TABLE wrapping paras: NATURAL=%d  GRID17.5=%d" % (ncell_nat, ncell_grid))
print(">>> if NATURAL dominates, Word uses small natural height in b35 cells => Oxi's 17.5 is the BUG")
print("wrote", OUT)
