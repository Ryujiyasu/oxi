# -*- coding: utf-8 -*-
"""S492v — COM-measure b35's ACTUAL intra-cell line pitch. b35 has adjustLineHeightInTable
=TRUE and all 13 snapToGrid=0 paras inside cells. Decisive: does Word use NATURAL (~14.0pt
for 10.5pt, opt-out wins) or GRID-snapped (17.5pt, flag wins) for these cells? Oxi outputs
17.5. For each table cell whose text wraps to >=2 lines, sample per-char Y (Information(6)),
find distinct line levels, report the pitch + font size. Writes ASCII results to file."""
import json
import win32com.client as win32

DOCX = r'c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\b35123fe8efc_tokumei_08_01.docx'
OUT = r'c:\tmp\b35_cell_pitch_word.json'
wdVertPos = 6

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0
out = []
try:
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    for ti in range(1, doc.Tables.Count + 1):
        tbl = doc.Tables(ti)
        nr = tbl.Rows.Count
        for ri in range(1, nr + 1):
            try:
                ncol = tbl.Rows(ri).Cells.Count
            except Exception:
                continue
            for ci in range(1, ncol + 1):
                try:
                    cell = tbl.Cell(ri, ci)
                except Exception:
                    continue
                rng = cell.Range
                start, end = rng.Start, rng.End - 1
                if end - start < 2:
                    continue
                # font size of the cell text
                try:
                    fs = float(rng.Font.Size)
                except Exception:
                    fs = None
                ys = []
                i = start
                while i < end:
                    cr = doc.Range(i, i)
                    try:
                        ys.append(float(cr.Information(wdVertPos)))
                    except Exception:
                        pass
                    i += 1
                if not ys:
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
                out.append({'table': ti, 'row': ri, 'col': ci, 'font_size': fs,
                            'n_lines': len(merged), 'pitch': med, 'deltas': deltas})
    doc.Close(False)
finally:
    word.Quit()

with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, indent=1)

# Word natural cell pitch (from repro, no-flag): 9->12,10->13,10.5->14,11->14,12->16
nat = {9.0: 12.0, 10.0: 13.0, 10.5: 14.0, 11.0: 14.0, 12.0: 16.0}
print("b35 ACTUAL Word cell line pitch (adjustLineHeightInTable=TRUE, cell paras snapToGrid=0):")
print(" tbl row col  fs    nlines  Word_pitch   natural   grid17.5   verdict")
for r in out:
    fs = round(r['font_size'], 1) if r['font_size'] else 0
    p = r['pitch']
    n = nat.get(fs)
    vn = abs(p - n) < 1.0 if n else False
    vg = abs(p - 17.5) < 1.0
    verdict = 'NATURAL' if vn else ('GRID17.5' if vg else '?')
    print("  %d  %2d  %2d  %4.1f   %3d     %5.2f       %s       %s     %s" % (
        r['table'], r['row'], r['col'], fs, r['n_lines'], p,
        ('%.1f' % n) if n else '-', '17.5', verdict))
# aggregate verdict
import collections
allp = [r['pitch'] for r in out]
print("\nall cell pitches:", sorted(collections.Counter(round(p) for p in allp).items()))
print("wrote", OUT)
