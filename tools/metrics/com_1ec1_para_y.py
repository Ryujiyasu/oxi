# -*- coding: utf-8 -*-
"""S492-1ec1 (S469 investigation): COM-measure Word's per-paragraph page-Y for 1ec1 (all
paras incl. those textboxes anchor to), Information(6) with R30 collapsed start. The 7 p1
textboxes anchor to blocks 2,3,11,13,28 and Oxi places them at block_y[anchor]+posY. If Oxi's
body paragraph Ys are MORE SPREAD than Word's (top too high, bottom too low), the box vertical
scatter (+3 top -> -8 bottom) is INHERITED from the body-Y error, not box-specific. Writes
page-1 paragraph Ys to a file (ASCII). cp932-safe (UTF-8 file, ASCII out, results to JSON)."""
import json
import win32com.client as win32

DOCX = None
import glob
DOCX = glob.glob(r'c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\1ec1091177b1*.docx')[0]
OUT = r'c:\tmp\1ec1_para_y.json'
wdVertPos = 6
wdActiveEndPageNumber = 3

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0
rows = []
try:
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    paras = doc.Paragraphs
    for i in range(1, paras.Count + 1):
        p = paras(i)
        rng = p.Range
        cr = doc.Range(rng.Start, rng.Start)
        try:
            y = float(cr.Information(wdVertPos))
            pg = int(cr.Information(wdActiveEndPageNumber))
        except Exception:
            y, pg = None, None
        intbl = False
        try:
            intbl = rng.Tables.Count > 0
        except Exception:
            pass
        nchars = rng.End - rng.Start
        rows.append({'i': i, 'page': pg, 'y_pt': y, 'in_table': bool(intbl), 'nchars': nchars})
    doc.Close(False)
finally:
    word.Quit()

with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(rows, f, ensure_ascii=False, indent=1)

p1 = [r for r in rows if r['page'] == 1 and r['y_pt'] is not None]
print("1ec1 Word page-1 paragraphs: %d (total %d)" % (len(p1), len(rows)))
print("idx  Y_pt   in_tbl  nchars")
for r in p1:
    print("  %3d  %6.1f   %s    %d" % (r['i'], r['y_pt'], 'T' if r['in_table'] else '.', r['nchars']))
print("wrote", OUT)
