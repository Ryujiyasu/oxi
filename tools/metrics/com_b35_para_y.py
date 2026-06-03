# -*- coding: utf-8 -*-
"""S492u COM ground-truth — measure Word's actual per-paragraph vertical Y on b35
(Information(6)=wdVerticalPositionRelativeToPage, with R30 collapsed-start fix), histogram
the consecutive Y-deltas. docGrid linesAndChars linePitch=350tw=17.5pt; 13 paras have
snapToGrid=false. Decisive test: does Word grid-snap the opt-out paras (all deltas ~17.5 /
multiples) or use natural (~18.0)? If Word has NO ~18.0 single-line deltas, Oxi's natural
18.0 on opt-out paras is the bug. cp932-safe (UTF-8 file, ASCII out, results to file)."""
import sys, json
import win32com.client as win32

DOCX = r'c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\b35123fe8efc_tokumei_08_01.docx'
OUT = r'c:\tmp\b35_word_paray.json'

wdVertPos = 6
wdActiveEndPageNumber = 3

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = 0
try:
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    rows = []
    paras = doc.Paragraphs
    n = paras.Count
    for i in range(1, n + 1):
        p = paras(i)
        rng = p.Range
        # R30: collapsed start range -> page/Y of the paragraph START (not active-end)
        cr = doc.Range(rng.Start, rng.Start)
        try:
            y = float(cr.Information(wdVertPos))
            pg = int(cr.Information(wdActiveEndPageNumber))
        except Exception:
            y, pg = None, None
        # snapToGrid: read from the paragraph's ParagraphFormat if exposed
        try:
            sg = p.Format.SnapToGrid  # bool; -1/True = on, 0/False = off
        except Exception:
            sg = None
        txt = (rng.Text or '')[:14]
        intable = bool(rng.Tables.Count > 0) if hasattr(rng, 'Tables') else None
        rows.append({'i': i, 'page': pg, 'y_pt': y, 'snap': sg, 'in_table': intable,
                     'text': txt})
    doc.Close(False)
finally:
    word.Quit()

# deltas within same page, consecutive paras
hist = {}
deltas = []
for a, b in zip(rows, rows[1:]):
    if a['page'] == b['page'] and a['y_pt'] is not None and b['y_pt'] is not None:
        d = round(b['y_pt'] - a['y_pt'], 2)
        if 5.0 < d < 60.0:
            deltas.append((a['i'], d, a['snap'], a['in_table']))
            k = round(d * 2) / 2
            hist[k] = hist.get(k, 0) + 1

with open(OUT, 'w', encoding='utf-8') as f:
    json.dump({'rows': rows, 'deltas': deltas}, f, ensure_ascii=False, indent=1)

print("b35 Word per-paragraph Y-delta histogram (0.5pt bins); grid=17.5pt")
for k in sorted(hist):
    mark = ' <-17.5 GRID' if abs(k - 17.5) < 0.26 else (' <-18.0' if abs(k - 18.0) < 0.26 else '')
    print("  %5.1fpt x%-3d %s%s" % (k, hist[k], '#' * hist[k], mark))

print("\nopt-out (snap=False) paragraph deltas (Word's actual advance for them):")
for i, d, sg, it in deltas:
    if sg == 0 or sg is False:
        print("  para i=%d  delta=%.2fpt  in_table=%s" % (i, d, it))
n_optout = sum(1 for r in rows if r['snap'] == 0 or r['snap'] is False)
print("\ntotal paras=%d  snap=False=%d  (XML reported 13 <snapToGrid val=0>)" % (len(rows), n_optout))
print("wrote", OUT)
