# -*- coding: utf-8 -*-
"""S559 wheel-(b) — COM-measure the REAL 3a4f p19 cell (労働基準法…, jc=left
left=459) wrap + cell geometry, to test whether Word RESERVES the default cellMar
for it. ⑦ reserves (S558: wraps 37 = gridCol−cellMar−firstLine). If real p19 also
reserves → the ⑦-vs-p19 asymmetry is NOT real (the {1:1323} was pure cascade, the
jc gate is a cascade-avoidance proxy). If p19 does NOT reserve → a real structural
discriminator distinguishes the two reserving/non-reserving cells.
"""
import os
import sys
import win32com.client as w32

DOCX = r'c:\tmp\3a4f9f.docx'
sys.stdout.reconfigure(encoding='utf-8')

# the two cells: anchor -> label
CELLS = {
    '⑦ (jc=both left=0)': u'常に整理整頓に努め',
    'p19 (jc=left left=459)': u'労働基準法においては、労働時間',
}


def measure(wdoc, anchor):
    rng = wdoc.Content
    f = rng.Find
    f.Text = anchor
    if not f.Execute():
        return None
    p = rng.Paragraphs(1).Range
    s, e = p.Start, p.End
    t = p.Text
    # line-1 char count + x of first/last char on L1 + cell left edge
    rows = []
    for i in range(min(e - s, 80)):
        ch = t[i] if i < len(t) else ''
        if ch in ('\r', '\n', '\x07', '\x0b'):
            continue
        x = wdoc.Range(s + i, s + i).Information(5)  # horiz pos rel to page
        y = wdoc.Range(s + i, s + i).Information(6)
        rows.append((ch, x, y))
    # group lines
    lines = []
    cur = []
    y0 = None
    for ch, x, y in rows:
        if y0 is None:
            y0 = y
        if abs(y - y0) > 0.5:
            lines.append(cur)
            cur = []
            y0 = y
        cur.append((ch, x))
    if cur:
        lines.append(cur)
    return lines


def main():
    word = w32.DispatchEx('Word.Application')
    word.Visible = False
    try:
        wdoc = word.Documents.Open(os.path.abspath(DOCX), ReadOnly=True)
        try:
            for label, anchor in CELLS.items():
                lines = measure(wdoc, anchor)
                print('\n=== %s ===' % label)
                if not lines:
                    print('  not found')
                    continue
                for li, ln in enumerate(lines[:3]):
                    x0 = ln[0][1]
                    xl = ln[-1][1]
                    print('  L%d n=%d  x0=%.2f  last_x=%.2f  text=%s'
                          % (li + 1, len(ln), x0, xl, ''.join(c for c, _ in ln)[:50]))
                print('  total lines=%d' % len(lines))
        finally:
            wdoc.Close(False)
    finally:
        word.Quit()


if __name__ == '__main__':
    main()
