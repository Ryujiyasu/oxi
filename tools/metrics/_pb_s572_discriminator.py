# -*- coding: utf-8 -*-
"""Which lines does Word actually compress under the legacy no-type-docGrid oikomi (S572)?

golden ikujidetail (compat 11, no-type docGrid linePitch=286 charSpace=-3531, jc=left,
compressPunctuation). S572 was derived on body paragraph i=199, where Word renders a
mid-line 、 at 9.0pt and a line-end 。 at 5.25 against a slack line's 11.25 -- demand-driven
oikomi. But paragraph i=17 「３　配偶者が従業員と同じ日から…」 (ind leftChars=100 left=430
hangingChars=100 hanging=220, sz 22) is the opposite: Word breaks it 42 / 41 / 41 / 4 while
Oxi packs 43 / 42 / 43, every Oxi line 1.6pt past the content right edge of 538.6. Line 2
carries 、x4 and a ・ -- ample material -- and Word still refuses to compress. So the S572
credit needs a discriminator beyond "marks are present".

Readout: Information(5) (wdHorizontalPositionRelToPage) per character, so each character's
advance is the difference to the next. Both paragraphs are printed with their line breaks
marked (a drop in x = a new line), and every mark's advance is called out against the
paragraph's plain-character advance.
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import win32com.client as win32

DOCX = os.path.abspath('tools/golden-test/documents/docx/ikujidetail_002197815.docx')
MARKS = '、。，．・：；（）「」［］'
app = win32.DispatchEx('Word.Application'); app.Visible = False
try:
    d = app.Documents.Open(DOCX, ReadOnly=True)
    for target in (199, 17, 5):
        p = d.Paragraphs(target)
        rng = p.Range
        n = rng.Characters.Count
        rows = []
        for k in range(1, min(n, 200) + 1):
            c = rng.Characters(k)
            rows.append((c.Text, c.Information(5), c.Information(6)))
        print(f'=== para i={target} chars={n} ===', flush=True)
        line = 0
        prev_x = None
        prev_y = None
        plain = []
        marks = []
        for i, (ch, x, y) in enumerate(rows):
            if prev_y is not None and y > prev_y + 1.0:
                print(f'  --- line {line} ends ---')
                line += 1
            adv = None
            if i + 1 < len(rows) and rows[i + 1][2] <= y + 1.0:
                adv = round(rows[i + 1][1] - x, 2)
            if adv is not None:
                (marks if ch in MARKS else plain).append(adv)
            prev_x, prev_y = x, y
        from collections import Counter
        print('  plain advances:', Counter(plain).most_common(4))
        print('  mark  advances:', Counter(marks).most_common(6))
        print('  n lines:', line + 1, flush=True)
    d.Close(False)
finally:
    app.Quit()
