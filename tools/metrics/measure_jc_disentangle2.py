"""S492 follow-up — WHAT mechanism gives jc=both its extra chars?

Mid-line punct measured 12.0 (natural) even under jc=both, yet jc=both fits 38
(comma) vs jc=left's 36. Candidates: (a) compression localized near the break,
(b) burasagari (hanging the trailing punct past the right text boundary),
(c) kinsoku oidashi difference. This dumps the FULL L1 char-by-char (codepoint,
X, advance) + a hang check for comma & open_kak under jc=both vs jc=left.

Geometry: left margin 1418tw=70.9pt; right text boundary = 595.3 - 70.9 = 524.4pt.
A char HANGS if its right edge (X+adv) exceeds 524.4pt.
"""
import os
import win32com.client as w32

OUT = os.path.abspath('tools/golden-test/repros/breakflip_jc')
WD_VPOS = 6
WD_HPOS = 5
RIGHT_BOUNDARY = 595.3 - 70.9  # 524.4pt text boundary
NAT = 12.0

word = w32.DispatchEx('Word.Application')
word.Visible = False
try:
    for k in ['comma', 'open_kak', 'close_paren', 'period']:
        for jc in ['both', 'left']:
            path = os.path.abspath(os.path.join(OUT, 'bf_%s_%s.docx' % (k, jc)))
            doc = word.Documents.Open(path, ReadOnly=True)
            rng = doc.Paragraphs(1).Range
            text = rng.Text
            start = rng.Start
            y0 = doc.Range(start, start).Information(WD_VPOS)
            row = []  # (codepoint_hex, char, X)
            for i in range(len(text)):
                ch = text[i]
                if ch in ('\r', '\n', '\x07'):
                    continue
                x = doc.Range(start + i, start + i).Information(WD_HPOS)
                if doc.Range(start + i, start + i).Information(WD_VPOS) > y0 + 2:
                    break
                row.append((ch, round(x, 2)))
            # advances
            advs = []
            for j in range(len(row) - 1):
                advs.append(round(row[j + 1][1] - row[j][1], 2))
            n = len(row)
            last_x = row[-1][1] if row else 0
            # estimate last char right edge: use natural 12.0 (or measured prev advance pattern)
            last_right_est = last_x + NAT
            hang = last_right_est - RIGHT_BOUNDARY
            print('--- %s jc=%s : L1=%d chars, last_x=%.2f, last_right(est +12)=%.2f, hang=%.2f%s'
                  % (k, jc, n, last_x, last_right_est, hang, '  <== HANGS' if hang > 1 else ''))
            # show last 6 advances (the break region)
            tail = advs[-6:] if len(advs) >= 6 else advs
            print('     last advances:', tail)
            # count how many advances < 11.5 (compressed)
            ncomp = sum(1 for a in advs if a < 11.5)
            print('     compressed (<11.5): %d / %d ; min adv=%.2f' % (ncomp, len(advs), min(advs) if advs else 0))
            doc.Close(False)
finally:
    word.Quit()
