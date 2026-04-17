"""Dense line-by-line measurement of Word's bullet paragraphs on p2/p3.

Oxi's pi=25 bullet paragraph splits:
  - p2 bottom: 2 lines (y=686, 740)
  - p3 top: 2 lines (y=74, 92)

Q: Does Word split the same paragraph across p2/p3? At which line?

Strategy: enumerate Word body paragraphs near the p2/p3 boundary and report
each line Y + page. Find the paragraph (or paragraphs) straddling the boundary.
"""
import os, sys, time
import win32com.client
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = os.path.abspath(
    r"tools\golden-test\documents\docx\d77a58485f16_20240705_resources_data_outline_08.docx"
)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False
try:
    doc = word.Documents.Open(DOC, ReadOnly=True); time.sleep(0.3)
    doc.Repaginate()
    print('Scanning paragraphs on p2 and p3 for bullet-starting text...')
    target_paras = []
    for pi, p in enumerate(doc.Paragraphs, 1):
        try:
            pg = p.Range.Information(3)
            if pg not in (2, 3):
                continue
            txt = p.Range.Text.replace('\r','').replace('\x07','')[:60]
            # Bullet paragraphs start with ・ or are full-width dash
            if txt.startswith('・') or (len(txt) > 3 and txt[0] == '\u30fb'):
                y = p.Range.Information(6)
                target_paras.append((pi, pg, y, txt))
        except Exception:
            pass
    target_paras.sort(key=lambda r: (r[1], r[2]))
    print(f'\nBullet-starting paras on p2/p3: {len(target_paras)}')
    for (pi, pg, y, txt) in target_paras:
        print(f'  para_idx={pi} p{pg} y={y:.1f} text={txt!r}')

    # Now dense line Y for each of these paragraphs
    for (pi, pg, y_first, txt) in target_paras:
        p = doc.Paragraphs(pi)
        pr = p.Range
        n = pr.Characters.Count
        if n < 5:
            continue
        ys = []
        for i in range(1, n + 1):
            try:
                ch = pr.Characters(i)
                y = ch.Information(6)
                chpg = ch.Information(3)
                ys.append((i, round(y, 1), chpg, ch.Text[:1] if ch.Text else ''))
            except Exception:
                pass
        uniq = sorted(set((yy, pp) for (_, yy, pp, _) in ys))
        print(f'\n=== para_idx={pi} n={n} text={txt[:50]!r}')
        for (yy, pp) in uniq:
            # Collect chars on this line
            line_chars = [c for (_, yyy, ppp, c) in ys if yyy == yy and ppp == pp]
            line_text = ''.join(line_chars)
            print(f'  p{pp} y={yy:.1f} chars={len(line_chars)} text={line_text[:50]}')

    doc.Close(False)
finally:
    word.Quit()
