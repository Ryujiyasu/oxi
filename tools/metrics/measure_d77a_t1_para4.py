"""Dense char-by-char Y measurement for d77a Table 1 paragraph 4.

Prior measurement showed last_y=216.5 (same as para1 first_y) — likely a
COM Information(6) artifact on trailing whitespace. Re-measure with
strict char iteration, reporting each unique Y.
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
    tbl = doc.Tables(1)
    cell = tbl.Cell(1, 1)
    paras = list(cell.Range.Paragraphs)
    print(f'Cell has {len(paras)} paragraphs')
    for pi, p in enumerate(paras, 1):
        pr = p.Range
        n = pr.Characters.Count
        # Sample EVERY char
        ys = []
        for i in range(1, n + 1):
            try:
                ch = pr.Characters(i)
                y = ch.Information(6)
                ys.append((i, round(y, 2), ch.Text[:1] if ch.Text else ''))
            except Exception:
                pass
        uniq_ys = sorted(set(y[1] for y in ys))
        print(f'\n=== para{pi} n={n} unique_ys={len(uniq_ys)} ===')
        print('  y values:', uniq_ys)
        # Show the text position break points (first char at each unique y)
        for target_y in uniq_ys:
            for (i, y, c) in ys:
                if y == target_y:
                    # collect all chars on this line
                    line_chars = [c2 for (i2, y2, c2) in ys if y2 == target_y]
                    line_text = ''.join(line_chars)
                    print(f'  y={y} chars={len(line_chars)} text={line_text[:50]}')
                    break
    doc.Close(False)
finally:
    word.Quit()
