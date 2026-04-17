"""COM-measure per-char x in Word d77a para 28 "・利用規約名..."

Oxi wraps this to 2 lines (37+2 chars), Word fits 1 line (40 chars).
Hypothesis check: does Word compress yakumono chars to fit?
If YES: locate which chars compress and by how much.
If NO: the difference is elsewhere (available width, font metrics).
"""
import os, sys, time, json
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
    # Iterate body paragraphs, find the one with "・利用規約名"
    target = "・利用規約名"
    found = None
    for i in range(1, doc.Paragraphs.Count + 1):
        try:
            p = doc.Paragraphs(i)
            txt = p.Range.Text[:20]
            if target in txt:
                found = (i, p)
                break
        except Exception:
            pass
    if not found:
        print("Not found", file=sys.stderr); sys.exit(1)
    idx, p = found
    print(f"Found at para {idx}", file=sys.stderr)
    pr = p.Range
    n = pr.Characters.Count
    print(f"n chars: {n}", file=sys.stderr)
    print(f"{'#':>3} {'char':>4} {'x':>8} {'y':>8} {'adv':>6}")
    prev_x = None
    prev_y = None
    for i in range(1, n + 1):
        try:
            ch = pr.Characters(i)
            x = ch.Information(5)  # horizontal position rel to page
            y = ch.Information(6)
            c = ch.Text
            adv = f"{x-prev_x:.2f}" if prev_x is not None and y == prev_y else "wrap"
            print(f"{i:3} {c:>4} {x:8.2f} {y:8.2f} {adv:>6}")
            prev_x = x; prev_y = y
        except Exception as e:
            print(f"{i}: err {e}")
    doc.Close(False)
finally:
    word.Quit()
