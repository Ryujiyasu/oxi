"""Measure Word actual line count for c7b9 P9 (217ch)."""
import win32com.client, time, sys, os
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = os.path.abspath("tools/golden-test/documents/docx/c7b923e5c616_20240705_resources_data_outline_06.docx")

word = win32com.client.Dispatch("Word.Application")
word.Visible = True
word.DisplayAlerts = False
doc = word.Documents.Open(DOC, ReadOnly=True)
time.sleep(1)

for pi in [8, 9, 13]:
    p = doc.Paragraphs(pi)
    rng = p.Range
    txt = rng.Text
    print(f"\n=== P{pi}: {len(txt)} chars ===")
    print(f"  font: {rng.Font.Name} sz={rng.Font.Size}")
    print(f"  align: {p.Alignment}")
    chars = rng.Characters
    n = chars.Count
    line_ys = []
    chars_per_line = []
    cur = 0
    prev_y = None
    for ci in range(1, n + 1):
        try:
            c = chars(ci)
            ch = c.Text
            if ch in ('\r', '\x07', '\x0b'):
                continue
            cy = c.Information(6)
            if prev_y is None or abs(cy - prev_y) > 0.5:
                if prev_y is not None:
                    line_ys.append(prev_y)
                    chars_per_line.append(cur)
                    cur = 0
                prev_y = cy
            cur += 1
        except:
            break
    if cur > 0:
        line_ys.append(prev_y)
        chars_per_line.append(cur)
    print(f"  Word lines: {len(line_ys)}, Y={line_ys}")
    print(f"  chars/line: {chars_per_line}")

doc.Close(SaveChanges=False)
word.Quit()
