"""Measure Word actual line count for db9c P11 (574ch)."""
import win32com.client, time, sys, os
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = os.path.abspath("tools/golden-test/documents/docx/db9ca18368cd_20241122_resource_open_data_01.docx")

word = win32com.client.Dispatch("Word.Application")
word.Visible = True
word.DisplayAlerts = False
doc = word.Documents.Open(DOC, ReadOnly=True)
time.sleep(1)

print(f"Total paragraphs: {doc.Paragraphs.Count}")
for pi in [4, 5, 6, 8, 11, 12, 15]:
    if pi > doc.Paragraphs.Count: break
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
    print(f"  Word lines: {len(line_ys)}")
    print(f"  Y values: {line_ys}")
    print(f"  chars/line: {chars_per_line}")

doc.Close(SaveChanges=False)
word.Quit()
