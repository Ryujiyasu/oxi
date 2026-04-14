"""Measure actual Word line count for d77a P10 (366ch para)."""
import win32com.client, time, sys, os
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = os.path.abspath("tools/golden-test/documents/docx/d77a58485f16_20240705_resources_data_outline_08.docx")

word = win32com.client.Dispatch("Word.Application")
word.Visible = True
word.DisplayAlerts = False
doc = word.Documents.Open(DOC, ReadOnly=True)
time.sleep(1)

# Find P10 (the 366ch one - actually paragraph index 10 in 1-based)
for pi in [10]:
    p = doc.Paragraphs(pi)
    rng = p.Range
    txt = rng.Text
    print(f"P{pi}: {len(txt)} chars, text: {txt[:80]!r}")
    print(f"  font: {rng.Font.Name} sz={rng.Font.Size}")
    print(f"  align: {p.Alignment}")
    # Measure all character Y positions, count distinct line Ys
    chars = rng.Characters
    n = chars.Count
    print(f"  n_chars: {n}")
    line_ys = []
    prev_y = None
    chars_per_line = []
    cur_chars = 0
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
                    chars_per_line.append(cur_chars)
                    cur_chars = 0
                prev_y = cy
            cur_chars += 1
        except:
            break
    if cur_chars > 0:
        line_ys.append(prev_y)
        chars_per_line.append(cur_chars)
    print(f"  Word actual lines: {len(line_ys)}")
    print(f"  Y values: {line_ys}")
    print(f"  chars/line: {chars_per_line}")
    if len(line_ys) >= 2:
        print(f"  line spacing: {line_ys[1]-line_ys[0]:.2f}pt")

doc.Close(SaveChanges=False)
word.Quit()
