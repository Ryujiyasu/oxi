"""Reproduce DML extract line counting for P196."""
import win32com.client
import os, time

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

path = os.path.abspath("tools/golden-test/documents/docx/459f05f1e877_kyodokenkyuyoushiki01.docx")
doc = word.Documents.Open(path, ReadOnly=True)
time.sleep(1)

para = doc.Paragraphs(196)
rng = para.Range
chars = rng.Characters
n_chars = chars.Count

prev_y = None
lines = []
line_start_x = None
line_char_count = 0
errors = 0

for ci in range(1, n_chars + 1):
    try:
        c = chars(ci)
        ch = c.Text
        if ch == '\r' or ch == '\x07':
            continue
        cy = c.Information(6)
        cx = c.Information(5)

        if prev_y is None or abs(cy - prev_y) > 0.5:
            if prev_y is not None:
                lines.append({
                    "y": round(prev_y, 2),
                    "x": round(line_start_x, 2),
                    "chars": line_char_count,
                })
            line_start_x = cx
            line_char_count = 0
            prev_y = cy

        line_char_count += 1
    except Exception as e:
        errors += 1
        continue

if prev_y is not None and line_char_count > 0:
    lines.append({
        "y": round(prev_y, 2),
        "x": round(line_start_x, 2),
        "chars": line_char_count,
    })

print(f"Lines detected: {len(lines)}")
for l in lines:
    print(f"  y={l['y']} x={l['x']} chars={l['chars']}")
print(f"Errors: {errors}")

doc.Close(SaveChanges=False)
word.Quit()
