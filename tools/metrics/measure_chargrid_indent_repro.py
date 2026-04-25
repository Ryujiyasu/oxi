"""COM-measure each L2*_*.docx repro: where does Word break line 1?

Output: per-repro JSON with {indent shape, body text len, line 1 break char,
chars on line 1, line 1 right-edge X}. Saved to
tools/metrics/chargrid_indent_repro/word_com_measurements.json.
"""
import win32com.client
import os, json
from pathlib import Path

word = win32com.client.gencache.EnsureDispatch('Word.Application')
word.Visible = False
repro_dir = r'C:\Users\ryuji\oxi-main\tools\metrics\chargrid_indent_repro'

names = sorted(f for f in os.listdir(repro_dir) if f.startswith('L2') and f.endswith('.docx'))
results = {}

for name in names:
    path = os.path.join(repro_dir, name)
    doc = word.Documents.Open(path, ReadOnly=True)
    try:
        # The first body paragraph is our test paragraph
        para = doc.Paragraphs(1)
        r = para.Range
        n_chars = r.Characters.Count
        # Text excluding paragraph mark
        text = r.Text.replace('\r', '').replace('\n', '')
        # Per-char line number (Information(10) = wdFirstCharacterLineNumber)
        # Find first char where line number != 1
        line1_last_char = None
        line2_first_char = None
        line1_x_left = None
        line1_x_right = None
        x_first = None
        for j in range(1, min(n_chars, 60) + 1):
            try:
                ch = r.Characters(j)
                line = ch.Information(10)
                x = ch.Information(1) / 20.0  # twips → pt
                if x_first is None:
                    x_first = x
                    line1_x_left = x
                if line == 1:
                    line1_last_char = j
                    line1_x_right = x
                else:
                    if line2_first_char is None:
                        line2_first_char = j
                    break
            except Exception:
                continue
        # Total line count
        try:
            line_count = max(r.Characters(j).Information(10) for j in range(1, min(n_chars, 60) + 1))
        except Exception:
            line_count = None
        results[name] = {
            'text_len': n_chars,
            'line1_last_char': line1_last_char,
            'line1_chars': line1_last_char,  # alias
            'line2_first_char': line2_first_char,
            'line1_x_left_pt': round(line1_x_left, 2) if line1_x_left else None,
            'line1_x_right_pt': round(line1_x_right, 2) if line1_x_right else None,
            'line_count': line_count,
        }
        n1 = line1_last_char or "?"
        x_left = round(line1_x_left, 2) if line1_x_left else "?"
        x_right = round(line1_x_right, 2) if line1_x_right else "?"
        print(f"{name:20s}  line1_chars={n1:>3}  x={x_left}→{x_right}  total_lines={line_count}")
    finally:
        doc.Close(SaveChanges=False)

word.Quit()

out = os.path.join(repro_dir, 'word_com_measurements.json')
with open(out, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nSaved {out}")
