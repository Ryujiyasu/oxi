"""Measure Word line break positions for d77a body paragraph via COM."""
import win32com.client
import os

docx_path = os.path.abspath("tools/golden-test/documents/docx/d77a58485f16_20240705_resources_data_outline_08.docx")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(docx_path)

try:
    for i in range(1, min(30, doc.Paragraphs.Count + 1)):
        p = doc.Paragraphs(i)
        text = p.Range.Text[:50]
        print(f"P{i}: {text[:40]}")
        if "前身" in text:
            print(f"\n=== Target para {i} ===")
            rng = p.Range
            text_full = rng.Text
            
            prev_line = -1
            line_starts = []
            for ci in range(min(len(text_full), 250)):
                char_rng = rng.Characters(ci + 1)
                line_num = char_rng.Information(10)
                if line_num != prev_line:
                    char_x = char_rng.Information(5)
                    char_y = char_rng.Information(6)
                    ch = text_full[ci]
                    line_starts.append((line_num, ci, ch, char_x, char_y))
                    prev_line = line_num
            
            for idx, (ln, ci, ch, cx, cy) in enumerate(line_starts[:8]):
                next_ci = line_starts[idx+1][1] if idx+1 < len(line_starts) else len(text_full)
                chars_on_line = next_ci - ci
                line_text = text_full[ci:next_ci]
                print(f"  Line {ln}: {chars_on_line} chars x={cx:.1f} y={cy:.1f} \"{line_text[:40]}\"")
            break
finally:
    doc.Close(False)
    word.Quit()
