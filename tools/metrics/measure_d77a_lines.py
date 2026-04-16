"""Measure Word line break positions for d77a paragraphs 10-12."""
import win32com.client, os
docx_path = os.path.abspath("tools/golden-test/documents/docx/d77a58485f16_20240705_resources_data_outline_08.docx")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(docx_path)
try:
    # Check paragraphs 10-15 (body text)
    for pi in range(10, 16):
        p = doc.Paragraphs(pi)
        rng = p.Range
        text = rng.Text
        if len(text) < 10: continue
        
        prev_line = -1
        line_starts = []
        for ci in range(min(len(text), 300)):
            c = rng.Characters(ci + 1)
            ln = c.Information(10)
            if ln != prev_line:
                line_starts.append((ln, ci))
                prev_line = ln
        
        print(f"P{pi} ({len(text)} chars, {len(line_starts)} lines):")
        for idx, (ln, ci) in enumerate(line_starts[:10]):
            next_ci = line_starts[idx+1][1] if idx+1 < len(line_starts) else len(text)
            chars = next_ci - ci
            print(f"  L{ln}: {chars} chars")
finally:
    doc.Close(False)
    word.Quit()
