"""Measure Line 4 (L13) of para 10."""
import win32com.client, os
docx_path = os.path.abspath("tools/golden-test/documents/docx/d77a58485f16_20240705_resources_data_outline_08.docx")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(docx_path)
try:
    p = doc.Paragraphs(10)
    rng = p.Range
    text = rng.Text
    # Find line 13 (4th line of paragraph)
    prev_line = -1
    line_starts = []
    for ci in range(min(len(text), 200)):
        c = rng.Characters(ci + 1)
        ln = c.Information(10)
        if ln != prev_line:
            line_starts.append((ln, ci))
            prev_line = ln
    
    # Line 4 = line_starts[3]
    if len(line_starts) > 3:
        ln, start_ci = line_starts[3]
        end_ci = line_starts[4][1] if len(line_starts) > 4 else len(text)
        line_text = text[start_ci:end_ci]
        print(f"Line 4 (L{ln}): {end_ci-start_ci} chars")
        print(f"  Text: {line_text[:60]}")
        
        # Check each char
        for ci in range(start_ci, min(end_ci, start_ci+5)):
            c = rng.Characters(ci + 1)
            x = c.Information(5)
            ch = text[ci]
            print(f"  C{ci}: x={x:.1f} '{ch}' U+{ord(ch):04X}")
        print("  ...")
        for ci in range(max(start_ci, end_ci-5), end_ci):
            c = rng.Characters(ci + 1)
            x = c.Information(5)
            ch = text[ci]
            print(f"  C{ci}: x={x:.1f} '{ch}' U+{ord(ch):04X}")
        
        # Check if halfwidth chars exist
        hw = [ch for ch in line_text if ord(ch) < 0x2000 and ch.strip()]
        print(f"\n  Halfwidth chars: {hw}")
finally:
    doc.Close(False)
    word.Quit()
