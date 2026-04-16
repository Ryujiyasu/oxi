"""Measure first character X position."""
import win32com.client, os
docx_path = os.path.abspath("tools/golden-test/documents/docx/d77a58485f16_20240705_resources_data_outline_08.docx")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(docx_path)
try:
    p = doc.Paragraphs(10)
    rng = p.Range
    # First 5 chars
    for ci in range(1, 6):
        c = rng.Characters(ci)
        x = c.Information(5)
        y = c.Information(6)
        ch = c.Text
        print(f"Char {ci}: x={x:.1f} y={y:.1f} '{ch}' U+{ord(ch):04X}")
    # Also check char 39 (line break point)
    for ci in [38, 39, 40, 41]:
        c = rng.Characters(ci)
        x = c.Information(5)
        y = c.Information(6)
        ln = c.Information(10)
        ch = c.Text
        print(f"Char {ci}: x={x:.1f} y={y:.1f} line={ln} '{ch}'")
finally:
    doc.Close(False)
    word.Quit()
