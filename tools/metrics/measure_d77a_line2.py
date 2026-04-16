"""Measure Line 2 char positions."""
import win32com.client, os
docx_path = os.path.abspath("tools/golden-test/documents/docx/d77a58485f16_20240705_resources_data_outline_08.docx")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(docx_path)
try:
    p = doc.Paragraphs(10)
    rng = p.Range
    text = rng.Text
    # Line 2 starts at char 40 (after 39 chars on line 1)
    # Check chars 38-42 to see the break point
    for ci in range(38, min(82, len(text))):
        c = rng.Characters(ci + 1)
        x = c.Information(5)
        y = c.Information(6)
        ln = c.Information(10)
        ch = text[ci]
        if ci < 42 or ci > 75:
            print(f"  C{ci}: x={x:.1f} y={y:.1f} L{ln} '{ch}' U+{ord(ch):04X}")
    # Total width of line 2
    c_start = rng.Characters(40)
    c_end = rng.Characters(77)
    x_start = c_start.Information(5)
    x_end = c_end.Information(5)
    print(f"\nLine 2: x_start={x_start:.1f} x_end={x_end:.1f} width={x_end-x_start:.1f}")
finally:
    doc.Close(False)
    word.Quit()
