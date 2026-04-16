"""Check LINE 2 chars 39-78."""
import win32com.client, os
docx_path = os.path.abspath("tools/golden-test/documents/docx/d77a58485f16_20240705_resources_data_outline_08.docx")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(docx_path)
try:
    p = doc.Paragraphs(10)
    rng = p.Range
    text = rng.Text
    print(f"Line 2 text (C39-C76): {text[39:77]}")
    print(f"Chars: {len(text[39:77])}")
    # Halfwidth chars
    hw = [(i+39, ch) for i, ch in enumerate(text[39:77]) if ord(ch) < 0x2000 and ch.strip()]
    print(f"Halfwidth: {hw}")
    # Check char 76-78
    for ci in [75, 76, 77, 78]:
        c = rng.Characters(ci + 1)
        x = c.Information(5)
        ln = c.Information(10)
        ch = text[ci]
        print(f"C{ci}: L{ln} x={x:.1f} '{ch}' U+{ord(ch):04X}")
finally:
    doc.Close(False)
    word.Quit()
