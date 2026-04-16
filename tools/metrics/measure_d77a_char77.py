"""Check char 77 (39th on line 2) position."""
import win32com.client, os
docx_path = os.path.abspath("tools/golden-test/documents/docx/d77a58485f16_20240705_resources_data_outline_08.docx")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(docx_path)
try:
    p = doc.Paragraphs(10)
    rng = p.Range
    text = rng.Text
    # Char 76 = last on L11, Char 77 = first on L12
    for ci in [75, 76, 77]:
        c = rng.Characters(ci + 1)
        x = c.Information(5)
        ln = c.Information(10)
        ch = text[ci]
        cw = c.Characters(1).Information(14) if hasattr(c, 'Information') else 0  # wdFrameWidth?
        print(f"C{ci}: L{ln} x={x:.1f} '{ch}' U+{ord(ch):04X}")
    # Char 76 is ー (U+30FC). What if ー has different width?
    print(f"\nLine 2 text (C39-C76): {text[39:77]}")
    print(f"Line 2 char count: {77-39}")
finally:
    doc.Close(False)
    word.Quit()
