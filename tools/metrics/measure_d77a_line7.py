import win32com.client, os
docx_path = os.path.abspath("tools/golden-test/documents/docx/d77a58485f16_20240705_resources_data_outline_08.docx")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(docx_path)
try:
    p = doc.Paragraphs(10)
    rng = p.Range
    text = rng.Text
    prev_line = -1
    line_starts = []
    for ci in range(min(len(text), 350)):
        c = rng.Characters(ci + 1)
        ln = c.Information(10)
        if ln != prev_line:
            line_starts.append((ln, ci))
            prev_line = ln
    # LINE 7 = index 6
    ln, sci = line_starts[6]
    eci = line_starts[7][1] if len(line_starts) > 7 else len(text)
    lt = text[sci:eci]
    print(f"LINE 7 (L{ln}): {eci-sci} chars")
    # Show each char with position
    for ci in range(sci, min(eci, sci+5)):
        c = rng.Characters(ci+1)
        x = c.Information(5)
        ch = text[ci]
        print(f"  C{ci}: x={x:.1f} '{ch}' U+{ord(ch):04X}")
    print("  ...")
    for ci in range(max(sci, eci-5), eci):
        c = rng.Characters(ci+1)
        x = c.Information(5)
        ch = text[ci]
        print(f"  C{ci}: x={x:.1f} '{ch}' U+{ord(ch):04X}")
    # Also check char after line end
    if eci < len(text):
        c = rng.Characters(eci+1)
        x = c.Information(5)
        ch = text[eci]
        print(f"  NEXT: C{eci}: x={x:.1f} '{ch}' U+{ord(ch):04X}")
    # Character spacing check
    print(f"\n  Full text: {lt}")
finally:
    doc.Close(False)
    word.Quit()
