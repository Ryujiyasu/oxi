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
    # Measure each char x position
    prev_x = None
    for ci in range(sci, eci):
        c = rng.Characters(ci + 1)
        x = c.Information(5)
        ch = text[ci]
        gap = x - prev_x if prev_x is not None else 0
        if ci < sci + 5 or ci > eci - 5:
            print(f"C{ci-sci}: x={x:.1f} gap={gap:.1f} '{ch}'")
        prev_x = x
    # Total width
    first_x = rng.Characters(sci + 1).Information(5)
    last_x = rng.Characters(eci).Information(5)
    print(f"\nTotal: first={first_x:.1f} last={last_x:.1f} span={last_x-first_x:.1f} chars={eci-sci}")
    print(f"Avg char width: {(last_x-first_x)/(eci-sci-1):.3f}pt")
finally:
    doc.Close(False)
    word.Quit()
