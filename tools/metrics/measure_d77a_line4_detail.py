"""Measure Line 4 (L13) of para 10 in detail, show all widths."""
import win32com.client, os
docx_path = os.path.abspath("tools/golden-test/documents/docx/d77a58485f16_20240705_resources_data_outline_08.docx")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(docx_path)
try:
    p = doc.Paragraphs(10)
    rng = p.Range
    text = rng.Text

    # Find line starts
    prev_line = -1
    line_starts = []
    for ci in range(min(len(text), 300)):
        c = rng.Characters(ci + 1)
        ln = c.Information(10)
        if ln != prev_line:
            line_starts.append((ln, ci))
            prev_line = ln

    # Dump LINE 4 in detail with char widths
    if len(line_starts) > 4:
        ln4, start4 = line_starts[3]
        ln5, start5 = line_starts[4]
        print(f"Line 4 (L{ln4}): chars {start4}..{start5-1} ({start5-start4} chars)")
        prev_x = None
        for ci in range(start4, start5 + 1):  # include first char of L5
            c = rng.Characters(ci + 1)
            try:
                x = c.Information(5)
                y = c.Information(6)
                ch = text[ci]
                advance = (x - prev_x) if prev_x is not None else None
                marker = "  "
                if ci == start5:
                    marker = "L5"
                print(f"  {marker} C{ci:3d}: x={x:6.1f} y={y:6.1f} '{ch}' U+{ord(ch):04X}" +
                      (f"  adv={advance:.1f}" if advance is not None else ""))
                prev_x = x
            except Exception as e:
                print(f"  C{ci}: ERR {e}")
                break
finally:
    doc.Close(False)
    word.Quit()
