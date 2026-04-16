import win32com.client, os
docx_path = os.path.abspath("tools/golden-test/documents/docx/d77a58485f16_20240705_resources_data_outline_08.docx")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(docx_path)
try:
    p = doc.Paragraphs(10)
    rng = p.Range
    text = rng.Text
    # Lines 3,4 of this para
    prev_line = -1
    line_starts = []
    for ci in range(min(len(text), 200)):
        c = rng.Characters(ci + 1)
        ln = c.Information(10)
        if ln != prev_line:
            line_starts.append((ln, ci))
            prev_line = ln
    for idx in [2, 3]:  # LINE 3, LINE 4
        ln, sci = line_starts[idx]
        eci = line_starts[idx+1][1] if idx+1 < len(line_starts) else len(text)
        lt = text[sci:eci]
        hw = [ch for ch in lt if ord(ch) < 0x2000 and ch.strip()]
        cp = [ch for ch in lt if ch in '、。「」（）『』【】〔〕']
        print(f"L{ln}: {eci-sci}chars hw={hw} compress_punct={cp}")
        print(f"  text: {lt[:50]}")
finally:
    doc.Close(False)
    word.Quit()
