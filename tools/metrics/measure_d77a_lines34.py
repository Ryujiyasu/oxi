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
    # Lines 3,4,7 (index 2,3,6)
    for idx in [2, 3, 6]:
        ln, sci = line_starts[idx]
        eci = line_starts[idx+1][1] if idx+1 < len(line_starts) else len(text)
        lt = text[sci:eci]
        hw = [ch for ch in lt if ord(ch) < 0x2000 and ch.strip()]
        # Compressible punct (same as kinsoku::is_cjk_compressible)
        comp = [ch for ch in lt if ch in '、。，．「」（）『』【】〔〕〘〖〙〗｛｝［］']
        # Check char gaps for compression evidence
        gaps = []
        prev_x = None
        for ci2 in range(sci, min(eci, sci+50)):
            c2 = rng.Characters(ci2+1)
            x2 = c2.Information(5)
            if prev_x is not None:
                gaps.append(x2 - prev_x)
            prev_x = x2
        non12 = [(i, g) for i, g in enumerate(gaps) if abs(g - 12.0) > 0.5]
        print(f"LINE {idx+1} (L{ln}): {eci-sci}ch hw={len(hw)} comp={len(comp)} non12gaps={non12}")
finally:
    doc.Close(False)
    word.Quit()
