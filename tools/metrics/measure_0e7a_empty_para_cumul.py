"""Measure cumulative line positions around empty paragraphs in 0e7a.
Goal: determine if Word uses doc_default or paragraph font for cumul round base."""
import win32com.client
import os

docx_path = os.path.abspath(r"tools\golden-test\documents\docx\0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

try:
    doc = word.Documents.Open(docx_path, ReadOnly=True)

    # Page 1: measure all paragraphs to see cumul pattern
    print("=== Page 1 paragraphs ===")
    prev_y = None
    for i in range(1, 40):
        para = doc.Paragraphs(i)
        rng = para.Range
        page = rng.Information(3)
        if page > 1:
            break
        y = rng.Information(6)
        ls = para.Format.LineSpacing
        font = rng.Font.Name
        fs = rng.Font.Size
        text = rng.Text[:30].replace('\r', '\\r').replace('\n', '\\n')
        gap = f"gap={y - prev_y:.1f}" if prev_y is not None else ""
        is_empty = len(rng.Text.strip()) <= 1
        print(f"P{i:3d} y={y:7.1f} ls={ls:5.1f} fs={fs:4.1f} {gap:>10} {'EMPTY' if is_empty else ''} [{font}] {text[:20]}")
        prev_y = y

    # Page 6: paragraphs around the boundary
    print("\n=== Page 6 paragraphs (around empty paras) ===")
    prev_y = None
    for i in range(205, 235):
        para = doc.Paragraphs(i)
        rng = para.Range
        page = rng.Information(3)
        if page < 6:
            continue
        if page > 7:
            break
        y = rng.Information(6)
        ls = para.Format.LineSpacing
        font = rng.Font.Name
        fs = rng.Font.Size
        text = rng.Text[:30].replace('\r', '\\r').replace('\n', '\\n')
        gap = f"gap={y - prev_y:.1f}" if prev_y is not None and page == 6 else ""
        is_empty = len(rng.Text.strip()) <= 1
        print(f"P{i:3d} p{page} y={y:7.1f} ls={ls:5.1f} fs={fs:4.1f} {gap:>10} {'EMPTY' if is_empty else ''} [{font}] {text[:20]}")
        prev_y = y if page == 6 else None

    doc.Close(False)
finally:
    word.Quit()
