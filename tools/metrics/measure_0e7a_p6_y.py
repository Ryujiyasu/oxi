"""Measure Y positions of all paragraphs on page 6 of 0e7a document via COM."""
import win32com.client
import json
import os

docx_path = os.path.abspath(r"tools\golden-test\documents\docx\0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

try:
    doc = word.Documents.Open(docx_path, ReadOnly=True)

    # wdVerticalPositionRelativeToPage = 6
    # wdActiveEndPageNumber = 3

    results = []
    total_paras = doc.Paragraphs.Count
    print(f"Total paragraphs: {total_paras}")

    for i in range(1, total_paras + 1):
        para = doc.Paragraphs(i)
        rng = para.Range
        page_num = rng.Information(3)  # wdActiveEndPageNumber

        if page_num < 5:
            continue
        if page_num > 7:
            break

        y_pos = rng.Information(6)  # wdVerticalPositionRelativeToPage
        text = rng.Text[:60].replace('\r', '\\r').replace('\n', '\\n')
        line_spacing = para.Format.LineSpacing
        space_before = para.Format.SpaceBefore
        space_after = para.Format.SpaceAfter

        results.append({
            "para": i,
            "page": page_num,
            "y_pt": round(y_pos, 2),
            "line_spacing": round(line_spacing, 2),
            "space_before": round(space_before, 2),
            "space_after": round(space_after, 2),
            "text": text
        })

        print(f"P{i:3d} page={page_num} y={y_pos:8.2f}pt ls={line_spacing:.2f} sb={space_before:.2f} sa={space_after:.2f} [{text[:40]}]")

    doc.Close(False)

    with open("pipeline_data/0e7a_p6_y_positions.json", "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"\nSaved {len(results)} paragraphs")

finally:
    word.Quit()
