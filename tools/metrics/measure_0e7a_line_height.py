"""
COM measurement: 0e7a document line height analysis
- MS Mincho 10.5pt, no docGrid (LayoutMode=0)
- Measure Y positions of consecutive paragraphs to get actual line height
- Also measure bordered paragraph spacing
"""
import win32com.client
import json
import os
import sys

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    docx_path = os.path.abspath(
        "tools/golden-test/documents/docx/"
        "0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx"
    )

    doc = word.Documents.Open(docx_path, ReadOnly=True)
    results = {}

    try:
        total_paras = doc.Paragraphs.Count
        print(f"Total paragraphs: {total_paras}")
        print(f"Total pages: {doc.ComputeStatistics(2)}")  # wdStatisticPages

        # Measure Y positions of first 40 paragraphs
        print("\n=== Paragraph Y positions (Information(6)) ===")
        prev_y = None
        for i in range(1, min(41, total_paras + 1)):
            para = doc.Paragraphs(i)
            r = para.Range
            y = r.Information(6)  # wdVerticalPositionRelativeToPage
            x = r.Information(5)  # wdHorizontalPositionRelativeToPage
            page = r.Information(3)  # wdActiveEndPageNumber
            text = r.Text[:40].replace('\r', '\\r').replace('\n', '\\n')

            gap = f" gap={y - prev_y:.2f}pt" if prev_y is not None and page == results.get(f'p{i-1}_page') else ""
            print(f"  P{i:3d} page={page} y={y:7.2f}pt x={x:7.2f}pt{gap}  \"{text}\"")

            results[f'p{i}_y'] = y
            results[f'p{i}_x'] = x
            results[f'p{i}_page'] = page
            prev_y = y

        # Measure line spacing setting
        print("\n=== Line spacing settings ===")
        for i in [1, 2, 3, 10, 20]:
            if i <= total_paras:
                para = doc.Paragraphs(i)
                fmt = para.Format
                print(f"  P{i}: LineSpacing={fmt.LineSpacing:.2f}pt "
                      f"LineSpacingRule={fmt.LineSpacingRule} "
                      f"SpaceBefore={fmt.SpaceBefore:.2f}pt "
                      f"SpaceAfter={fmt.SpaceAfter:.2f}pt")

        # Measure bordered section in detail (paragraphs around "総則")
        # Find "総則" or "第１条"
        print("\n=== Bordered section Y gaps (page 2+) ===")
        prev_y = None
        prev_page = None
        for i in range(1, min(total_paras + 1, 100)):
            para = doc.Paragraphs(i)
            r = para.Range
            page = r.Information(3)
            if page < 2:
                continue
            y = r.Information(6)
            text = r.Text[:50].replace('\r', '\\r').replace('\n', '\\n')

            gap = ""
            if prev_y is not None and page == prev_page:
                gap = f" gap={y - prev_y:.2f}pt"

            print(f"  P{i:3d} p{page} y={y:7.2f}pt{gap}  \"{text}\"")
            prev_y = y
            prev_page = page

            if page > 3:
                break

        # Measure font-specific line height
        print("\n=== Font line height test ===")
        # Check first body paragraph's actual rendered height
        for i in [4, 5, 6, 7, 8]:
            if i <= total_paras and i + 1 <= total_paras:
                p1 = doc.Paragraphs(i)
                p2 = doc.Paragraphs(i + 1)
                y1 = p1.Range.Information(6)
                y2 = p2.Range.Information(6)
                page1 = p1.Range.Information(3)
                page2 = p2.Range.Information(3)
                if page1 == page2:
                    text1 = p1.Range.Text[:30].replace('\r', '\\r')
                    print(f"  P{i}→P{i+1}: y1={y1:.2f} y2={y2:.2f} gap={y2-y1:.2f}pt  \"{text1}\"")

    finally:
        doc.Close(False)
        word.Quit()

    # Save results
    out_path = "tools/metrics/output/0e7a_line_height.json"
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, 'w') as f:
        json.dump(results, f, indent=2)
    print(f"\nSaved to {out_path}")

if __name__ == "__main__":
    main()
