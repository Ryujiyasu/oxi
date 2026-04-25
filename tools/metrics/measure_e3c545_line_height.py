"""
COM measurement: e3c545 line height analysis
- Meiryo 10.5pt (sz=21), no docGrid, no w:spacing
- pgMar top=1134tw=56.7pt
- Text expected: default "auto" line spacing with Meiryo font natural metrics
- Measure Y positions of P1..P30 to get actual line height
- Compare against Oxi's layout output
"""
import win32com.client
import json
import os

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    docx_path = os.path.abspath(
        "tools/golden-test/documents/docx/e3c545fac7a7_LOD_Handbook.docx"
    )

    doc = word.Documents.Open(docx_path, ReadOnly=True)
    results = {"paragraphs": []}

    try:
        total_paras = doc.Paragraphs.Count
        print(f"Total paragraphs: {total_paras}")
        print(f"Total pages: {doc.ComputeStatistics(2)}")

        print("\n=== First 30 paragraph Y positions ===")
        prev_y = None
        prev_page = None
        for i in range(1, min(31, total_paras + 1)):
            para = doc.Paragraphs(i)
            r = para.Range
            y = r.Information(6)
            x = r.Information(5)
            page = r.Information(3)
            text = r.Text[:40].replace('\r', '\\r').replace('\n', '\\n')

            fmt = para.Format
            ls = fmt.LineSpacing
            lsr = fmt.LineSpacingRule
            sb = fmt.SpaceBefore
            sa = fmt.SpaceAfter

            gap_str = ""
            if prev_y is not None and page == prev_page:
                gap_str = f" gap={y-prev_y:.2f}"

            print(f"  P{i:3d} p{page} y={y:7.2f} x={x:7.2f}{gap_str}  ls={ls:.1f}({lsr}) sb={sb:.1f} sa={sa:.1f}  \"{text}\"")

            results["paragraphs"].append({
                "index": i, "page": page, "y": y, "x": x,
                "line_spacing": ls, "line_spacing_rule": lsr,
                "space_before": sb, "space_after": sa,
                "text": text,
            })
            prev_y = y
            prev_page = page

        # Measure line-to-line within a wrapped paragraph
        # P3 is body para that wraps to 5 lines per visual
        print("\n=== Wrapped-paragraph internal line Y (P3) ===")
        p3 = doc.Paragraphs(3)
        r = p3.Range
        print(f"P3 text length: {len(r.Text)}")
        # Sample Y at every 10 chars
        text = r.Text
        for i in range(0, min(len(text), 400), 10):
            sub = doc.Range(r.Start + i, r.Start + i + 1)
            y = sub.Information(6)
            ch = text[i].replace('\r', '\\r').replace('\n', '\\n')
            print(f"  offset {i:3d} ch='{ch}' y={y:.2f}")

    finally:
        doc.Close(False)
        word.Quit()

    out_path = "pipeline_data/e3c545_line_height.json"
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved to {out_path}")

if __name__ == "__main__":
    main()
