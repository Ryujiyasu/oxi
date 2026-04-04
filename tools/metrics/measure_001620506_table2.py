"""Measure Table 2 (相対的必要記載事項BOX) cell paragraph positions in 001620506.docx.

Table 2 = body element index 33, 1 row, 1 cell, 16 paragraphs.
Oxi renders 17 lines x 18pt = 306pt, Word appears ~108pt.
Need to measure exact Y positions and line heights.
"""
import win32com.client
import os, json, time

DOCX = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..",
    "tools", "golden-test", "documents", "docx", "3a4f9fbe1a83_001620506.docx"))

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    time.sleep(2)

    results = {}

    # Get page setup
    ps = doc.PageSetup
    results["page_setup"] = {
        "page_width": ps.PageWidth,
        "page_height": ps.PageHeight,
        "margin_top": ps.TopMargin,
        "margin_bottom": ps.BottomMargin,
        "margin_left": ps.LeftMargin,
        "margin_right": ps.RightMargin,
    }
    print(f"Page: {ps.PageWidth}x{ps.PageHeight}, margins T={ps.TopMargin} B={ps.BottomMargin} L={ps.LeftMargin} R={ps.RightMargin}")

    # Find tables
    tables = doc.Tables
    print(f"Total tables: {tables.Count}")

    # Table 1 (first table on page 2)
    for ti in [1, 2]:
        tbl = tables(ti)
        cell = tbl.Cell(1, 1)
        r = cell.Range
        # Cell dimensions
        print(f"\n=== Table {ti} ===")
        print(f"  Cell width: {cell.Width}")
        print(f"  Cell height: {cell.Height}")

        # Measure each paragraph in cell
        cell_paras = r.Paragraphs
        print(f"  Cell paragraphs: {cell_paras.Count}")

        tbl_data = {
            "cell_width": cell.Width,
            "cell_height": cell.Height,
            "paragraphs": []
        }

        for pi in range(1, cell_paras.Count + 1):
            para = cell_paras(pi)
            pr = para.Range
            # Y position
            y = pr.Information(6)  # wdVerticalPositionRelativeToPage
            x = pr.Information(5)  # wdHorizontalPositionRelativeToPage
            # Line spacing
            fmt = para.Format
            ls = fmt.LineSpacing
            ls_rule = fmt.LineSpacingRule
            sb = fmt.SpaceBefore
            sa = fmt.SpaceAfter
            # Font size
            font = pr.Font
            sz = font.Size
            # Text
            text = pr.Text[:40].replace('\r', '').replace('\x07', '')
            # Page
            page = pr.Information(3)  # wdActiveEndPageNumber

            para_data = {
                "index": pi,
                "page": page,
                "x": x,
                "y": y,
                "font_size": sz,
                "line_spacing": ls,
                "line_spacing_rule": ls_rule,
                "space_before": sb,
                "space_after": sa,
                "text": text,
            }
            tbl_data["paragraphs"].append(para_data)
            print(f"  P{pi:2d}: page={page} y={y:6.1f} x={x:5.1f} sz={sz} ls={ls} rule={ls_rule} sb={sb} sa={sa} text={text[:30]}")

        results[f"table_{ti}"] = tbl_data

    # Also measure body paragraphs around table area (P29-P32 body indices)
    print("\n=== Body paragraphs around tables ===")
    body_paras = doc.Paragraphs
    print(f"Total body paragraphs: {body_paras.Count}")

    # Find paragraphs on page 2
    results["body_p2"] = []
    for pi in range(1, min(body_paras.Count + 1, 100)):
        para = body_paras(pi)
        pr = para.Range
        page = pr.Information(3)
        if page == 2:
            y = pr.Information(6)
            text = pr.Text[:40].replace('\r', '').replace('\x07', '')
            fmt = para.Format
            ls = fmt.LineSpacing
            sb = fmt.SpaceBefore
            sa = fmt.SpaceAfter
            sz = pr.Font.Size
            results["body_p2"].append({
                "body_index": pi,
                "y": y,
                "font_size": sz,
                "line_spacing": ls,
                "space_before": sb,
                "space_after": sa,
                "text": text[:30],
            })
            print(f"  Body P{pi:3d}: y={y:6.1f} sz={sz} ls={ls} sb={sb} sa={sa} text={text[:30]}")
        elif page > 2:
            break

    doc.Close(False)
    word.Quit()

    out_path = os.path.join(os.path.dirname(__file__), "..", "..",
        "pipeline_data", "com_measurements", "001620506_table2.json")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved to {out_path}")

if __name__ == "__main__":
    main()
