"""
Ra: テーブル行の実レンダリング高さをCOM計測
対象: 04b88e7e0b25_index-19 (Word 5p vs Oxi 4p)

計測内容:
1. 各段落のY座標 (wdVerticalPositionRelativeToPage)
2. 各テーブル行の開始Y / 終了Y / 行高さ
3. ページ分割ポイントの特定
"""
import sys
import os
import json
import win32com.client
import pythoncom

DOCX = os.path.abspath(os.path.join(
    os.path.dirname(__file__), "..", "..",
    "tools", "golden-test", "documents", "docx",
    "04b88e7e0b25_index-19.docx"
))
OUT = os.path.abspath(os.path.join(
    os.path.dirname(__file__), "..", "..",
    "pipeline_data", "com_measurements", "index19_table_row_heights.json"
))

def main():
    pythoncom.CoInitialize()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    word.AutomationSecurity = 3

    try:
        doc = word.Documents.Open(DOCX, ReadOnly=True, AddToRecentFiles=False)
        try:
            results = {"paragraphs": [], "tables": [], "page_breaks": []}

            # 1. 全段落のY座標とページ番号
            print(f"Paragraphs: {doc.Paragraphs.Count}")
            for i in range(1, doc.Paragraphs.Count + 1):
                para = doc.Paragraphs(i)
                rng = para.Range
                try:
                    y = rng.Information(6)   # wdVerticalPositionRelativeToPage
                    page = rng.Information(3) # wdActiveEndPageNumber
                    text = rng.Text[:40].replace('\r', '').replace('\n', '')
                    results["paragraphs"].append({
                        "idx": i,
                        "page": page,
                        "y_pt": round(y, 2),
                        "text": text,
                    })
                except Exception as e:
                    results["paragraphs"].append({"idx": i, "error": str(e)})

            # 2. テーブル行の詳細計測
            print(f"Tables: {doc.Tables.Count}")
            for t_idx in range(1, doc.Tables.Count + 1):
                table = doc.Tables(t_idx)
                table_data = {"table_idx": t_idx, "rows": []}

                for r_idx in range(1, table.Rows.Count + 1):
                    row = table.Rows(r_idx)
                    try:
                        # Row height settings
                        height_rule = row.HeightRule  # 0=auto, 1=atLeast, 2=exact
                        height_val = row.Height

                        # First cell's range Y position
                        first_cell = table.Cell(r_idx, 1)
                        rng = first_cell.Range
                        y = rng.Information(6)
                        page = rng.Information(3)

                        # Cell text
                        text = rng.Text[:30].replace('\r', '').replace('\n', '').replace('\x07', '')

                        row_data = {
                            "row_idx": r_idx,
                            "page": page,
                            "y_pt": round(y, 2),
                            "height_rule": height_rule,
                            "height_val": round(height_val, 2),
                            "text": text,
                        }

                        # Try to get row bottom by checking next row or paragraph after table
                        if r_idx < table.Rows.Count:
                            next_cell = table.Cell(r_idx + 1, 1)
                            next_y = next_cell.Range.Information(6)
                            next_page = next_cell.Range.Information(3)
                            if next_page == page:
                                row_data["rendered_height"] = round(next_y - y, 2)
                            else:
                                row_data["rendered_height"] = "cross-page"

                        table_data["rows"].append(row_data)
                    except Exception as e:
                        table_data["rows"].append({"row_idx": r_idx, "error": str(e)})

                results["tables"].append(table_data)

            # 3. ページ分割ポイント特定
            prev_page = 1
            for p in results["paragraphs"]:
                if "page" in p and p["page"] != prev_page:
                    results["page_breaks"].append({
                        "from_page": prev_page,
                        "to_page": p["page"],
                        "para_idx": p["idx"],
                        "y_pt": p["y_pt"],
                        "text": p.get("text", ""),
                    })
                    prev_page = p["page"]

            # 4. ページ設定
            sec = doc.Sections(1).PageSetup
            results["page_setup"] = {
                "page_height": round(sec.PageHeight, 2),
                "page_width": round(sec.PageWidth, 2),
                "margin_top": round(sec.TopMargin, 2),
                "margin_bottom": round(sec.BottomMargin, 2),
                "margin_left": round(sec.LeftMargin, 2),
                "margin_right": round(sec.RightMargin, 2),
            }

            # docGrid
            try:
                grid = doc.Sections(1).PageSetup
                results["page_setup"]["line_pitch"] = round(grid.LinePitch, 2)
            except:
                pass

            # Save
            os.makedirs(os.path.dirname(OUT), exist_ok=True)
            with open(OUT, "w", encoding="utf-8") as f:
                json.dump(results, f, indent=2, ensure_ascii=False)
            print(f"Saved to {OUT}")

            # Summary
            print(f"\n=== Page Setup ===")
            ps = results["page_setup"]
            print(f"  Page: {ps['page_width']}x{ps['page_height']}pt")
            print(f"  Margins: T={ps['margin_top']} B={ps['margin_bottom']} L={ps['margin_left']} R={ps['margin_right']}")
            content_h = ps['page_height'] - ps['margin_top'] - ps['margin_bottom']
            print(f"  Content area height: {content_h:.2f}pt")

            print(f"\n=== Page Breaks ===")
            for pb in results["page_breaks"]:
                print(f"  p.{pb['from_page']}→p.{pb['to_page']}: para {pb['para_idx']} y={pb['y_pt']}pt \"{pb['text'][:30]}\"")

            print(f"\n=== Tables ===")
            for t in results["tables"]:
                print(f"  Table {t['table_idx']}:")
                for r in t["rows"]:
                    if "error" in r:
                        print(f"    Row {r['row_idx']}: ERROR {r['error']}")
                    else:
                        rh = r.get("rendered_height", "?")
                        print(f"    Row {r['row_idx']}: p.{r['page']} y={r['y_pt']}pt h={rh} rule={r['height_rule']} spec={r['height_val']} \"{r['text'][:20]}\"")

        finally:
            doc.Close(SaveChanges=False)
    finally:
        word.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
