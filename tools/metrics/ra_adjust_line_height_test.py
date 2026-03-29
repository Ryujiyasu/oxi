"""
Ra: adjustLineHeightInTable の正確な意味をCOM計測で確定
index-19のcompat設定と、テーブルセル内の段落行高さを計測
"""
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
    "pipeline_data", "com_measurements", "adjust_line_height_in_table.json"
))


def main():
    pythoncom.CoInitialize()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    word.AutomationSecurity = 3

    results = {}

    try:
        doc = word.Documents.Open(DOCX, ReadOnly=True, AddToRecentFiles=False)
        try:
            # 1. Compat settings
            compat_map = {
                1: "noTabHangInd",
                2: "noSpaceRaiseLower",
                3: "suppressSpBfAfterPgBrk",
                8: "suppressTopSpacing",
                12: "adjustLineHeightInTable",
                13: "noLeading",
            }
            compat_results = {}
            for k, v in compat_map.items():
                try:
                    val = doc.Compatibility(k)
                    compat_results[v] = val
                    print(f"  Compat({k}) {v} = {val}")
                except Exception as e:
                    compat_results[v] = f"error: {e}"
            results["compat"] = compat_results

            # 2. Table cell paragraphs: check their actual line spacing
            print("\n--- Table cell paragraph properties ---")
            table = doc.Tables(1)
            for r_idx in range(1, min(4, table.Rows.Count + 1)):
                cell = table.Cell(r_idx, 1)
                para = cell.Range.Paragraphs(1)
                fmt = para.Format
                ls = fmt.LineSpacing
                lsr = fmt.LineSpacingRule
                sa = fmt.SpaceAfter
                sb = fmt.SpaceBefore
                rule_names = {0: "single", 1: "1.5", 2: "double", 3: "atLeast", 4: "exactly", 5: "multiple"}
                print(f"  Table1 Row{r_idx} Cell1: LineSpacing={ls:.2f}pt Rule={rule_names.get(lsr, lsr)} SA={sa:.2f} SB={sb:.2f}")
                results[f"table1_row{r_idx}"] = {
                    "line_spacing": round(ls, 2),
                    "line_spacing_rule": rule_names.get(lsr, str(lsr)),
                    "space_after": round(sa, 2),
                    "space_before": round(sb, 2),
                }

            # 3. Normal paragraph properties (outside table)
            print("\n--- Normal paragraph properties ---")
            for i in [1, 2, 3]:
                para = doc.Paragraphs(i)
                fmt = para.Format
                ls = fmt.LineSpacing
                lsr = fmt.LineSpacingRule
                sa = fmt.SpaceAfter
                sb = fmt.SpaceBefore
                rule_names = {0: "single", 1: "1.5", 2: "double", 3: "atLeast", 4: "exactly", 5: "multiple"}
                text = para.Range.Text[:30].strip()
                print(f"  Para {i}: LineSpacing={ls:.2f}pt Rule={rule_names.get(lsr, lsr)} SA={sa:.2f} SB={sb:.2f}")
                results[f"para{i}"] = {
                    "line_spacing": round(ls, 2),
                    "line_spacing_rule": rule_names.get(lsr, str(lsr)),
                    "space_after": round(sa, 2),
                    "space_before": round(sb, 2),
                }

            # 4. Measure exact table row heights
            print("\n--- Table row Y positions ---")
            for t_idx in [1, 6, 7]:
                table = doc.Tables(t_idx)
                print(f"  Table {t_idx} ({table.Rows.Count} rows):")
                prev_y = None
                for r_idx in range(1, min(table.Rows.Count + 1, 6)):
                    try:
                        cell = table.Cell(r_idx, 1)
                        y = cell.Range.Information(6)
                        page = cell.Range.Information(3)
                        diff = f" diff={y-prev_y:.2f}" if prev_y and page == cell.Range.Information(3) else ""
                        if prev_y:
                            diff_val = y - prev_y
                            if abs(diff_val) < 100:
                                diff = f" diff={diff_val:.2f}"
                        print(f"    Row {r_idx}: y={y:.2f} p.{page}{diff}")
                        prev_y = y
                    except Exception as e:
                        print(f"    Row {r_idx}: error {e}")

            # Save
            os.makedirs(os.path.dirname(OUT), exist_ok=True)
            with open(OUT, "w", encoding="utf-8") as f:
                json.dump(results, f, indent=2, ensure_ascii=False)
            print(f"\nSaved to {OUT}")

        finally:
            doc.Close(SaveChanges=False)
    finally:
        word.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
