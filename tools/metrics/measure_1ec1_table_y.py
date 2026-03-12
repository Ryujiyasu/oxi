"""#5: Measure the exact Y position of the 納付計画記載欄 table via Word COM."""
import win32com.client
import os
import time

docx_path = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx\1ec1091177b1_006.docx")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(docx_path)
    time.sleep(1)

    print(f"Total tables: {doc.Tables.Count}")

    for ti in range(1, doc.Tables.Count + 1):
        table = doc.Tables(ti)
        tbl_range = table.Range

        # Get table position
        word.Selection.SetRange(tbl_range.Start, tbl_range.Start)
        table_y = float(word.Selection.Information(6))  # wdVerticalPositionRelativeToPage
        table_x = float(word.Selection.Information(5))  # wdHorizontalPositionRelativeToPage
        table_page = tbl_range.Information(3)

        print(f"\nTable {ti}: page={table_page}, top_y={table_y:.2f}pt, left_x={table_x:.2f}pt")
        print(f"  Rows: {table.Rows.Count}, Columns: {table.Columns.Count}")

        # Measure each row's Y position
        for ri in range(1, table.Rows.Count + 1):
            row = table.Rows(ri)
            row_range = row.Range
            word.Selection.SetRange(row_range.Start, row_range.Start)
            row_y = float(word.Selection.Information(6))
            row_height = row.Height
            row_height_rule = row.HeightRule  # 0=wdRowHeightAuto, 1=wdRowHeightAtLeast, 2=wdRowHeightExactly

            rule_names = {0: "Auto", 1: "AtLeast", 2: "Exactly"}
            rule_str = rule_names.get(row_height_rule, str(row_height_rule))

            # First cell text preview
            first_cell_text = ""
            try:
                first_cell_text = table.Cell(ri, 1).Range.Text[:30].replace('\r', '\\r').replace('\x07', '')
            except:
                pass

            print(f"  Row {ri}: y={row_y:.2f}pt, height={row_height:.2f}pt ({rule_str}), text=\"{first_cell_text}\"")

        # Table width
        try:
            print(f"  PreferredWidth: {table.PreferredWidth}")
        except:
            pass

    doc.Close(False)
finally:
    word.Quit()
    print("\nDone.")
