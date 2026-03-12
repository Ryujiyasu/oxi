"""#5: Measure table position via Word COM — handle merged cells."""
import win32com.client
import os
import time

docx_path = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx\1ec1091177b1_006.docx")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(docx_path)
    time.sleep(1)

    table = doc.Tables(1)
    tbl_range = table.Range
    word.Selection.SetRange(tbl_range.Start, tbl_range.Start)
    table_y = float(word.Selection.Information(6))
    table_x = float(word.Selection.Information(5))
    print(f"Table 1: top_y={table_y:.2f}pt, left_x={table_x:.2f}pt")
    print(f"  Rows: {table.Rows.Count}, Columns: {table.Columns.Count}")

    # Access cells individually to handle merged cells
    for ri in range(1, table.Rows.Count + 1):
        for ci in range(1, table.Columns.Count + 1):
            try:
                cell = table.Cell(ri, ci)
                cell_range = cell.Range
                word.Selection.SetRange(cell_range.Start, cell_range.Start)
                cell_y = float(word.Selection.Information(6))
                cell_x = float(word.Selection.Information(5))
                text = cell_range.Text[:25].replace('\r', '\\r').replace('\x07', '').replace('\n', '\\n')
                print(f"  Cell({ri},{ci}): y={cell_y:.2f}, x={cell_x:.2f}, text=\"{text}\"")
            except Exception as e:
                if "merged" in str(e).lower() or "削除" in str(e) or "25305" in str(e):
                    print(f"  Cell({ri},{ci}): MERGED/DELETED")
                else:
                    print(f"  Cell({ri},{ci}): ERROR: {e}")

    # Also get table bottom by checking position at end of table range
    word.Selection.SetRange(tbl_range.End - 1, tbl_range.End - 1)
    table_bottom_y = float(word.Selection.Information(6))
    print(f"\nTable bottom (last char): y={table_bottom_y:.2f}pt")

    # Table width from preferred width
    try:
        pw = table.PreferredWidth
        pwt = table.PreferredWidthType  # 1=wdPreferredWidthPercent, 2=wdPreferredWidthPoints, 3=wdPreferredWidthAuto
        type_names = {1: "Percent", 2: "Points", 3: "Auto"}
        print(f"  PreferredWidth: {pw} ({type_names.get(pwt, pwt)})")
    except:
        pass

    doc.Close(False)
finally:
    word.Quit()
    print("\nDone.")
