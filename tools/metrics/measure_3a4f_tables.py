"""Measure table row heights for 3a4f9fbe1a83_001620506.docx via COM."""
import win32com.client
import os, time

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    path = os.path.abspath("tools/golden-test/documents/docx/3a4f9fbe1a83_001620506.docx")
    doc = word.Documents.Open(path, ReadOnly=True)
    time.sleep(1)

    print(f"Tables: {doc.Tables.Count}")

    for i in range(1, min(doc.Tables.Count + 1, 11)):
        tbl = doc.Tables(i)
        rows = tbl.Rows.Count
        cols = tbl.Columns.Count

        # Get table position (first row, first cell)
        try:
            rng = tbl.Cell(1, 1).Range
            y = rng.Information(6)  # wdVerticalPositionRelativeToPage
            page = rng.Information(3)  # wdActiveEndPageNumber

            # Get row height
            row_h = tbl.Rows(1).Height
            row_rule = tbl.Rows(1).HeightRule  # 0=auto, 1=atLeast, 2=exact

            # Last row bottom
            last_rng = tbl.Cell(rows, 1).Range
            last_y = last_rng.Information(6)

            print(f"Table {i}: {rows}x{cols} page={page} y={y:.2f} row_h={row_h:.2f} rule={row_rule} last_y={last_y:.2f}")
        except Exception as e:
            print(f"Table {i}: error {e}")

    doc.Close(SaveChanges=False)
    word.Quit()

if __name__ == "__main__":
    measure()
