"""Measure table row Y positions in gen_tables.docx"""
import win32com.client
import time, os

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True
    word.DisplayAlerts = 0
    time.sleep(1)

    docx = os.path.abspath("tools/golden-test/documents/docx/gen_tables.docx")
    doc = word.Documents.Open(docx, ReadOnly=True)
    time.sleep(1)

    try:
        sec = doc.Sections(1)
        lm = sec.PageSetup.LayoutMode
        print(f"LayoutMode: {lm}")
        print(f"TopMargin: {sec.PageSetup.TopMargin:.2f}pt")

        for ti in range(1, doc.Tables.Count + 1):
            tbl = doc.Tables(ti)
            print(f"\nTable {ti} ({tbl.Rows.Count} rows, {tbl.Columns.Count} cols):")
            prev_y = None
            for r in range(1, tbl.Rows.Count + 1):
                try:
                    y = tbl.Cell(r, 1).Range.Information(6)
                    h = tbl.Rows(r).Height
                    rule = tbl.Rows(r).HeightRule
                    text = tbl.Cell(r, 1).Range.Text.strip()[:25]
                    gap = f" gap={y-prev_y:.2f}" if prev_y else ""
                    prev_y = y
                    print(f"  Row {r}: y={y:.2f}pt, h={h:.2f}pt, rule={rule}{gap} '{text}'")
                except Exception as e:
                    print(f"  Row {r}: error {e}")

        # Check border widths
        for ti in range(1, min(doc.Tables.Count + 1, 3)):
            tbl = doc.Tables(ti)
            try:
                bw = tbl.Borders(1).LineWidth  # wdBorderTop
                print(f"\nTable {ti} top border width: {bw}")
                bw_l = tbl.Borders(4).LineWidth  # wdBorderLeft
                print(f"Table {ti} left border width: {bw_l}")
            except:
                pass

        doc.Close(0)
    finally:
        word.Quit()

if __name__ == "__main__":
    measure()
