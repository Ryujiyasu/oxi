"""Measure table row heights in tokumei_08_01-1."""
import win32com.client
import time, os

DOCX = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..",
    "tools", "golden-test", "documents", "docx",
    "de6e32b5960b_tokumei_08_01-1.docx"))

def main():
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    time.sleep(2)

    print(f"LayoutMode: {doc.PageSetup.LayoutMode}")

    # First 3 tables, row heights
    tables = doc.Tables
    print(f"Tables: {tables.Count}")
    for ti in range(1, min(4, tables.Count + 1)):
        tbl = tables(ti)
        rows = tbl.Rows
        print(f"\nTable {ti}: rows={rows.Count}")
        for ri in range(1, min(6, rows.Count + 1)):
            row = rows(ri)
            h = row.Height
            hr = row.HeightRule
            # First cell Y
            cell = tbl.Cell(ri, 1)
            y = cell.Range.Information(6)
            text = cell.Range.Text[:20].replace('\r','').replace('\x07','')
            print(f"  Row{ri}: y={y:7.2f} h={h:6.1f} rule={hr} text='{text}'")

    # Also measure paragraph Y positions on page 1
    print("\n--- P1 paragraphs ---")
    ps = doc.Paragraphs
    for i in range(1, min(20, ps.Count + 1)):
        p = ps(i)
        r = p.Range
        y = r.Information(6)
        page = r.Information(3)
        if page > 1: break
        text = r.Text[:25].replace('\r','').replace('\x07','')
        print(f"P{i:2d}: y={y:7.2f} text='{text}'")

    doc.Close(False)
    word.Quit()

if __name__ == "__main__":
    main()
