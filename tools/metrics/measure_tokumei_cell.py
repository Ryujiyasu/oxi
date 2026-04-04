"""Measure table cell padding and line height in tokumei_08_01-1."""
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

    tbl = doc.Tables(1)
    print(f"Table1: rows={tbl.Rows.Count}, cols={tbl.Columns.Count}")

    # Default cell margins
    try:
        print(f"TopPadding: {tbl.TopPadding}")
        print(f"BottomPadding: {tbl.BottomPadding}")
        print(f"LeftPadding: {tbl.LeftPadding}")
        print(f"RightPadding: {tbl.RightPadding}")
    except:
        print("Table padding not accessible")

    # Row heights and cell content
    for ri in range(1, min(5, tbl.Rows.Count + 1)):
        row = tbl.Rows(ri)
        h = row.Height
        hr = row.HeightRule
        cell = tbl.Cell(ri, 1)
        cy = cell.Range.Information(6)

        # Cell margins
        try:
            ct = cell.TopPadding
            cb = cell.BottomPadding
        except:
            ct = cb = "?"

        # Paragraph info
        paras = cell.Range.Paragraphs
        print(f"\nRow{ri}: y={cy:.1f} h={h} rule={hr} cellTopPad={ct} cellBotPad={cb} paras={paras.Count}")
        for pi in range(1, min(3, paras.Count + 1)):
            p = paras(pi)
            py = p.Range.Information(6)
            fmt = p.Format
            sb = fmt.SpaceBefore
            sa = fmt.SpaceAfter
            ls = fmt.LineSpacing
            lr = fmt.LineSpacingRule
            text = p.Range.Text[:30].replace('\r','').replace('\x07','')
            print(f"  P{pi}: y={py:.1f} sb={sb:.1f} sa={sa:.1f} ls={ls:.1f} lr={lr} text='{text}'")

    doc.Close(False)
    word.Quit()

if __name__ == "__main__":
    main()
