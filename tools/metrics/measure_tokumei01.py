"""Measure table row heights for b35123fe8efc_tokumei_08_01."""
import win32com.client
import time, os

DOCX = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..",
    "tools", "golden-test", "documents", "docx",
    "b35123fe8efc_tokumei_08_01.docx"))

def main():
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(DOCX, ReadOnly=True)
    time.sleep(2)

    tbl = doc.Tables(1)
    print(f"Table1: rows={tbl.Rows.Count} cols={tbl.Columns.Count}")

    for ri in range(1, min(10, tbl.Rows.Count + 1)):
        try:
            c = tbl.Cell(ri, 1)
            y = c.Range.Information(6)
            h = tbl.Rows(ri).Height
            hr = tbl.Rows(ri).HeightRule
            text = c.Range.Text[:25].replace('\r','').replace('\x07','')
            # Count paragraphs in cell
            pcount = c.Range.Paragraphs.Count
            # First para spacing
            p1 = c.Range.Paragraphs(1)
            sb = p1.Format.SpaceBefore
            sa = p1.Format.SpaceAfter
            ls = p1.Format.LineSpacing
            lr = p1.Format.LineSpacingRule
            print(f"Row{ri}: y={y:7.1f} h={h:.1f} rule={hr} paras={pcount} sb={sb:.1f} sa={sa:.1f} ls={ls:.1f} lr={lr} text='{text}'")
        except Exception as e:
            print(f"Row{ri}: error {e}")

    # Gaps
    print("\nRow gaps:")
    for ri in range(1, min(10, tbl.Rows.Count)):
        try:
            y1 = tbl.Cell(ri, 1).Range.Information(6)
            y2 = tbl.Cell(ri+1, 1).Range.Information(6)
            print(f"  Row{ri}->Row{ri+1}: gap={y2-y1:.1f}")
        except:
            pass

    doc.Close(False)
    word.Quit()

if __name__ == "__main__":
    main()
