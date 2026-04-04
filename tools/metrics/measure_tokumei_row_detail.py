"""Verify: is Row2 height really 32.9pt or is it actually 29.5pt?
Check actual row boundaries, not just text Y positions."""
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

    # Check each row's actual rendered height
    for ri in range(1, min(8, tbl.Rows.Count + 1)):
        row = tbl.Rows(ri)
        h = row.Height
        hr = row.HeightRule
        # Get cell 1 text Y and cell range boundaries
        c = tbl.Cell(ri, 1)
        ty = c.Range.Information(6)
        # Try to get row top/bottom via selection
        print(f"Row{ri}: textY={ty:.1f} Height={h:.1f} Rule={hr}")

    # More precise: use Row boundaries
    # Row top = previous row bottom, or table top for row 1
    # Table top = Row1 textY - spaceBefore - paddingTop
    print("\nRow gaps (textY difference):")
    for ri in range(1, min(7, tbl.Rows.Count)):
        y1 = tbl.Cell(ri, 1).Range.Information(6)
        y2 = tbl.Cell(ri+1, 1).Range.Information(6)
        print(f"  Row{ri}->Row{ri+1}: gap={y2-y1:.1f}")

    # Also check Row2 cell paragraphs for spaceBefore/After
    print("\nRow2 cell details:")
    for ci in range(1, tbl.Columns.Count + 1):
        try:
            c = tbl.Cell(2, ci)
            r = c.Range
            paras = r.Paragraphs
            py = r.Information(6)
            for pi in range(1, min(3, paras.Count+1)):
                p = paras(pi)
                sb = p.Format.SpaceBefore
                sa = p.Format.SpaceAfter
                ls = p.Format.LineSpacing
                lr = p.Format.LineSpacingRule
                print(f"  Cell({2},{ci}) P{pi}: y={p.Range.Information(6):.1f} sb={sb:.1f} sa={sa:.1f} ls={ls:.1f} lr={lr}")
        except:
            pass

    doc.Close(False)
    word.Quit()

if __name__ == "__main__":
    main()
