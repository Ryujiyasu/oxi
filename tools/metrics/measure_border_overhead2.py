"""Measure border overhead on table row height - clean test.

Creates individual documents per border width to avoid interference.
Uses wide tables with short text.
"""
import win32com.client
import sys

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        results = []
        for bw_eighths in [0, 2, 4, 8, 12, 16, 24, 36, 48]:
            bw_pt = bw_eighths / 8.0
            doc = word.Documents.Add()
            doc.PageSetup.TopMargin = 72  # 1 inch

            # Create a 5x2 table (5 rows to see middle row behavior)
            sel = word.Selection
            rng = sel.Range
            tbl = doc.Tables.Add(rng, 5, 2)

            # Set font explicitly
            tbl.Range.Font.Name = "Calibri"
            tbl.Range.Font.Size = 11

            # Set borders
            if bw_eighths > 0:
                for bid in [-1, -2, -3, -4, -5, -6]:
                    try:
                        b = tbl.Borders(bid)
                        b.LineStyle = 1  # single
                        b.LineWidth = bw_eighths
                    except:
                        pass
            else:
                for bid in [-1, -2, -3, -4, -5, -6]:
                    try:
                        tbl.Borders(bid).LineStyle = 0
                    except:
                        pass

            # Short text in each cell
            for ri in range(1, 6):
                for ci in range(1, 3):
                    tbl.Cell(ri, ci).Range.Text = "A"

            # Measure
            ys = []
            for ri in range(1, 6):
                try:
                    y = tbl.Cell(ri, 1).Range.Paragraphs(1).Range.Information(6)
                    ys.append(round(y, 2))
                except:
                    ys.append(-1)

            heights = []
            for i in range(len(ys) - 1):
                heights.append(round(ys[i+1] - ys[i], 2))

            # Also measure cell top padding (distance from cell border to text)
            # Try: Range.Information(wdVerticalPositionRelativeToTextBoundary)
            # Actually, let's measure the table top position
            tbl_y = tbl.Range.Information(6)  # first para in table

            print(f"bw={bw_pt:5.2f}pt  ys={ys}  h={heights}  tbl_y={round(tbl_y, 2)}")
            results.append({"bw": bw_pt, "ys": ys, "heights": heights})

            doc.Close(False)

        return results
    finally:
        word.Quit()


if __name__ == "__main__":
    results = measure()
    print("\n=== Row Height vs Border Width ===")
    print(f"{'bw':>6s}  {'row1':>6s}  {'row2':>6s}  {'row3':>6s}  {'row4':>6s}")
    for r in results:
        hs = r["heights"]
        cols = [f"{h:.2f}" if h >= 0 else "?" for h in hs]
        print(f"{r['bw']:6.2f}  {'  '.join(cols)}")
