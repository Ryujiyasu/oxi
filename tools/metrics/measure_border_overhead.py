"""Measure whether border width affects table row height in Word.

Creates test tables with different border widths and measures actual row heights.
"""
import win32com.client
import sys
import os

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Add()

        # Set page margins small for more space
        doc.PageSetup.TopMargin = 72  # 1 inch

        results = []

        for border_width_eighths in [0, 2, 4, 8, 12, 16, 24]:
            # border_width in 1/8 pt: 0=none, 2=0.25pt, 4=0.5pt, 8=1pt, 12=1.5pt, 16=2pt, 24=3pt
            bw_pt = border_width_eighths / 8.0

            # Add a paragraph before each table
            sel = word.Selection
            sel.TypeText(f"Border width: {bw_pt}pt")
            sel.TypeParagraph()

            # Create a 3x2 table
            rng = sel.Range
            tbl = doc.Tables.Add(rng, 3, 2)

            # Set all borders
            if border_width_eighths > 0:
                for border_id in [-1, -2, -3, -4, -5, -6]:  # top, left, bottom, right, insideH, insideV
                    try:
                        b = tbl.Borders(border_id)
                        b.LineStyle = 1  # wdLineStyleSingle
                        b.LineWidth = border_width_eighths
                    except:
                        pass
            else:
                # No borders
                for border_id in [-1, -2, -3, -4, -5, -6]:
                    try:
                        tbl.Borders(border_id).LineStyle = 0  # wdLineStyleNone
                    except:
                        pass

            # Add content to cells
            for ri in range(1, 4):
                for ci in range(1, 3):
                    tbl.Cell(ri, ci).Range.Text = f"R{ri}C{ci}"

            # Move after table
            sel.MoveDown(5, 1)  # wdLine
            sel.TypeParagraph()

            # Measure row Y positions
            row_ys = []
            for ri in range(1, 4):
                try:
                    y = tbl.Cell(ri, 1).Range.Paragraphs(1).Range.Information(6)
                    row_ys.append(round(y, 2))
                except:
                    row_ys.append(-1)

            # Get heights
            row_heights = []
            for i in range(len(row_ys) - 1):
                if row_ys[i] >= 0 and row_ys[i+1] >= 0:
                    row_heights.append(round(row_ys[i+1] - row_ys[i], 2))
                else:
                    row_heights.append(-1)

            # Get Row.Height property
            spec_heights = []
            for ri in range(1, 4):
                try:
                    spec_heights.append(round(tbl.Rows(ri).Height, 2))
                except:
                    spec_heights.append(-1)

            result = {
                "border_width_pt": bw_pt,
                "row_ys": row_ys,
                "actual_row_heights": row_heights,
                "spec_row_heights": spec_heights,
            }
            results.append(result)
            print(f"border={bw_pt:5.2f}pt  row_ys={row_ys}  actual_h={row_heights}")

        doc.Close(False)
        return results
    finally:
        word.Quit()


if __name__ == "__main__":
    results = measure()
    print("\n=== Summary ===")
    print(f"{'border':>8s}  {'row1_h':>8s}  {'row2_h':>8s}")
    for r in results:
        hs = r["actual_row_heights"]
        h1 = f"{hs[0]:.2f}" if len(hs) > 0 and hs[0] >= 0 else "?"
        h2 = f"{hs[1]:.2f}" if len(hs) > 1 and hs[1] >= 0 else "?"
        print(f"{r['border_width_pt']:8.2f}  {h1:>8s}  {h2:>8s}")
