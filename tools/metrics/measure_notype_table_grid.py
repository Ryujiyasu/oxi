"""COM measurement: table cell line height with no-type docGrid.
Tests whether grid snap applies to table cells when docGrid has no type attribute.
"""
import win32com.client
import os, sys, time

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True
    word.DisplayAlerts = 0
    time.sleep(1)

    try:
        # Create a test document with no-type docGrid (linePitch=360 only)
        doc = word.Documents.Add()
        time.sleep(1)

        # Set page margins to standard
        sec = doc.Sections(1)
        sec.PageSetup.TopMargin = 72  # 1 inch
        sec.PageSetup.BottomMargin = 72
        sec.PageSetup.LeftMargin = 90  # 1.25 inch
        sec.PageSetup.RightMargin = 90

        # Check LayoutMode (grid snap status)
        lm = sec.PageSetup.LayoutMode
        print(f"Default LayoutMode: {lm}")

        # Add a table with 3 rows, 2 cols
        rng = doc.Range(0, 0)
        tbl = doc.Tables.Add(rng, 3, 2)

        # Fill cells with text
        for r in range(1, 4):
            for c in range(1, 3):
                tbl.Cell(r, c).Range.Text = f"Cell R{r}C{c}"
                tbl.Cell(r, c).Range.Font.Name = "Calibri"
                tbl.Cell(r, c).Range.Font.Size = 11

        # Force layout
        doc.Repaginate()
        time.sleep(0.5)

        # Measure row heights via Y coordinates
        print("\n=== No-type docGrid (default) ===")
        print(f"LayoutMode: {sec.PageSetup.LayoutMode}")
        for r in range(1, 4):
            cell_rng = tbl.Cell(r, 1).Range
            y = cell_rng.Information(6)  # wdVerticalPositionRelativeToPage
            # Also get row height property
            row_h = tbl.Rows(r).Height
            rule = tbl.Rows(r).HeightRule
            print(f"  Row {r}: y={y:.2f}pt, Height={row_h:.2f}pt, Rule={rule}")

        # Get paragraph Y in first cell
        for r in range(1, 4):
            p = tbl.Cell(r, 1).Range.Paragraphs(1)
            py = p.Range.Information(6)
            ls = p.Format.LineSpacing
            lsr = p.Format.LineSpacingRule
            print(f"  Row {r} Para: y={py:.2f}pt, LineSpacing={ls:.2f}pt, Rule={lsr}")

        # Now add body paragraphs below the table for comparison
        end_rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        end_rng.InsertAfter("\nBody paragraph 1\nBody paragraph 2\n")
        doc.Repaginate()
        time.sleep(0.3)

        # Measure body paragraph Y
        total_paras = doc.Paragraphs.Count
        print(f"\nBody paragraphs (total {total_paras}):")
        for i in range(1, min(total_paras + 1, 10)):
            p = doc.Paragraphs(i)
            py = p.Range.Information(6)
            text = p.Range.Text.strip()[:30]
            print(f"  P{i}: y={py:.2f}pt, text='{text}'")

        # Now explicitly set docGrid type=lines and compare
        print("\n=== Setting docGrid type=lines ===")
        # Access via XML
        sec.PageSetup.LayoutMode = 1  # wdLayoutModeLineGrid
        doc.Repaginate()
        time.sleep(0.5)

        lm2 = sec.PageSetup.LayoutMode
        print(f"LayoutMode after set: {lm2}")

        for r in range(1, 4):
            cell_rng = tbl.Cell(r, 1).Range
            y = cell_rng.Information(6)
            print(f"  Row {r}: y={y:.2f}pt")

        for i in range(1, min(total_paras + 1, 10)):
            p = doc.Paragraphs(i)
            py = p.Range.Information(6)
            text = p.Range.Text.strip()[:30]
            print(f"  P{i}: y={py:.2f}pt, text='{text}'")

        # Now set LayoutMode=0 (no grid) for comparison
        print("\n=== Setting LayoutMode=0 (no grid) ===")
        sec.PageSetup.LayoutMode = 0
        doc.Repaginate()
        time.sleep(0.5)

        for r in range(1, 4):
            cell_rng = tbl.Cell(r, 1).Range
            y = cell_rng.Information(6)
            print(f"  Row {r}: y={y:.2f}pt")

        for i in range(1, min(total_paras + 1, 10)):
            p = doc.Paragraphs(i)
            py = p.Range.Information(6)
            text = p.Range.Text.strip()[:30]
            print(f"  P{i}: y={py:.2f}pt, text='{text}'")

        doc.Close(0)

    finally:
        word.Quit()

if __name__ == "__main__":
    measure()
