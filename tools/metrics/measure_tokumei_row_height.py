"""Deep measure: why is tokumei Row1 auto height 25pt when content = 16.3pt?

Row1: 1 cell (gridSpan=4), 1 paragraph, sb=4.3 ls=12(exact) sa=0 padding=0.
Content height = 4.3+12 = 16.3pt. But Word renders 25pt (136-111).

Hypothesis: Word uses natural font height (not exact) for row height calculation,
then applies exact only for text positioning within the line.
"""
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

    # Measure Row1 in detail
    row1 = tbl.Rows(1)
    cell = tbl.Cell(1, 1)
    r = cell.Range

    # Y of Row1 content
    y1 = r.Information(6)
    # Y of Row2 content
    cell2 = tbl.Cell(2, 1)
    y2 = cell2.Range.Information(6)
    row_height = y2 - y1

    print(f"Row1 y={y1:.1f}, Row2 y={y2:.1f}, row_height={row_height:.1f}")
    print(f"Row1.Height={row1.Height}, HeightRule={row1.HeightRule}")

    # Paragraph details
    p = cell.Range.Paragraphs(1)
    fmt = p.Format
    print(f"P1: sb={fmt.SpaceBefore:.2f} sa={fmt.SpaceAfter:.2f}")
    print(f"    ls={fmt.LineSpacing:.2f} lr={fmt.LineSpacingRule}")
    print(f"    font={r.Font.Name} sz={r.Font.Size}")

    # Now create a test: same setup, measure what determines the height
    print("\n--- Fresh document test ---")
    doc2 = word.Documents.Add()
    time.sleep(1)

    # Set grid
    doc2.PageSetup.LayoutMode = 1
    # Insert table
    tbl2 = doc2.Tables.Add(doc2.Range(0,0), 3, 1)

    # Set cell content with exact spacing + beforeLines
    cell_r = tbl2.Cell(1, 1).Range
    cell_r.Font.Name = "MS 明朝"
    cell_r.Font.Size = 10.5
    cell_r.Text = "Test line"

    # Set spacing
    p2 = tbl2.Cell(1,1).Range.Paragraphs(1)
    p2.Format.LineSpacing = 12
    p2.Format.LineSpacingRule = 4  # wdLineSpaceExactly
    p2.Format.SpaceBefore = 4.3
    p2.Format.SpaceAfter = 0

    # Row 2 text
    tbl2.Cell(2,1).Range.Text = "Row 2"

    time.sleep(1)

    y1 = tbl2.Cell(1,1).Range.Information(6)
    y2 = tbl2.Cell(2,1).Range.Information(6)
    print(f"Fresh: Row1 y={y1:.1f}, Row2 y={y2:.1f}, gap={y2-y1:.1f}")

    # Now try without exact spacing (single/auto)
    p2.Format.LineSpacingRule = 0  # wdLineSpaceAutomatic (single)
    p2.Format.SpaceBefore = 4.3
    time.sleep(0.5)

    y1b = tbl2.Cell(1,1).Range.Information(6)
    y2b = tbl2.Cell(2,1).Range.Information(6)
    print(f"Auto: Row1 y={y1b:.1f}, Row2 y={y2b:.1f}, gap={y2b-y1b:.1f}")

    # Now exact 12pt, no spaceBefore
    p2.Format.LineSpacingRule = 4  # exact
    p2.Format.LineSpacing = 12
    p2.Format.SpaceBefore = 0
    time.sleep(0.5)

    y1c = tbl2.Cell(1,1).Range.Information(6)
    y2c = tbl2.Cell(2,1).Range.Information(6)
    print(f"Exact no sb: Row1 y={y1c:.1f}, Row2 y={y2c:.1f}, gap={y2c-y1c:.1f}")

    # Now exact 12pt, beforeLines=30 via XML would require different approach
    # Try spaceBefore = 4.38 (= 30% of 14.6)
    p2.Format.SpaceBefore = 4.38
    time.sleep(0.5)

    y1d = tbl2.Cell(1,1).Range.Information(6)
    y2d = tbl2.Cell(2,1).Range.Information(6)
    print(f"Exact sb=4.38: Row1 y={y1d:.1f}, Row2 y={y2d:.1f}, gap={y2d-y1d:.1f}")

    doc2.Close(False)
    doc.Close(False)
    word.Quit()

if __name__ == "__main__":
    main()
