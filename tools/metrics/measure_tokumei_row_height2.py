"""Test with linePitch=292 (14.6pt) matching tokumei document."""
import win32com.client
import time

def main():
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    doc = word.Documents.Add()
    time.sleep(1)

    doc.PageSetup.LayoutMode = 1

    # Set linePitch to 292 twips
    # PageSetup doesn't have direct linePitch setter, need to use XML
    # Instead, try to match by using linesAndChars with appropriate settings
    # Actually, just set LinesPage to change effective pitch
    # linePitch = (pageHeight - marginTop - marginBottom) / linesPage
    # Default: (841.9 - 72 - 72) / lines = 14.6 -> lines = 697.9/14.6 = 47.8 -> 48?
    # Actually linePitch is set in sectPr/docGrid, not via COM directly

    # Create table
    tbl = doc.Tables.Add(doc.Range(0,0), 3, 1)

    cell_r = tbl.Cell(1, 1).Range
    cell_r.Font.Name = "ＭＳ 明朝"
    cell_r.Font.Size = 10.5
    cell_r.Text = "Test beforeLines"

    p = tbl.Cell(1,1).Range.Paragraphs(1)
    p.Format.LineSpacing = 12
    p.Format.LineSpacingRule = 4  # exact
    p.Format.SpaceBefore = 4.35
    p.Format.SpaceAfter = 0

    tbl.Cell(2,1).Range.Text = "Row 2"
    tbl.Cell(3,1).Range.Text = "Row 3"

    time.sleep(1)

    print(f"LayoutMode: {doc.PageSetup.LayoutMode}")
    print(f"LinesPage: {doc.PageSetup.LinesPage}")

    y1 = tbl.Cell(1,1).Range.Information(6)
    y2 = tbl.Cell(2,1).Range.Information(6)
    y3 = tbl.Cell(3,1).Range.Information(6)
    print(f"Default pitch 360:")
    print(f"  Row1 y={y1:.1f}, Row2 y={y2:.1f}, Row3 y={y3:.1f}")
    print(f"  Row1 height={y2-y1:.1f}, Row2 height={y3-y2:.1f}")

    # Try different line spacings
    for ls in [10, 12, 14, 14.6, 16, 18]:
        p.Format.LineSpacing = ls
        time.sleep(0.3)
        y1 = tbl.Cell(1,1).Range.Information(6)
        y2 = tbl.Cell(2,1).Range.Information(6)
        print(f"  exact {ls:5.1f}pt sb=4.35: gap={y2-y1:.1f}")

    # Try different space_before values with exact 12pt
    p.Format.LineSpacing = 12
    for sb in [0, 2, 4, 4.35, 6, 8, 10]:
        p.Format.SpaceBefore = sb
        time.sleep(0.3)
        y1 = tbl.Cell(1,1).Range.Information(6)
        y2 = tbl.Cell(2,1).Range.Information(6)
        print(f"  exact 12pt sb={sb:5.2f}: gap={y2-y1:.1f}")

    doc.Close(False)
    word.Quit()

if __name__ == "__main__":
    main()
