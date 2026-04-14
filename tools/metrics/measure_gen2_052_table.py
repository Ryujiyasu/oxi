"""Measure table cell positions for gen2_052_Privacy_Policy."""
import win32com.client, os, time

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc_path = os.path.abspath("tools/golden-test/documents/docx/gen2_052_Privacy_Policy.docx")
    doc = word.Documents.Open(doc_path, ReadOnly=True)
    time.sleep(1)

    ps = doc.Sections(1).PageSetup
    print(f"Margins L/R: {ps.LeftMargin:.2f} / {ps.RightMargin:.2f}")
    print(f"Content width: {ps.PageWidth - ps.LeftMargin - ps.RightMargin:.2f}")

    for ti in range(1, min(doc.Tables.Count + 1, 3)):
        t = doc.Tables(ti)
        print(f"\nTable {ti}: {t.Rows.Count} rows x {t.Columns.Count} cols")

        # Table position
        for ri in range(1, min(t.Rows.Count + 1, 4)):
            for ci in range(1, min(t.Columns.Count + 1, 4)):
                cell = t.Cell(ri, ci)
                rng = cell.Range
                x = rng.Information(5)  # wdHorizontalPositionRelativeToPage
                y = rng.Information(6)  # wdVerticalPositionRelativeToPage
                w = cell.Width
                h = cell.Height
                text = rng.Text[:15].replace('\r', '').replace('\x07', '')
                print(f"  R{ri}C{ci}: x={x:.1f} y={y:.1f} w={w:.1f} h={h:.1f} \"{text}\"")

        # Also check column widths
        print(f"  Column widths:")
        for ci in range(1, min(t.Columns.Count + 1, 5)):
            print(f"    Col{ci}: {t.Columns(ci).Width:.2f}pt")

    doc.Close(False)
    word.Quit()

if __name__ == '__main__':
    measure()
