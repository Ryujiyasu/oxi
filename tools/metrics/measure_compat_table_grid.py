"""COM measurement: table cell grid snap behavior per compatibilityMode.
Tests if compatMode=14 (Word 2010) disables grid snap in table cells.
"""
import win32com.client
import time

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True
    word.DisplayAlerts = 0
    time.sleep(1)

    try:
        for compat_mode in [15, 14, 12]:
            doc = word.Documents.Add()
            time.sleep(0.5)

            # Set compat mode via XML manipulation
            sec = doc.Sections(1)

            # Add a table
            rng = doc.Range(0, 0)
            tbl = doc.Tables.Add(rng, 3, 2)
            for r in range(1, 4):
                for c in range(1, 3):
                    tbl.Cell(r, c).Range.Text = f"R{r}C{c}"
                    tbl.Cell(r, c).Range.Font.Name = "Calibri"
                    tbl.Cell(r, c).Range.Font.Size = 11

            # Set compat mode
            if compat_mode != 15:
                doc.SetCompatibilityMode(compat_mode)

            doc.Repaginate()
            time.sleep(0.5)

            lm = sec.PageSetup.LayoutMode
            print(f"\n=== CompatMode={compat_mode}, LayoutMode={lm} ===")

            for r in range(1, 4):
                y = tbl.Cell(r, 1).Range.Information(6)
                print(f"  Row {r}: y={y:.2f}pt")

            if tbl.Rows.Count >= 2:
                y1 = tbl.Cell(1, 1).Range.Information(6)
                y2 = tbl.Cell(2, 1).Range.Information(6)
                print(f"  Row gap: {y2-y1:.2f}pt")

            # Also add body paragraph after table for reference
            end_rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
            end_rng.InsertAfter("\nBody text")
            end_rng.Font.Name = "Calibri"
            end_rng.Font.Size = 11
            doc.Repaginate()
            time.sleep(0.3)

            # Find body paragraph
            total_paras = doc.Paragraphs.Count
            for i in range(1, total_paras + 1):
                p = doc.Paragraphs(i)
                text = p.Range.Text.strip()
                if text == "Body text":
                    py = p.Range.Information(6)
                    print(f"  Body para: y={py:.2f}pt")
                    break

            doc.Close(0)

    finally:
        word.Quit()

if __name__ == "__main__":
    measure()
