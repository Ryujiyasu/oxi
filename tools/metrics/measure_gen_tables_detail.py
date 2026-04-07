"""Detailed measurement of gen_tables.docx - paragraphs + table positions"""
import win32com.client
import time, os

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True
    word.DisplayAlerts = 0
    time.sleep(1)

    docx = os.path.abspath("tools/golden-test/documents/docx/gen_tables.docx")
    doc = word.Documents.Open(docx, ReadOnly=True)
    time.sleep(1)

    try:
        # All paragraphs before first table
        print("=== Paragraphs ===")
        for i in range(1, min(doc.Paragraphs.Count + 1, 30)):
            p = doc.Paragraphs(i)
            py = p.Range.Information(6)
            text = p.Range.Text.strip()[:50]
            ls = p.Format.LineSpacing
            lsr = p.Format.LineSpacingRule
            sb = p.Format.SpaceBefore
            sa = p.Format.SpaceAfter
            fn = p.Range.Font.Name
            fs = p.Range.Font.Size
            style = p.Style.NameLocal
            if text or True:
                print(f"  P{i}: y={py:.2f}pt ls={ls:.2f} lsr={lsr} sb={sb:.2f} sa={sa:.2f} fn={fn} fs={fs:.1f} style={style}")
                if text:
                    print(f"        text='{text}'")

        # Table details
        for ti in range(1, min(doc.Tables.Count + 1, 3)):
            tbl = doc.Tables(ti)
            print(f"\n=== Table {ti} ===")
            # Cell padding
            try:
                print(f"  TopPadding: {tbl.TopPadding:.2f}")
                print(f"  BottomPadding: {tbl.BottomPadding:.2f}")
                print(f"  LeftPadding: {tbl.LeftPadding:.2f}")
                print(f"  RightPadding: {tbl.RightPadding:.2f}")
            except:
                pass
            # Borders
            try:
                for bi, bname in [(1, "top"), (2, "left"), (3, "bottom"), (4, "right")]:
                    b = tbl.Borders(bi)
                    print(f"  Border {bname}: width={b.LineWidth}, style={b.LineStyle}")
            except:
                pass

            # Row details
            prev_y = None
            for r in range(1, min(tbl.Rows.Count + 1, 6)):
                y = tbl.Cell(r, 1).Range.Information(6)
                p = tbl.Cell(r, 1).Range.Paragraphs(1)
                ls = p.Format.LineSpacing
                fn = p.Range.Font.Name
                fs = p.Range.Font.Size
                gap = f" gap={y-prev_y:.2f}" if prev_y else ""
                prev_y = y
                text = tbl.Cell(r, 1).Range.Text.strip()[:20]
                print(f"  Row {r}: y={y:.2f}pt ls={ls:.2f} fn={fn} fs={fs:.1f}{gap} '{text}'")

        doc.Close(0)
    finally:
        word.Quit()

if __name__ == "__main__":
    measure()
