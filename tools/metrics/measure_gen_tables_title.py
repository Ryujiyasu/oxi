"""Measure exact Title paragraph height in gen_tables.docx"""
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
        sec = doc.Sections(1)
        print(f"LayoutMode: {sec.PageSetup.LayoutMode}")
        print(f"TopMargin: {sec.PageSetup.TopMargin:.2f}pt")

        # Title paragraph
        p1 = doc.Paragraphs(1)
        y1 = p1.Range.Information(6)
        ls1 = p1.Format.LineSpacing
        lsr1 = p1.Format.LineSpacingRule
        sb1 = p1.Format.SpaceBefore
        sa1 = p1.Format.SpaceAfter
        fn1 = p1.Range.Font.Name
        fs1 = p1.Range.Font.Size
        bold = p1.Range.Font.Bold
        style = p1.Style.NameLocal
        print(f"P1: y={y1:.2f} ls={ls1:.2f} lsr={lsr1} sb={sb1:.2f} sa={sa1:.2f}")
        print(f"    fn={fn1} fs={fs1:.1f} bold={bold} style={style}")

        # Check if there's a paragraph border
        bdr = p1.Format.Borders
        for bi in range(1, 5):
            try:
                b = bdr(bi)
                if b.LineStyle > 0:
                    print(f"    Border {bi}: style={b.LineStyle} width={b.LineWidth} color={b.ColorIndex}")
            except:
                pass

        # Second element (table header)
        p2 = doc.Paragraphs(2)
        y2 = p2.Range.Information(6)
        print(f"P2: y={y2:.2f} (table header cell)")
        print(f"P1→P2 gap: {y2-y1:.2f}")
        print(f"Expected: line_height + max(sa={sa1:.1f}, sb=0) = lh + {sa1:.1f}")
        implied_lh = y2 - y1 - sa1
        print(f"Implied line_height: {implied_lh:.2f}")

        # Check if there's a subtitle between title and table
        for i in range(1, 6):
            p = doc.Paragraphs(i)
            py = p.Range.Information(6)
            text = p.Range.Text.strip()[:30]
            print(f"  P{i}: y={py:.2f} '{text}'")

        doc.Close(0)
    finally:
        word.Quit()

if __name__ == "__main__":
    measure()
