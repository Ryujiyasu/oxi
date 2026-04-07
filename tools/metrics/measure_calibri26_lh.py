"""Direct measurement: Calibri 26pt line height in different configs"""
import win32com.client
import time

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True
    word.DisplayAlerts = 0
    time.sleep(1)

    try:
        doc = word.Documents.Add()
        time.sleep(0.5)
        sec = doc.Sections(1)

        # Set LayoutMode to 0 (no grid)
        sec.PageSetup.LayoutMode = 0
        time.sleep(0.3)

        # Add two paragraphs with Calibri 26pt
        rng = doc.Range(0, 0)
        rng.Text = "Line 1\nLine 2\n"
        rng.Font.Name = "Calibri"
        rng.Font.Size = 26
        # Set Single spacing, sa=0, sb=0
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            p.Format.LineSpacingRule = 0  # Single
            p.Format.SpaceBefore = 0
            p.Format.SpaceAfter = 0

        doc.Repaginate()
        time.sleep(0.5)

        print(f"LayoutMode: {sec.PageSetup.LayoutMode}")
        for i in range(1, min(doc.Paragraphs.Count + 1, 4)):
            p = doc.Paragraphs(i)
            py = p.Range.Information(6)
            ls = p.Format.LineSpacing
            text = p.Range.Text.strip()[:20]
            print(f"P{i}: y={py:.2f}pt, ls={ls:.2f}, text='{text}'")

        if doc.Paragraphs.Count >= 2:
            y1 = doc.Paragraphs(1).Range.Information(6)
            y2 = doc.Paragraphs(2).Range.Information(6)
            print(f"\nLine height (gap) = {y2-y1:.2f}pt")
            print(f"In twips = {(y2-y1)*20:.1f}")

        # Also test with other fonts
        for font, size in [("Cambria", 11), ("Calibri", 11), ("MS Gothic", 10.5)]:
            rng = doc.Range(0, 0)
            doc.Content.Delete()
            rng = doc.Range(0, 0)
            rng.Text = "Test1\nTest2\n"
            rng.Font.Name = font
            rng.Font.Size = size
            for i in range(1, doc.Paragraphs.Count + 1):
                p = doc.Paragraphs(i)
                p.Format.LineSpacingRule = 0
                p.Format.SpaceBefore = 0
                p.Format.SpaceAfter = 0
            doc.Repaginate()
            time.sleep(0.3)
            y1 = doc.Paragraphs(1).Range.Information(6)
            y2 = doc.Paragraphs(2).Range.Information(6)
            gap = y2 - y1
            print(f"{font} {size}pt: gap={gap:.2f}pt ({gap*20:.0f}tw)")

        doc.Close(0)
    finally:
        word.Quit()

if __name__ == "__main__":
    measure()
