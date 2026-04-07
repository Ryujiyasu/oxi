"""Measure line heights for LayoutMode≥1 with Multiple spacing (1.15x)"""
import win32com.client
import time

def measure():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True
    word.DisplayAlerts = 0
    time.sleep(1)

    try:
        for lm, name in [(0, "LM0"), (1, "LM1-lines"), (2, "LM2-lineGrid")]:
            doc = word.Documents.Add()
            time.sleep(0.5)
            sec = doc.Sections(1)
            sec.PageSetup.LayoutMode = lm
            time.sleep(0.3)

            rng = doc.Range(0, 0)
            rng.Text = "Line1\nLine2\nLine3\n"
            rng.Font.Name = "Calibri"
            rng.Font.Size = 11

            # Set 1.15x Multiple spacing
            for i in range(1, doc.Paragraphs.Count + 1):
                p = doc.Paragraphs(i)
                p.Format.LineSpacingRule = 5  # wdLineSpaceMultiple
                p.Format.LineSpacing = 13.8  # 1.15 * 12
                p.Format.SpaceBefore = 0
                p.Format.SpaceAfter = 0

            doc.Repaginate()
            time.sleep(0.3)

            print(f"\n=== {name} (LayoutMode={sec.PageSetup.LayoutMode}) ===")
            for i in range(1, min(doc.Paragraphs.Count + 1, 4)):
                p = doc.Paragraphs(i)
                py = p.Range.Information(6)
                ls = p.Format.LineSpacing
                lsr = p.Format.LineSpacingRule
                print(f"  P{i}: y={py:.2f}pt, ls={ls:.2f}, lsr={lsr}")

            if doc.Paragraphs.Count >= 2:
                y1 = doc.Paragraphs(1).Range.Information(6)
                y2 = doc.Paragraphs(2).Range.Information(6)
                print(f"  Gap = {y2-y1:.2f}pt ({(y2-y1)*20:.0f}tw)")

            # Also test Single spacing
            for i in range(1, doc.Paragraphs.Count + 1):
                p = doc.Paragraphs(i)
                p.Format.LineSpacingRule = 0  # Single
                p.Format.SpaceBefore = 0
                p.Format.SpaceAfter = 0
            doc.Repaginate()
            time.sleep(0.3)
            y1 = doc.Paragraphs(1).Range.Information(6)
            y2 = doc.Paragraphs(2).Range.Information(6)
            print(f"  Single gap = {y2-y1:.2f}pt ({(y2-y1)*20:.0f}tw)")

            doc.Close(0)

    finally:
        word.Quit()

if __name__ == "__main__":
    measure()
