"""Compare body start P0_y in LM0 vs LM2 for same font/size combinations.

Goal: determine if a single formula explains both LM0 and LM2 body start positions.
"""
import win32com.client, time, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

def measure(font, size, lm, top_margin=72.0):
    doc = word.Documents.Add()
    time.sleep(0.2)
    ps = doc.PageSetup
    ps.PageWidth = 612; ps.LeftMargin = 90; ps.RightMargin = 90
    ps.TopMargin = top_margin; ps.BottomMargin = 72
    try:
        ps.LayoutMode = lm
    except Exception:
        pass
    rng = doc.Range()
    rng.InsertAfter("ABC\nDEF")
    rng = doc.Range()
    rng.Font.Size = size
    rng.Font.Name = font
    time.sleep(0.1)
    try:
        y1 = doc.Paragraphs(1).Range.Information(6)
        y2 = doc.Paragraphs(2).Range.Information(6)
    except Exception:
        doc.Close(SaveChanges=False)
        return None
    doc.Close(SaveChanges=False)
    return (y1, y2 - y1, y1 - top_margin)

print(f"{'font':<14} {'size':<5} {'LM':<3} {'P0_y':<7} {'P0_h':<7} {'offset_above_topMar':<12}")
for font in ["Times New Roman", "ＭＳ 明朝", "Yu Mincho", "Meiryo"]:
    for size in [10.5, 12, 14, 18, 24]:
        for lm in [0, 2]:
            r = measure(font, size, lm)
            if r:
                y1, p0h, offset = r
                print(f"{font:<14} {size:<5} {lm:<3} {y1:<7.2f} {p0h:<7.2f} {offset:<12.2f}")
        print()

word.Quit()
