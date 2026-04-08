"""Investigate TNR 14pt in LM1: nat=19pt, snap=21pt (NOT a multiple of 18 pitch).

Hypotheses to test:
1. P2 height includes spacing-after we missed
2. Word uses ceil(natural * 1.x) for grid-incompatible cases
3. Per-LM mode interaction with default linePitch
"""
import win32com.client, time, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

def measure_full(font, size, lm, line_240ths=None, after_tw=None):
    """Returns dict with various Y measurements."""
    doc = word.Documents.Add()
    time.sleep(0.2)
    ps = doc.PageSetup
    ps.PageWidth = 612; ps.LeftMargin = 90; ps.RightMargin = 90
    ps.TopMargin = 72; ps.BottomMargin = 72
    ps.LayoutMode = lm
    rng = doc.Range()
    rng.InsertAfter("ABC\nDEF\nGHI")
    rng = doc.Range()
    rng.Font.Size = size
    rng.Font.Name = font
    if line_240ths is not None:
        for p in [doc.Paragraphs(1), doc.Paragraphs(2), doc.Paragraphs(3)]:
            p.LineSpacingRule = 5
            p.LineSpacing = size * (line_240ths / 240.0)
    if after_tw is not None:
        for p in [doc.Paragraphs(1), doc.Paragraphs(2), doc.Paragraphs(3)]:
            p.SpaceAfter = after_tw / 20.0
    time.sleep(0.1)
    try:
        y1 = doc.Paragraphs(1).Range.Information(6)
        y2 = doc.Paragraphs(2).Range.Information(6)
        y3 = doc.Paragraphs(3).Range.Information(6)
        sa1 = doc.Paragraphs(1).SpaceAfter
        ls1 = doc.Paragraphs(1).LineSpacing
        lsr1 = doc.Paragraphs(1).LineSpacingRule
    except Exception as e:
        doc.Close(SaveChanges=False)
        return None
    doc.Close(SaveChanges=False)
    return {
        "y1": y1, "y2": y2, "y3": y3,
        "p1_height": round(y2 - y1, 3),
        "p2_height": round(y3 - y2, 3),
        "spaceAfter": sa1, "lineSpacing": ls1, "lineSpacingRule": lsr1,
    }

print("=== TNR 14pt anomaly investigation ===")
print()

print("LM=0 (no grid) variants:")
for label, line, after in [
    ("default", None, None),
    ("line=240 after=0", 240, 0),
    ("line=240 after=200", 240, 200),
    ("line=276 after=0", 276, 0),
]:
    r = measure_full("Times New Roman", 14, 0, line, after)
    print(f"  {label:<22}: P1h={r['p1_height']} P2h={r['p2_height']} SA={r['spaceAfter']} LS={r['lineSpacing']} LSR={r['lineSpacingRule']}")

print()
print("LM=1 (lines) variants:")
for label, line, after in [
    ("default", None, None),
    ("line=240 after=0", 240, 0),
    ("line=240 after=200", 240, 200),
    ("line=276 after=0", 276, 0),
]:
    r = measure_full("Times New Roman", 14, 1, line, after)
    print(f"  {label:<22}: P1h={r['p1_height']} P2h={r['p2_height']} SA={r['spaceAfter']} LS={r['lineSpacing']} LSR={r['lineSpacingRule']}")

print()
print("LM=2 (linesAndChars) variants:")
for label, line, after in [
    ("default", None, None),
    ("line=240 after=0", 240, 0),
]:
    r = measure_full("Times New Roman", 14, 2, line, after)
    print(f"  {label:<22}: P1h={r['p1_height']} P2h={r['p2_height']} SA={r['spaceAfter']} LS={r['lineSpacing']} LSR={r['lineSpacingRule']}")

word.Quit()
