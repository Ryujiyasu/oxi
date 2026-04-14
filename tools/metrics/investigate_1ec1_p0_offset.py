"""Investigate body start offset formula for oversized P0 in linesAndChars mode."""
import win32com.client, time, os, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

# Open the actual doc and read P0 details
path = os.path.abspath("tools/golden-test/documents/docx/1ec1091177b1_006.docx")
doc = word.Documents.Open(path, ReadOnly=True)
time.sleep(0.5)

ps = doc.PageSetup
print(f"page={ps.PageWidth}x{ps.PageHeight} top={ps.TopMargin} layoutMode={ps.LayoutMode}")
sec = doc.Sections(1)
print(f"Section LineNumbering active: {sec.PageSetup.LineNumbering.Active}")

p1 = doc.Paragraphs(1)
print(f"P1 text: {p1.Range.Text[:50]!r}")
print(f"P1 font: {p1.Range.Font.Name} sz={p1.Range.Font.Size}")
print(f"P1 LineSpacingRule={p1.LineSpacingRule} LineSpacing={p1.LineSpacing}")
print(f"P1 SpaceBefore={p1.SpaceBefore} SpaceAfter={p1.SpaceAfter}")
print(f"P1 Y (Information6) = {p1.Range.Information(6)}")

p2 = doc.Paragraphs(2)
print(f"\nP2 font: {p2.Range.Font.Name} sz={p2.Range.Font.Size}")
print(f"P2 Y = {p2.Range.Information(6)}")
print(f"P1→P2 = {p2.Range.Information(6) - p1.Range.Information(6)}")

doc.Close(SaveChanges=False)

# Now create test docs with various oversized fonts in linesAndChars mode
print("\n=== Synthetic test: P0 oversized in linesAndChars ===")
print(f"{'font':<14} {'size':<5} {'pitch':<7} {'P0_y':<7} {'P1_y':<7} {'p1h':<6}")

def measure(font, size, pitch_tw):
    doc = word.Documents.Add()
    time.sleep(0.2)
    ps = doc.PageSetup
    ps.PageWidth = 595; ps.PageHeight = 842
    ps.TopMargin = 56.7  # match 1ec
    ps.BottomMargin = 56.7
    ps.LeftMargin = 42.55
    ps.RightMargin = 42.55
    try:
        ps.LayoutMode = 2
    except Exception:
        pass
    rng = doc.Range()
    rng.InsertAfter("テスト\n本文")
    rng = doc.Range()
    rng.Font.Name = font
    # Set sizes per paragraph
    p1 = doc.Paragraphs(1)
    p1.Range.Font.Size = size
    p2 = doc.Paragraphs(2)
    p2.Range.Font.Size = 10.5
    time.sleep(0.1)
    try:
        y1 = p1.Range.Information(6)
        y2 = p2.Range.Information(6)
    except Exception as e:
        doc.Close(SaveChanges=False)
        return None, None, None
    doc.Close(SaveChanges=False)
    return y1, y2, y2 - y1

for font in ["ＭＳ 明朝", "Yu Mincho", "Meiryo"]:
    for size in [10.5, 12, 14, 18, 20, 24]:
        y1, y2, p1h = measure(font, size, 357)
        if y1:
            print(f"{font:<14} {size:<5} 17.85   {y1:<7.2f} {y2:<7.2f} {p1h:<6.2f}")

word.Quit()
