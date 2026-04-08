"""Round 24: Pin down 0.25pt rounding direction in LM2 first-paragraph centering.

Previous Round 23 found residual 0.25pt mismatches in (grid_n - LM0_lh)/2 formula:
  - Latin TNR 18pt/24pt: predicted 7.75/4.25 → measured 7.5/4.0 (floor)
  - CJK MS Mincho/Yu Mincho/Meiryo: predicted 6.25 → measured 6.5 (ceil)

Hypothesis: Latin floors 0.25, CJK ceils 0.25.
Alternative: COM 0.5pt quantization noise.

Test by sweeping more (font, size) combinations that produce *.25 fractional offsets.
"""
import win32com.client, time, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

def measure(font, size):
    doc = word.Documents.Add()
    time.sleep(0.15)
    ps = doc.PageSetup
    ps.PageWidth = 612; ps.LeftMargin = 90; ps.RightMargin = 90
    ps.TopMargin = 72; ps.BottomMargin = 72
    try: ps.LayoutMode = 2  # linesAndChars
    except: pass
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
        doc.Close(SaveChanges=False); return None
    doc.Close(SaveChanges=False)
    return (y1, y2 - y1)

print(f"{'font':<18} {'size':<6} {'P0_y':<7} {'P0_h':<7} {'offset':<8} {'cls':<6}")
LATIN = ["Times New Roman", "Calibri", "Arial", "Cambria", "Century"]
CJK = ["ＭＳ 明朝", "ＭＳ ゴシック", "Yu Mincho", "Yu Gothic", "Meiryo"]
SIZES = [10.5, 11, 11.5, 12, 13, 13.5, 14, 15, 16, 17, 18, 19, 20, 21, 22, 24, 26, 28]

for cls, fonts in [("Latin", LATIN), ("CJK", CJK)]:
    for font in fonts:
        for size in SIZES:
            r = measure(font, size)
            if r is None: continue
            y1, p0h = r
            offset = y1 - 72.0
            print(f"{font:<18} {size:<6} {y1:<7.2f} {p0h:<7.2f} {offset:<8.2f} {cls:<6}")
        print()

word.Quit()
