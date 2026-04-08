"""Round 26: Extended font sweep for LM2 first-paragraph 0.25pt rounding direction.

Round 24 found Latin floors / CJK ceils the 0.25 residual in
  P0_y = topMargin + (grid_n*pitch - LM0_lh) / 2

Hypothesis: Word uses font-class-specific rounding of LM0_lh in the centering
numerator (Latin → ceil, CJK → floor). Test with additional fonts:
  - HG series (excluded from CJK 83/64 whitelist per §1.2)
  - Latin Times (different from Times New Roman)
  - Garamond / Constantia (Latin)
  - SimSun / SimHei (Chinese)
  - Malgun Gothic (Korean)
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

    # Also need LM0 to compute the predicted offset
    rng = doc.Range()
    rng.InsertAfter("ABC\nDEF")
    rng = doc.Range()
    rng.Font.Size = size
    rng.Font.Name = font
    time.sleep(0.1)
    try:
        y1_lm2 = doc.Paragraphs(1).Range.Information(6)
        y2_lm2 = doc.Paragraphs(2).Range.Information(6)
        p0h_lm2 = y2_lm2 - y1_lm2
    except Exception:
        doc.Close(SaveChanges=False); return None
    doc.Close(SaveChanges=False)

    # LM0 measurement
    doc = word.Documents.Add()
    time.sleep(0.15)
    ps = doc.PageSetup
    ps.PageWidth = 612; ps.LeftMargin = 90; ps.RightMargin = 90
    ps.TopMargin = 72; ps.BottomMargin = 72
    try: ps.LayoutMode = 0
    except: pass
    rng = doc.Range()
    rng.InsertAfter("ABC\nDEF")
    rng = doc.Range()
    rng.Font.Size = size
    rng.Font.Name = font
    time.sleep(0.1)
    try:
        y1_lm0 = doc.Paragraphs(1).Range.Information(6)
        y2_lm0 = doc.Paragraphs(2).Range.Information(6)
        lm0_lh = y2_lm0 - y1_lm0
    except Exception:
        doc.Close(SaveChanges=False); return None
    doc.Close(SaveChanges=False)
    return (y1_lm2, p0h_lm2, lm0_lh)

print(f"{'class':<7} {'font':<22} {'size':<5} {'LM0_lh':<7} {'P0_h':<6} {'offset':<7} {'raw':<7} {'delta':<6}")

# Each tuple: (class, font_name)
TESTS = [
    ("Latin", "Times"),  # If installed (different from Times New Roman)
    ("Latin", "Garamond"),
    ("Latin", "Constantia"),
    ("Latin", "Georgia"),
    ("Latin", "Verdana"),
    ("CJK", "HGS明朝E"),  # HG series — excluded from CJK 83/64
    ("CJK", "HGP明朝E"),
    ("CJK", "HGSｺﾞｼｯｸM"),
    ("CJK", "SimSun"),  # zh-CN
    ("CJK", "SimHei"),
    ("CJK", "Malgun Gothic"),  # ko
    ("CJK", "Batang"),
]
SIZES = [10.5, 11, 12, 13, 14, 16, 17, 18, 19, 20, 22, 24]

for cls, font in TESTS:
    any_ok = False
    for size in SIZES:
        r = measure(font, size)
        if r is None:
            continue
        y1, p0h, lm0 = r
        offset = y1 - 72.0
        raw = (p0h - lm0) / 2.0
        delta = offset - raw
        any_ok = True
        print(f"{cls:<7} {font:<22} {size:<5} {lm0:<7.2f} {p0h:<6.2f} {offset:<7.2f} {raw:<7.3f} {delta:<6.2f}")
    if not any_ok:
        print(f"{cls:<7} {font:<22} (font not available)")
    print()

word.Quit()
