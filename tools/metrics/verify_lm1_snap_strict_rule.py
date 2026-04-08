"""Verify LM1 snap rule:
1. Measure natural_lh in LM0 with EXPLICIT line=240 (true single)
2. Measure same font/size in LM1 (default linePitch=18pt) with EXPLICIT line=240
3. Verify (LM1 result) == (floor(natural/pitch) + 1) * pitch
"""
import win32com.client, time, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

def measure(font, size, layout_mode, line_240ths=240):
    doc = word.Documents.Add()
    time.sleep(0.2)
    ps = doc.PageSetup
    ps.PageWidth = 612; ps.LeftMargin = 90; ps.RightMargin = 90
    ps.TopMargin = 72; ps.BottomMargin = 72
    ps.LayoutMode = layout_mode
    rng = doc.Range()
    rng.InsertAfter("ABC\nDEF")
    rng = doc.Range()
    rng.Font.Size = size
    rng.Font.Name = font
    for p in [doc.Paragraphs(1), doc.Paragraphs(2)]:
        p.LineSpacingRule = 5  # wdLineSpaceMultiple
        p.LineSpacing = size * (line_240ths / 240.0)
        p.SpaceAfter = 0
    time.sleep(0.1)
    try:
        y1 = doc.Paragraphs(1).Range.Information(6)
        y2 = doc.Paragraphs(2).Range.Information(6)
    except Exception:
        doc.Close(SaveChanges=False)
        return None
    doc.Close(SaveChanges=False)
    return round(y2 - y1, 3)

# Verify rule across multiple fonts and sizes
SAMPLES = [
    ("Times New Roman", 12), ("Times New Roman", 14), ("Times New Roman", 16),
    ("ＭＳ 明朝", 11), ("ＭＳ 明朝", 12), ("ＭＳ 明朝", 14), ("ＭＳ 明朝", 18),
    ("Yu Mincho", 10), ("Yu Mincho", 11), ("Yu Mincho", 12),
    ("Meiryo", 9), ("Meiryo", 10), ("Meiryo", 12),
]
PITCH = 18.0

print(f"{'font':<16} {'size':<5} {'LM0_nat':<9} {'LM1_snap':<9} {'rule_pred':<10} {'match':<6}")
for font, size in SAMPLES:
    nat = measure(font, size, 0, 240)
    snap = measure(font, size, 1, 240)
    if nat is None or snap is None:
        print(f"{font:<16} {size:<5} ERR")
        continue
    pred_strict = (int(nat / PITCH) + 1) * PITCH  # strict greater rule
    pred_ceil = ((int(nat / PITCH) + (1 if nat > int(nat/PITCH)*PITCH else 0)) or 1) * PITCH  # ceil with min 1
    if abs(nat - int(nat/PITCH)*PITCH) < 1e-6:
        pred_ceil_v = (int(nat / PITCH)) * PITCH if int(nat/PITCH) > 0 else PITCH
    else:
        pred_ceil_v = (int(nat / PITCH) + 1) * PITCH
    match_strict = "✓" if abs(snap - pred_strict) < 0.1 else "✗"
    print(f"{font:<16} {size:<5} {nat:<9} {snap:<9} {pred_strict:<10} {match_strict}")

word.Quit()
