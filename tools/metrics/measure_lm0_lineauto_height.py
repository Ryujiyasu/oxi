"""COM measure: LM0 (no docGrid type) + lineRule=auto line height per font/size/multiplier.

Method: Create 2-paragraph doc with various settings, measure P2_y - P1_y = effective line+after height.
For lineRule=auto, w:line is in 240ths of 'natural single-spaced' line height.

Sweep:
- Font: Times New Roman, MS明朝, MS Gothic, Yu Mincho
- Size: 10, 10.5, 11, 12, 14
- line: 240(=1.0x), 276(=1.15x default), 360(=1.5x), 480(=2.0x)
- after: 0, 200(=10pt)
"""
import win32com.client, time, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

def measure_para_height(font, size, line_240ths, after_tw):
    """Returns dy = (P2_y - P1_y), the effective height of P1 (1 line + after)."""
    doc = word.Documents.Add()
    time.sleep(0.2)
    ps = doc.PageSetup
    ps.PageWidth = 612; ps.LeftMargin = 90; ps.RightMargin = 90
    ps.TopMargin = 72; ps.BottomMargin = 72
    try:
        ps.LayoutMode = 0  # wdLayoutModeDefault — no grid
    except Exception:
        pass
    if not hasattr(measure_para_height, "_logged"):
        try:
            print(f"  [diag] LayoutMode={ps.LayoutMode} CharsLine={ps.CharsLine} LinesPage={ps.LinesPage}")
        except Exception as e:
            print(f"  [diag] err: {e}")
        measure_para_height._logged = True
    # Insert two short paragraphs
    rng = doc.Range()
    rng.InsertAfter("ABC\nDEF")
    rng = doc.Range()
    rng.Font.Size = size
    rng.Font.Name = font
    # Set line spacing on both paras
    for p in [doc.Paragraphs(1), doc.Paragraphs(2)]:
        p.LineSpacingRule = 5  # wdLineSpaceMultiple - we'll use wdLineSpaceAtLeast=3, wdLineSpaceMultiple=5
        # Actually for "auto" we use wdLineSpaceSingle=0 and SpaceAfter
        # But we want to test custom line= multiplier, so use wdLineSpaceMultiple
        p.LineSpacing = size * (line_240ths / 240.0)
        # Hmm actually wdLineSpaceMultiple expects pt value = size * multiplier
        p.SpaceAfter = after_tw / 20.0  # tw → pt
    time.sleep(0.1)
    # Read Y of P1 and P2
    try:
        y1 = doc.Paragraphs(1).Range.Information(6)
        y2 = doc.Paragraphs(2).Range.Information(6)
    except Exception as e:
        doc.Close(SaveChanges=False)
        return None
    doc.Close(SaveChanges=False)
    return round(y2 - y1, 3)

print("LM0 + lineRule=auto sweep (no docGrid type, no grid snap)")
print(f"{'font':<20} {'size':<6} {'line':<5} {'after':<6} {'measured':<10}")

import json
results = {}
for font in ["Times New Roman", "ＭＳ 明朝", "Yu Mincho", "Yu Gothic", "Calibri", "Meiryo"]:
    results[font] = {}
    for size_x10 in range(70, 251, 5):  # 7.0, 7.5, ..., 25.0
        size = size_x10 / 10.0
        try:
            dy = measure_para_height(font, size, 240, 0)
        except Exception as e:
            dy = None
        ppem_round = round(size * 96 / 72)
        results[font][size] = dy
        print(f"{font:<18} size={size:<5} ppem={ppem_round:<3} measured={dy}")
import os
os.makedirs("tools/metrics/output", exist_ok=True)
with open("tools/metrics/output/lm0_lineauto_full.json", "w", encoding="utf-8") as f:
    json.dump(results, f, ensure_ascii=False, indent=2)

word.Quit()
