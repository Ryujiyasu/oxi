"""Minimal verify: MS Mincho 10.5pt single line height in LM0 (lineRule=auto).

Creates a 3-paragraph doc with MS Mincho 10.5pt, default line spacing (lineRule=auto).
Measures P2_y - P1_y and P3_y - P2_y to compare with lm0_lineauto.json (12.0pt).
Memory claim: Word P1→P2 = 14.0pt, P2→P3 = 13.5pt (under-estimated by 1.5pt).
"""
import win32com.client, time, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

doc = word.Documents.Add()
time.sleep(0.3)
ps = doc.PageSetup
ps.PageWidth = 595.3  # A4
ps.PageHeight = 841.9
ps.LeftMargin = 99.25; ps.RightMargin = 99.25  # ~35mm
ps.TopMargin = 113.4; ps.BottomMargin = 113.4  # ~40mm
try:
    ps.LayoutMode = 0  # wdLayoutModeDefault — LM0
except Exception as e:
    print(f"LayoutMode set fail: {e}")

print(f"LayoutMode={ps.LayoutMode} CharsLine={ps.CharsLine}")

rng = doc.Range()
rng.InsertAfter("あいうえお\nかきくけこ\nさしすせそ")
for p in [doc.Paragraphs(1), doc.Paragraphs(2), doc.Paragraphs(3)]:
    p.Range.Font.Name = "ＭＳ 明朝"
    p.Range.Font.Size = 10.5
    p.LineSpacingRule = 0  # wdLineSpaceSingle (lineRule=auto semantics)
    p.SpaceBefore = 0; p.SpaceAfter = 0

time.sleep(0.3)
y1 = doc.Paragraphs(1).Range.Information(6)
y2 = doc.Paragraphs(2).Range.Information(6)
y3 = doc.Paragraphs(3).Range.Information(6)
print(f"P1_y={y1:.3f}  P2_y={y2:.3f}  P3_y={y3:.3f}")
print(f"gap P1→P2 = {y2-y1:.3f}pt")
print(f"gap P2→P3 = {y3-y2:.3f}pt")
print(f"lm0_lineauto.json says: 12.0pt")

doc.Close(SaveChanges=False)
word.Quit()
