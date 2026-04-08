"""COM measure: LM≥1 (charGrid type=lines or linesAndChars) line height per font/size.

LM=1 (lines): grid snap on lines only
LM=2 (linesAndChars): grid snap on both lines and chars
LM=3 (snapToGrid): genko mode

Sweep with the same fonts/sizes as Round 9 to compare against LM=0.
"""
import win32com.client, time, sys, json, os
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

def measure(font, size, layout_mode, line_pitch_tw=360):
    doc = word.Documents.Add()
    time.sleep(0.2)
    ps = doc.PageSetup
    ps.PageWidth = 612; ps.LeftMargin = 90; ps.RightMargin = 90
    ps.TopMargin = 72; ps.BottomMargin = 72
    try:
        ps.LayoutMode = layout_mode
    except Exception as e:
        print(f"  ERR setup: {e}")
        doc.Close(SaveChanges=False)
        return None
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
    return round(y2 - y1, 3)

print("LM modes line height sweep (linePitch=360tw=18pt)")
print(f"{'mode':<6} {'font':<14} {'size':<5} {'measured':<8}")
results = {}
for lm in [0, 1, 2]:
    label = {0: "LM0", 1: "LM1", 2: "LM2"}[lm]
    results[label] = {}
    for font in ["Times New Roman", "ＭＳ 明朝", "Yu Mincho", "Meiryo"]:
        results[label][font] = {}
        for size in [9, 10, 10.5, 11, 12, 14, 16, 18]:
            dy = measure(font, size, lm)
            results[label][font][size] = dy
            print(f"{label:<6} {font:<14} {size:<5} {dy}")
        print()

os.makedirs("tools/metrics/output", exist_ok=True)
with open("tools/metrics/output/lm_modes_lineheight.json", "w", encoding="utf-8") as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
print("Saved: tools/metrics/output/lm_modes_lineheight.json")

word.Quit()
