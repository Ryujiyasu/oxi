"""COM measure CJK-adjacent Latin space width at 12pt for fonts in
japanese_font_mixing_baseline (MS明朝/TNR mixed run)."""
import win32com.client
import time
import sys

sys.stdout.reconfigure(encoding="utf-8")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False


def measure(text, font_name, size):
    doc = word.Documents.Add()
    time.sleep(0.1)
    ps = doc.PageSetup
    ps.PageWidth = 612.0; ps.PageHeight = 792.0
    ps.LeftMargin = 90.0; ps.RightMargin = 90.0
    ps.TopMargin = 72.0; ps.BottomMargin = 72.0
    rng = doc.Range()
    rng.InsertAfter(text)
    rng = doc.Range()
    rng.Font.Name = font_name
    rng.Font.Size = size
    doc.Paragraphs(1).Alignment = 0
    time.sleep(0.05)
    chars = doc.Range().Characters
    out = []
    for i in range(1, chars.Count + 1):
        try:
            c = chars(i); ch = c.Text
            if ch in ("\r","\x07"): continue
            out.append((ch, c.Information(5)))
        except: pass
    doc.Close(False)
    widths = [(out[i][0], round(out[i+1][1] - out[i][1], 4)) for i in range(len(out)-1)]
    return widths


def report(label, text, font, sz):
    ws = measure(text, font, sz)
    line = "  ".join(f"{ch!r}={w}" for ch, w in ws)
    print(f"{label:35s} {font:18s} {sz:5.1f}pt: {line}")


# 12pt tests
print("=== 12pt CJK-adjacent space ===")
report("A A (Latin context)",   "A A B B",      "Times New Roman", 12)
report("A あ (CJK adj)",         "A あ A",        "Times New Roman", 12)
report("あ A (CJK adj)",         "あ A あ",       "Times New Roman", 12)
report("MS明朝 A A",             "A A B B",      "ＭＳ 明朝", 12)
report("MS明朝 A あ",            "A あ A",        "ＭＳ 明朝", 12)
report("Calibri A A",            "A A B B",      "Calibri", 12)
report("Calibri A あ",           "A あ A",        "Calibri", 12)

print("\n=== 12pt actual problem text (mixed run) ===")
# The actual problematic text
TEXT = "これはMS明朝フォントのテストです。This is Times New Roman font test. 日本語と英語が混在している行です。"
ws = measure(TEXT, "Times New Roman", 12)
print(f"Total chars: {len(ws)+1}")
# Just show space-related transitions
for i, (ch, w) in enumerate(ws):
    if ch == ' ':
        prev_ch = ws[i-1][0] if i > 0 else '?'
        nxt_ch = ws[i+1][0] if i+1 < len(ws) else '?'
        print(f"  [{i:2d}] {prev_ch!r}→' '→? width={w}")

word.Quit()
