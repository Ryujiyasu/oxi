"""Test if Word compresses CJK punctuation during line breaking.

Hypothesis: Word renders CJK punctuation （） 「」 etc at half-width (50%)
when computing line break positions, even without justify alignment.

Test:
1. Create paragraph with 40 CJK chars + N punctuation chars
2. Vary line width to find break points
3. Compare to: full-width (40+N chars at 11pt = (40+N)*11pt)
                half-width (40 + N*0.5 chars at 11pt)
"""
import win32com.client
import time

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

def measure(text, font="ＭＳ 明朝", size=11.0, page_width_tw=12240, margin_tw=1800):
    """Returns list of (line_y, char_count, line_width_pt, last_char_x)."""
    doc = word.Documents.Add()
    time.sleep(0.2)
    # Set page size and margins
    doc.PageSetup.PageWidth = page_width_tw / 20  # twips → pt
    doc.PageSetup.LeftMargin = margin_tw / 20
    doc.PageSetup.RightMargin = margin_tw / 20

    rng = doc.Range()
    rng.InsertAfter(text)
    rng = doc.Range()
    rng.Font.Name = font
    rng.Font.Size = size
    doc.Paragraphs(1).Alignment = 0  # left
    time.sleep(0.2)

    chars = doc.Range().Characters
    lines = {}  # y -> [(x, ch)]
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci)
            ch = c.Text
            if ch in ("\r", "\x07"):
                continue
            cx = c.Information(5)
            cy = c.Information(6)
            lines.setdefault(round(cy, 2), []).append((cx, ch))
        except Exception:
            continue
    doc.Close(SaveChanges=False)
    out = []
    for y in sorted(lines.keys()):
        chs = lines[y]
        text_line = "".join(c for _, c in chs)
        last_x = max(x for x, _ in chs)
        first_x = min(x for x, _ in chs)
        out.append((y, len(chs), text_line, first_x, last_x))
    return out


# Test 1: pure CJK ideographs (no punct) — baseline
print("=== Test 1: pure CJK (40 ideographs) ===")
text1 = "あ" * 50
for y, n, t, fx, lx in measure(text1):
    print(f"  y={y} chars={n} width={lx-fx:.2f}pt first..last")

# Test 2: CJK with full-width parens — does Word compress?
print("\n=== Test 2: CJK + パンクチュエーション ===")
# 漢字 + ( reading ) pattern x N
text2 = "漢字（かな）" * 8 + "終わり"
for y, n, t, fx, lx in measure(text2):
    print(f"  y={y} chars={n} width={lx-fx:.2f}pt first..last")

# Test 3: CJK with brackets at same positions
print("\n=== Test 3: CJK + brackets ===")
text3 = "「漢字」" * 12 + "終"
for y, n, t, fx, lx in measure(text3):
    print(f"  y={y} chars={n} width={lx-fx:.2f}pt first..last")

# Test 4: measure individual char widths in different contexts
print("\n=== Test 4: char widths ===")
def char_widths(text, font="ＭＳ 明朝", size=11.0):
    doc = word.Documents.Add()
    time.sleep(0.2)
    rng = doc.Range()
    rng.InsertAfter(text)
    rng = doc.Range()
    rng.Font.Name = font
    rng.Font.Size = size
    doc.Paragraphs(1).Alignment = 0
    time.sleep(0.1)
    chars = doc.Range().Characters
    xs = []
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci)
            ch = c.Text
            if ch in ("\r", "\x07"):
                continue
            xs.append((ch, c.Information(5)))
        except:
            continue
    doc.Close(SaveChanges=False)
    widths = []
    for i in range(len(xs) - 1):
        widths.append((xs[i][0], round(xs[i+1][1] - xs[i][1], 4)))
    return widths

# Pure: 漢字
print("  '漢字漢字':", char_widths("漢字漢字"))
# Mixed: 漢（か）字
print("  '漢（か）字':", char_widths("漢（か）字"))
# Pure punct
print("  '（（））':", char_widths("（（））"))
# 漢字「漢字」
print("  '漢字「漢字」漢':", char_widths("漢字「漢字」漢"))

word.Quit()
