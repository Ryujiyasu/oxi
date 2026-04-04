"""COM: Measure character widths for MS Mincho 8pt systematically.

Test many CJK chars to find the pattern of 7.5pt vs 8.0pt widths.
"""
import win32com.client
import os, time

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

path = os.path.abspath("tools/golden-test/documents/docx/459f05f1e877_kyodokenkyuyoushiki01.docx")
doc = word.Documents.Open(path, ReadOnly=False)
time.sleep(1)

# Use P196 which is MS Mincho 8pt, left-align for natural widths
para = doc.Paragraphs(196)
rng = para.Range
original_text = rng.Text

# Switch to left align
para.Alignment = 0
time.sleep(0.1)

# Measure widths of ALL chars in the full paragraph text
chars = rng.Characters
n = chars.Count

widths_7_5 = []
widths_8_0 = []
other_widths = {}

prev_x = None
prev_ch = None
for i in range(1, n + 1):
    try:
        c = chars(i)
        ch = c.Text
        if ch in ('\r', '\x07'):
            continue
        cx = c.Information(5)
        if prev_x is not None and prev_ch is not None:
            w = round(cx - prev_x, 1)
            if w > 0 and w < 20:
                cp = ord(prev_ch)
                if w == 7.5:
                    widths_7_5.append(prev_ch)
                elif w == 8.0:
                    widths_8_0.append(prev_ch)
                else:
                    if w not in other_widths:
                        other_widths[w] = []
                    other_widths[w].append(prev_ch)
        prev_x = cx
        prev_ch = ch
    except:
        pass

# Restore
para.Alignment = 3

print(f"=== MS Mincho 8pt character widths ===")
print(f"7.5pt chars ({len(widths_7_5)}): {sorted(set(widths_7_5))}")
print(f"8.0pt chars ({len(widths_8_0)}): first 20 = {sorted(set(widths_8_0))[:20]}")
for w, chars_list in sorted(other_widths.items()):
    print(f"{w}pt chars ({len(chars_list)}): {sorted(set(chars_list))[:10]}")

# Print unicode codepoints of 7.5pt chars
print(f"\n7.5pt codepoints:")
for ch in sorted(set(widths_7_5)):
    print(f"  U+{ord(ch):04X} '{ch}'")

# Now test with a wider range using a temporary paragraph
# Create test with specific chars
test_chars = "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをん"
test_chars += "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン"
test_chars += "一二三四五六七八九十百千万億兆円年月日時分秒"
test_chars += "記究様添表申出書者税務大学校共同研究際以下通個票"

# Add a new paragraph at end with test chars
rng2 = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
rng2.InsertAfter('\r' + test_chars)
time.sleep(0.1)

# Find the new paragraph and measure
last_para = doc.Paragraphs(doc.Paragraphs.Count)
last_para.Alignment = 0  # left
last_rng = last_para.Range
last_rng.Font.Name = "ＭＳ 明朝"
last_rng.Font.Size = 8
time.sleep(0.1)

chars2 = last_rng.Characters
n2 = chars2.Count

results = {}
prev_x = None
prev_ch = None
for i in range(1, n2 + 1):
    try:
        c = chars2(i)
        ch = c.Text
        if ch in ('\r', '\x07', '\n'):
            continue
        cx = c.Information(5)
        if prev_x is not None and prev_ch is not None:
            w = round(cx - prev_x, 1)
            if 0 < w < 20:
                results[prev_ch] = w
        prev_x = cx
        prev_ch = ch
    except:
        pass

print(f"\n=== Systematic test results ===")
w75 = [ch for ch, w in results.items() if w == 7.5]
w80 = [ch for ch, w in results.items() if w == 8.0]
print(f"7.5pt ({len(w75)}): {''.join(sorted(w75))}")
print(f"8.0pt ({len(w80)}): {''.join(sorted(w80))}")
for ch in sorted(w75):
    print(f"  U+{ord(ch):04X} '{ch}' = 7.5pt")

doc.Close(SaveChanges=False)
word.Quit()
