"""Verify: are P196 chars truly 7.5pt in LEFT alignment?
Measure existing paragraph vs fresh insertion."""
import win32com.client
import os, time

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

# Test 1: Fresh document with same text, MS Mincho 8pt
doc = word.Documents.Add()
time.sleep(0.5)

test_text = "５　①から⑨の記載及び研究計画書（様式３）の添付は、代表者になっている申出者の申出書に行うこととして、その他"
rng = doc.Range()
rng.InsertAfter(test_text)
rng = doc.Range()
rng.Font.Name = "ＭＳ 明朝"
rng.Font.Size = 8
doc.Paragraphs(1).Alignment = 0  # left
time.sleep(0.1)

chars = doc.Range().Characters
prev_x = None
prev_ch = None
widths = []
for i in range(1, chars.Count + 1):
    c = chars(i)
    ch = c.Text
    if ch in ('\r', '\x07'):
        continue
    cx = c.Information(5)
    if prev_x is not None:
        w = round(cx - prev_x, 2)
        widths.append((prev_ch, w))
    prev_x = cx
    prev_ch = ch

print("=== Fresh doc, MS Mincho 8pt, LEFT ===")
for ch, w in widths:
    marker = "***" if abs(w - 8.0) > 0.1 else ""
    if marker:
        print(f"  '{ch}' U+{ord(ch):04X}: {w}pt {marker}")

# Count
w75 = sum(1 for _, w in widths if abs(w - 7.5) < 0.1)
w80 = sum(1 for _, w in widths if abs(w - 8.0) < 0.1)
print(f"\n7.5pt: {w75}, 8.0pt: {w80}, other: {len(widths) - w75 - w80}")

doc.Close(SaveChanges=False)

# Test 2: Original document P196
path = os.path.abspath("tools/golden-test/documents/docx/459f05f1e877_kyodokenkyuyoushiki01.docx")
doc2 = word.Documents.Open(path, ReadOnly=False)
time.sleep(1)

para = doc2.Paragraphs(196)
para.Alignment = 0  # left
time.sleep(0.1)

chars2 = para.Range.Characters
prev_x = None
prev_ch = None
widths2 = []
for i in range(1, chars2.Count + 1):
    c = chars2(i)
    ch = c.Text
    if ch in ('\r', '\x07'):
        continue
    cx = c.Information(5)
    if prev_x is not None:
        w = round(cx - prev_x, 2)
        widths2.append((prev_ch, w))
    prev_x = cx
    prev_ch = ch

print("\n=== Original doc P196, MS Mincho 8pt, LEFT ===")
for ch, w in widths2[:60]:
    marker = "***" if abs(w - 8.0) > 0.1 else ""
    if marker:
        print(f"  '{ch}' U+{ord(ch):04X}: {w}pt {marker}")

w75 = sum(1 for _, w in widths2[:60] if abs(w - 7.5) < 0.1)
w80 = sum(1 for _, w in widths2[:60] if abs(w - 8.0) < 0.1)
print(f"\n7.5pt: {w75}, 8.0pt: {w80}, other: {len(widths2[:60]) - w75 - w80}")

# Check: is charGrid active?
ps = doc2.Sections(1).PageSetup
print(f"\nCharsLine: {ps.CharsLine}, LayoutMode: {ps.LayoutMode}")

# Check if docGrid type=lines affects char spacing
print(f"Grid active? LineSpacing={para.Format.LineSpacing:.1f}")

doc2.Close(SaveChanges=False)
word.Quit()
