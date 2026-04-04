"""Quick COM measurement: Meiryo 10.5pt CJK char widths."""
import win32com.client
import time

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Add()
time.sleep(0.5)

text = "本資料は自治体等の機関が所有している統計データをLOD形式で公開"
rng = doc.Range()
rng.InsertAfter(text)
rng = doc.Range()
rng.Font.Name = "メイリオ"
rng.Font.Size = 10.5
doc.Paragraphs(1).Alignment = 0  # left
time.sleep(0.1)

chars = doc.Range().Characters
prev_x = None
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

for ch, w in widths:
    is_ascii = ord(ch) < 128
    print(f"  '{ch}' U+{ord(ch):04X}: {w}pt {'(ASCII)' if is_ascii else ''}")

# Stats
cjk_widths = [w for ch, w in widths if ord(ch) >= 0x3000]
ascii_widths = [w for ch, w in widths if ord(ch) < 128]
print(f"\nCJK avg: {sum(cjk_widths)/len(cjk_widths):.3f}pt")
if ascii_widths:
    print(f"ASCII avg: {sum(ascii_widths)/len(ascii_widths):.3f}pt")

doc.Close(SaveChanges=False)
word.Quit()
