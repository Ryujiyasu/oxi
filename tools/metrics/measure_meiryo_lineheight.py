"""COM: Measure Meiryo line height at various sizes."""
import win32com.client
import time

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Add()
time.sleep(0.5)

# Insert paragraphs at different sizes
sizes = [8, 9, 10, 10.5, 11, 12, 14, 16]
for fs in sizes:
    rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    rng.InsertAfter("あ\r")
    para = doc.Paragraphs(doc.Paragraphs.Count - 1)
    para.Range.Font.Name = "メイリオ"
    para.Range.Font.Size = fs
    para.Alignment = 0

# Add a second line for reference
rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
rng.InsertAfter("い\r")
doc.Paragraphs(doc.Paragraphs.Count - 1).Range.Font.Name = "メイリオ"
doc.Paragraphs(doc.Paragraphs.Count - 1).Range.Font.Size = 10.5

time.sleep(0.3)

# Measure Y positions
print("Meiryo line heights (Y diff between consecutive paragraphs):")
prev_y = None
for pi in range(1, doc.Paragraphs.Count + 1):
    para = doc.Paragraphs(pi)
    rng = para.Range
    text = rng.Text.rstrip('\r')
    if not text:
        continue
    y = rng.Information(6)
    fs = rng.Font.Size
    ls = para.Format.LineSpacing

    if prev_y is not None:
        gap = y - prev_y
        print(f"  size={prev_fs:.1f}pt: Y gap={gap:.2f}pt (lineSpacing={prev_ls:.2f})")

    prev_y = y
    prev_fs = fs
    prev_ls = ls

# Also measure with explicit lineSpacing settings
print("\nWith exact lineSpacing=20pt:")
doc2 = word.Documents.Add()
time.sleep(0.5)
for i in range(3):
    rng = doc2.Range(doc2.Content.End - 1, doc2.Content.End - 1)
    rng.InsertAfter(f"テスト{i}\r")
    para = doc2.Paragraphs(doc2.Paragraphs.Count - 1)
    para.Range.Font.Name = "メイリオ"
    para.Range.Font.Size = 10.5
    para.Format.LineSpacingRule = 4  # wdLineSpaceExactly
    para.Format.LineSpacing = 20
    para.Format.SpaceBefore = 0
    para.Format.SpaceAfter = 0
    para.Alignment = 0
time.sleep(0.1)

prev_y = None
for pi in range(1, doc2.Paragraphs.Count + 1):
    para = doc2.Paragraphs(pi)
    rng = para.Range
    text = rng.Text.rstrip('\r')
    if not text:
        continue
    y = rng.Information(6)
    if prev_y is not None:
        print(f"  gap={y - prev_y:.2f}pt")
    prev_y = y

# Default single spacing
print("\nDefault single spacing (Meiryo 10.5pt):")
doc3 = word.Documents.Add()
time.sleep(0.5)
for i in range(3):
    rng = doc3.Range(doc3.Content.End - 1, doc3.Content.End - 1)
    rng.InsertAfter(f"テスト行{i}\r")
    para = doc3.Paragraphs(doc3.Paragraphs.Count - 1)
    para.Range.Font.Name = "メイリオ"
    para.Range.Font.Size = 10.5
    para.Format.SpaceBefore = 0
    para.Format.SpaceAfter = 0
    para.Alignment = 0
time.sleep(0.1)

prev_y = None
for pi in range(1, doc3.Paragraphs.Count + 1):
    para = doc3.Paragraphs(pi)
    rng = para.Range
    text = rng.Text.rstrip('\r')
    if not text:
        continue
    y = rng.Information(6)
    ls = para.Format.LineSpacing
    lsr = para.Format.LineSpacingRule
    if prev_y is not None:
        print(f"  gap={y - prev_y:.2f}pt lineSpacing={ls:.2f} rule={lsr}")
    prev_y = y

doc.Close(SaveChanges=False)
doc2.Close(SaveChanges=False)
doc3.Close(SaveChanges=False)
word.Quit()
