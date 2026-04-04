"""COM: Measure empty paragraph height for Meiryo."""
import win32com.client
import time

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Add()
time.sleep(0.5)

# Set font
rng = doc.Range()
rng.Font.Name = "メイリオ"
rng.Font.Size = 10.5

# Insert: text, empty, text, empty, text
texts = ["テスト1", "", "テスト2", "", "テスト3"]
for i, t in enumerate(texts):
    if i > 0:
        doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertAfter('\r')
    if t:
        doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertAfter(t)

# Set all paragraphs to same font
for pi in range(1, doc.Paragraphs.Count + 1):
    p = doc.Paragraphs(pi)
    p.Range.Font.Name = "メイリオ"
    p.Range.Font.Size = 10.5
    p.Format.SpaceBefore = 0
    p.Format.SpaceAfter = 0
    p.Alignment = 0

time.sleep(0.2)

print("Meiryo 10.5pt: text and empty paragraph heights")
prev_y = None
for pi in range(1, doc.Paragraphs.Count + 1):
    p = doc.Paragraphs(pi)
    rng = p.Range
    text = rng.Text.rstrip('\r')
    y = rng.Information(6)
    is_empty = len(text) == 0

    if prev_y is not None:
        gap = y - prev_y
        print(f"  P{pi}: y={y:.2f} gap={gap:.2f}pt {'(empty)' if is_empty else text}")
    else:
        print(f"  P{pi}: y={y:.2f} {'(empty)' if is_empty else text}")
    prev_y = y

# Also test MS Mincho for comparison
print("\nMS Mincho 10.5pt:")
doc2 = word.Documents.Add()
time.sleep(0.5)
for i, t in enumerate(texts):
    if i > 0:
        doc2.Range(doc2.Content.End - 1, doc2.Content.End - 1).InsertAfter('\r')
    if t:
        doc2.Range(doc2.Content.End - 1, doc2.Content.End - 1).InsertAfter(t)

for pi in range(1, doc2.Paragraphs.Count + 1):
    p = doc2.Paragraphs(pi)
    p.Range.Font.Name = "ＭＳ 明朝"
    p.Range.Font.Size = 10.5
    p.Format.SpaceBefore = 0
    p.Format.SpaceAfter = 0
    p.Alignment = 0

time.sleep(0.2)

prev_y = None
for pi in range(1, doc2.Paragraphs.Count + 1):
    p = doc2.Paragraphs(pi)
    rng = p.Range
    text = rng.Text.rstrip('\r')
    y = rng.Information(6)
    is_empty = len(text) == 0

    if prev_y is not None:
        gap = y - prev_y
        print(f"  P{pi}: y={y:.2f} gap={gap:.2f}pt {'(empty)' if is_empty else text}")
    else:
        print(f"  P{pi}: y={y:.2f} {'(empty)' if is_empty else text}")
    prev_y = y

doc.Close(SaveChanges=False)
doc2.Close(SaveChanges=False)
word.Quit()
