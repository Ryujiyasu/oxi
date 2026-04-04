"""COM: Measure justify absorption threshold.

When a paragraph is justify-aligned and the next-line's first char can fit
by compressing inter-character spacing, Word absorbs it. Measure the exact
threshold by creating test cases with controlled overflow amounts.
"""
import win32com.client
import os, time

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

doc = word.Documents.Add()
time.sleep(0.5)

# Set narrow margins for controlled width
ps = doc.Sections(1).PageSetup
ps.LeftMargin = 72  # 1 inch
ps.RightMargin = 72
# content_w = 595.3 - 72 - 72 = 451.3pt

content_w = ps.PageWidth - ps.LeftMargin - ps.RightMargin
print(f"Content width: {content_w:.1f}pt")

# MS Mincho 10.5pt: each fullwidth char = 10.5pt
# content_w / 10.5 = 43.0 chars per line
chars_per_line = int(content_w / 10.5)
print(f"Chars per line (natural): {chars_per_line}")

# Test: insert paragraphs with increasing char counts
# At chars_per_line+1, does justify compress to fit on 1 line?
results = []
for extra in range(0, 8):
    n = chars_per_line + extra
    text = "あ" * n  # all same-width chars

    # Add paragraph
    if doc.Paragraphs.Count > 1 or doc.Range().Text.strip():
        doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertAfter('\r')
    rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    rng.InsertAfter(text)

    # Set font and justify
    para = doc.Paragraphs(doc.Paragraphs.Count)
    para.Range.Font.Name = "ＭＳ 明朝"
    para.Range.Font.Size = 10.5
    para.Alignment = 3  # justify

time.sleep(0.3)

# Measure line counts
for pi in range(1, doc.Paragraphs.Count + 1):
    para = doc.Paragraphs(pi)
    rng = para.Range
    text = rng.Text.rstrip('\r')
    if not text:
        continue

    n = len(text)
    overflow_pt = (n - chars_per_line) * 10.5

    # Count lines by Y position
    chars = rng.Characters
    first_y = chars(1).Information(6)
    last_y = chars(min(n, chars.Count)).Information(6)
    is_multiline = abs(last_y - first_y) > 1.0

    # Also check with left alignment
    para.Alignment = 0
    time.sleep(0.05)
    left_last_y = chars(min(n, chars.Count)).Information(6)
    left_multiline = abs(left_last_y - first_y) > 1.0
    para.Alignment = 3
    time.sleep(0.05)

    absorbed = left_multiline and not is_multiline
    print(f"  n={n} (+{n-chars_per_line}) overflow={overflow_pt:.1f}pt: "
          f"left={'2+lines' if left_multiline else '1line'} "
          f"justify={'2+lines' if is_multiline else '1line'} "
          f"{'ABSORBED' if absorbed else ''}")

# More precise test with 8pt font
print(f"\n=== MS Mincho 8pt test ===")
doc2 = word.Documents.Add()
time.sleep(0.5)
ps2 = doc2.Sections(1).PageSetup
ps2.LeftMargin = 72
ps2.RightMargin = 72
content_w2 = ps2.PageWidth - ps2.LeftMargin - ps2.RightMargin
chars_per_line_8 = int(content_w2 / 8.0)
print(f"Content width: {content_w2:.1f}pt, natural chars/line: {chars_per_line_8}")

for extra in range(0, 8):
    n = chars_per_line_8 + extra
    text = "あ" * n

    if doc2.Paragraphs.Count > 1 or doc2.Range().Text.strip():
        doc2.Range(doc2.Content.End - 1, doc2.Content.End - 1).InsertAfter('\r')
    rng = doc2.Range(doc2.Content.End - 1, doc2.Content.End - 1)
    rng.InsertAfter(text)

    para = doc2.Paragraphs(doc2.Paragraphs.Count)
    para.Range.Font.Name = "ＭＳ 明朝"
    para.Range.Font.Size = 8
    para.Alignment = 3

time.sleep(0.3)

for pi in range(1, doc2.Paragraphs.Count + 1):
    para = doc2.Paragraphs(pi)
    rng = para.Range
    text = rng.Text.rstrip('\r')
    if not text:
        continue

    n = len(text)
    overflow_pt = (n - chars_per_line_8) * 8.0

    chars = rng.Characters
    first_y = chars(1).Information(6)
    last_y = chars(min(n, chars.Count)).Information(6)
    is_multiline = abs(last_y - first_y) > 1.0

    para.Alignment = 0
    time.sleep(0.05)
    left_last_y = chars(min(n, chars.Count)).Information(6)
    left_multiline = abs(left_last_y - first_y) > 1.0
    para.Alignment = 3
    time.sleep(0.05)

    absorbed = left_multiline and not is_multiline
    print(f"  n={n} (+{n-chars_per_line_8}) overflow={overflow_pt:.1f}pt: "
          f"left={'2+lines' if left_multiline else '1line'} "
          f"justify={'2+lines' if is_multiline else '1line'} "
          f"{'ABSORBED' if absorbed else ''}")

doc2.Close(SaveChanges=False)
doc.Close(SaveChanges=False)
word.Quit()
