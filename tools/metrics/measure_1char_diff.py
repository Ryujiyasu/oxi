"""COM: Measure exact character widths and available width for ±1 char/line difference.

kyodokenkyuyoushiki01 P196: Word 55ch vs Oxi 54ch on line 1.
Measure per-character X positions to determine actual char widths.
"""
import win32com.client
import os, time

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

path = os.path.abspath("tools/golden-test/documents/docx/459f05f1e877_kyodokenkyuyoushiki01.docx")
doc = word.Documents.Open(path, ReadOnly=True)
time.sleep(1)

para = doc.Paragraphs(196)
rng = para.Range
text = rng.Text.rstrip('\r')
print(f"P196: {len(text)} chars total")
print(f"Font: {rng.Font.Name}, Size: {rng.Font.Size}")
print(f"Alignment: {para.Alignment}")  # 3=justify

# Paragraph format
fmt = para.Format
print(f"LeftIndent: {fmt.LeftIndent:.2f}")
print(f"RightIndent: {fmt.RightIndent:.2f}")
print(f"FirstLineIndent: {fmt.FirstLineIndent:.2f}")

# Page setup
ps = doc.Sections(1).PageSetup
content_w = ps.PageWidth - ps.LeftMargin - ps.RightMargin
print(f"ContentWidth: {content_w:.2f}")

# First line available width
first_line_left = fmt.LeftIndent + fmt.FirstLineIndent
avail_first = content_w - first_line_left - fmt.RightIndent
print(f"FirstLineAvail: {avail_first:.2f} (left={first_line_left:.2f})")

# Measure char positions on line 1
chars = rng.Characters
n = chars.Count
line1_y = chars(1).Information(6)

print(f"\nLine 1 chars (y={line1_y:.1f}):")
char_data = []
for i in range(1, min(n+1, 70)):
    c = chars(i)
    ch = c.Text
    if ch in ('\r', '\x07'):
        continue
    cy = c.Information(6)
    if abs(cy - line1_y) > 1.0:
        print(f"  -- line break at char {i} (y={cy:.1f}) --")
        break
    cx = c.Information(5)
    char_data.append((i, ch, cx, cy))

# Calculate per-character widths from positions
print(f"\nTotal line 1 chars: {len(char_data)}")
if len(char_data) >= 2:
    first_x = char_data[0][2]
    last_x = char_data[-1][2]
    span = last_x - first_x
    n_gaps = len(char_data) - 1
    avg_gap = span / n_gaps if n_gaps > 0 else 0
    print(f"Span: {first_x:.2f} to {last_x:.2f} = {span:.2f}pt")
    print(f"Avg char spacing: {avg_gap:.4f}pt ({n_gaps} gaps)")

    # Show individual widths for first 20 chars
    print(f"\nPer-char widths (first 20):")
    for j in range(min(20, n_gaps)):
        w = char_data[j+1][2] - char_data[j][2]
        print(f"  [{j+1}] '{char_data[j][1]}' width={w:.3f}pt")

    # Check if uniform or varied
    widths = [char_data[j+1][2] - char_data[j][2] for j in range(n_gaps)]
    unique_w = set(round(w, 2) for w in widths)
    print(f"\nUnique widths: {sorted(unique_w)}")

# Now switch to LEFT alignment and check
print(f"\n=== LEFT alignment test ===")
# Need read-write access
doc.Close(SaveChanges=False)
doc = word.Documents.Open(os.path.abspath(path), ReadOnly=False)
time.sleep(0.5)
para = doc.Paragraphs(196)
rng = para.Range

# Switch to left
para.Alignment = 0
time.sleep(0.1)

chars = rng.Characters
line1_y = chars(1).Information(6)
left_char_data = []
for i in range(1, min(n+1, 70)):
    c = chars(i)
    ch = c.Text
    if ch in ('\r', '\x07'):
        continue
    cy = c.Information(6)
    if abs(cy - line1_y) > 1.0:
        print(f"  Left align: line break at char {i}, total line1={len(left_char_data)}ch")
        break
    cx = c.Information(5)
    left_char_data.append((i, ch, cx, cy))

if len(left_char_data) >= 2:
    first_x = left_char_data[0][2]
    last_x = left_char_data[-1][2]
    span = last_x - first_x
    n_gaps = len(left_char_data) - 1
    avg_gap = span / n_gaps if n_gaps > 0 else 0
    print(f"Left: {len(left_char_data)}ch on line 1")
    print(f"Left span: {first_x:.2f} to {last_x:.2f} = {span:.2f}pt")
    print(f"Left avg spacing: {avg_gap:.4f}pt")

    # Per-char widths (natural, no justify)
    left_widths = [left_char_data[j+1][2] - left_char_data[j][2] for j in range(n_gaps)]
    unique_lw = set(round(w, 2) for w in left_widths)
    print(f"Left unique widths: {sorted(unique_lw)}")

    # Total natural text width
    natural_w = span + left_widths[-1] if left_widths else span
    print(f"Natural text width (line1): {natural_w:.2f}pt")

# Restore
para.Alignment = 3
doc.Close(SaveChanges=False)
word.Quit()
