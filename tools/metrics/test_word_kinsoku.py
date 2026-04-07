"""Test what chars Word actually treats as line-start-prohibited."""
import win32com.client
import time
import sys

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

def test_lines(text):
    """Render text and see where Word breaks."""
    doc = word.Documents.Add()
    time.sleep(0.2)
    rng = doc.Range()
    rng.InsertAfter(text)
    rng = doc.Range()
    rng.Font.Name = "ＭＳ 明朝"
    rng.Font.Size = 10.5
    doc.Paragraphs(1).Alignment = 0
    time.sleep(0.1)
    chars = doc.Range().Characters
    prev_y = None
    line_chars = []
    out_lines = []
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci)
            ch = c.Text
            if ch in ("\r", "\x07"):
                continue
            cy = c.Information(6)
            if prev_y is not None and abs(cy - prev_y) > 0.5:
                out_lines.append("".join(line_chars))
                line_chars = []
            prev_y = cy
            line_chars.append(ch)
        except:
            continue
    if line_chars:
        out_lines.append("".join(line_chars))
    doc.Close(SaveChanges=False)
    return out_lines


# Setup: 41 chars of CJK then a small kana — does Word allow it at line start?
# We need a long enough text to force a break

# Test small kana characters
test_chars = ['ョ', 'ャ', 'ュ', 'ぁ', 'ぃ', 'っ', 'ヮ', 'ー', 'ヽ', '・', '。', '）', '」']
# Pad with regular kanji to force line break around char 41
for tc in test_chars:
    # 40 kanji + tc + more chars
    text = "あ" * 40 + tc + "い" * 20
    lines = test_lines(text)
    if len(lines) >= 2:
        l1_first = lines[1][0] if lines[1] else "?"
        marker = "→ ALLOWED at line start" if l1_first == tc else f"→ blocked (L1 starts with {l1_first!r})"
    else:
        marker = "→ no break"
    sys.stdout.buffer.write(
        (f"'{tc}' (U+{ord(tc):04X}): " + marker + "\n").encode("utf-8", "replace")
    )

word.Quit()
