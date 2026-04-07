"""Test if Word breaks 'マッチョ' in the middle.

Various padding lengths to force the break point at different chars.
"""
import win32com.client
import time
import sys

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

def test_lines(text, font="ＭＳ 明朝", size=10.5):
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


# The phrase: あなたはマッチョですね。
# Pad with あ before to force break at different positions inside マッチョ
phrase = "あなたはマッチョですね。"

# Try various paddings to position the break at different points
for pad in range(36, 50):
    text = "あ" * pad + phrase
    lines = test_lines(text)
    if len(lines) >= 2:
        l0_tail = lines[0][-6:]
        l1_head = lines[1][:8]
        sys.stdout.buffer.write(
            (f"pad={pad} L0[{len(lines[0])}]: ...{l0_tail!r} | L1[{len(lines[1])}]: {l1_head!r}...\n").encode("utf-8", "replace")
        )
    else:
        sys.stdout.buffer.write(f"pad={pad}: no break\n".encode("utf-8"))

word.Quit()
