"""V2: Force the break to fall INSIDE マッチョ.

Use ASCII padding (half-width) so the break point lands inside the phrase.
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


# Phrases to test
phrases = [
    "あなたはマッチョですね。",   # マッチョ
    "私たちは、元気です。",       # comma + period + small kana
]

for phrase in phrases:
    sys.stdout.buffer.write(f"\n=== '{phrase}' ===\n".encode("utf-8", "replace"))
    # Vary CJK padding so the phrase straddles the break
    # Each CJK char ~10.5pt; line fits ~40 chars
    for pad in range(28, 42):
        text = "漢" * pad + phrase
        lines = test_lines(text)
        if len(lines) >= 2:
            l0 = lines[0]
            l1 = lines[1]
            sys.stdout.buffer.write(
                (f"  pad={pad:2d} L0[{len(l0)}]={l0[-12:]!r} | L1[{len(l1)}]={l1[:14]!r}\n").encode("utf-8", "replace")
            )

word.Quit()
