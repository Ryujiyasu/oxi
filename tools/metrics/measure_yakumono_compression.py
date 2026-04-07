"""Detailed test of yakumono (CJK punctuation) compression rules.

Based on initial finding: consecutive CJK punctuation chars compress 2nd+ to 50%.
"""
import win32com.client
import time

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

def widths(text, font="ＭＳ 明朝", size=11.0):
    doc = word.Documents.Add()
    time.sleep(0.15)
    rng = doc.Range()
    rng.InsertAfter(text)
    rng = doc.Range()
    rng.Font.Name = font
    rng.Font.Size = size
    doc.Paragraphs(1).Alignment = 0
    time.sleep(0.05)
    chars = doc.Range().Characters
    xs = []
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci)
            ch = c.Text
            if ch in ("\r", "\x07"):
                continue
            xs.append((ch, c.Information(5)))
        except:
            continue
    doc.Close(SaveChanges=False)
    out = []
    for i in range(len(xs) - 1):
        out.append((xs[i][0], round(xs[i+1][1] - xs[i][1], 4)))
    return out


tests = [
    # Pairs of openers
    "（（漢", "「「漢", "（「漢",
    # Pairs of closers
    "漢））", "漢」」", "漢）」",
    # Mixed: closer then opener
    "漢）（漢", "漢」「漢",
    # Comma/period combinations
    "漢、。漢", "漢。、漢", "漢、、漢", "漢。。漢",
    # The actual ruby_text_lineheight_11 pattern
    "漢字（かな）」「漢", # ）」「 sequence
    "漢）」漢", "漢」）漢",
    # ASCII colon
    "漢：漢", "漢：「漢",
    # Triple
    "漢）」）漢",
    # Punct after small kana
    "漢っ）漢", "漢ょ）漢",
    # Numeric
    "漢）１漢",
    # Different punct widths in MS Mincho mono
    "漢、漢", "漢。漢",
]

for t in tests:
    w = widths(t)
    s = "  ".join(f"{ch}={ww}" for ch, ww in w)
    import sys
    sys.stdout.buffer.write((repr(t) + ": " + s + "\n").encode("utf-8", errors="replace"))

word.Quit()
