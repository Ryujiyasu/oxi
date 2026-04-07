"""Test if yakumono compression is font-specific or universal."""
import win32com.client
import time
import sys

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

def widths(text, font, size):
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


# Test "漢）」漢" pattern across fonts and sizes
test_strs = [
    "漢）」漢",   # closer pair
    "（（漢",     # opener pair
    "漢、。漢",   # comma-period
    "漢：「漢",   # colon + opener
]

fonts = [
    ("ＭＳ 明朝", 11.0),
    ("ＭＳ 明朝", 14.0),
    ("游ゴシック", 11.0),
    ("游ゴシック", 14.0),
    ("游明朝", 11.0),
    ("游明朝", 14.0),
    ("メイリオ", 11.0),
    ("メイリオ", 14.0),
    ("ＭＳ Ｐゴシック", 11.0),
    ("MS Gothic", 11.0),
]

for font, sz in fonts:
    sys.stdout.buffer.write(f"\n=== {font} {sz}pt ===\n".encode("utf-8", "replace"))
    for t in test_strs:
        try:
            w = widths(t, font, sz)
            s = "  ".join(f"{ch}={ww}" for ch, ww in w)
            sys.stdout.buffer.write((repr(t) + ": " + s + "\n").encode("utf-8", "replace"))
        except Exception as e:
            sys.stdout.buffer.write((repr(t) + f" ERR: {e}\n").encode("utf-8", "replace"))

word.Quit()
