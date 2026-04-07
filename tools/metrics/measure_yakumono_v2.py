"""V2: confirm rules for ：；？！ and various punct combos."""
import win32com.client
import time

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

def widths(text):
    doc = word.Documents.Add()
    time.sleep(0.15)
    rng = doc.Range()
    rng.InsertAfter(text)
    rng = doc.Range()
    rng.Font.Name = "ＭＳ 明朝"
    rng.Font.Size = 11.0
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
    # ：；？！with neighbors
    "漢：）漢", "漢）：漢",
    "漢；）漢", "漢）；漢",
    "漢？「漢", "漢「？漢",
    "漢！」漢", "漢」！漢",
    # 、。 with brackets
    "漢、）漢", "漢）、漢",
    "漢。「漢", "漢「。漢",
    # interpunct ・
    "漢・「漢", "漢」・漢",
    # 全角ハイフン ー
    "漢ー）漢",
    # mid-dot ・
    "漢・・漢",
    # ） before line-end (no next char)
    "漢、",
    # Different brackets
    "漢〕〔漢", "漢】【漢", "漢》《漢",
]

for t in tests:
    w = widths(t)
    s = "  ".join(f"{ch}={ww}" for ch, ww in w)
    import sys
    sys.stdout.buffer.write((repr(t) + ": " + s + "\n").encode("utf-8", errors="replace"))

word.Quit()
