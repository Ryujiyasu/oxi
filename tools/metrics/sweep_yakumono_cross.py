"""Test yakumono cross-type adjacency: close+open, punct+open, etc."""
import win32com.client
import time
import json
import sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

def advances(text, font="ＭＳ 明朝", size=11.0):
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
        except Exception:
            continue
    doc.Close(SaveChanges=False)
    return [(xs[i][0], round(xs[i+1][1] - xs[i][1], 4)) for i in range(len(xs) - 1)]

# Cross-type tests
TESTS = [
    "漢」（漢",       # close+open between CJK
    "漢、（漢",       # punct+open
    "漢」、漢",       # close+punct
    "漢。「漢",       # punct+open
    "漢「」漢",       # open+close (empty quote)
    "漢、。漢",       # punct+punct
    "漢）。漢",       # close+punct
    "漢（）漢",       # paren pair
    "漢、！漢",       # punct + non-yakumono punct (!)
    "漢！、漢",       # ! + comma — does ! get compressed?
    "漢！？漢",       # ! + ? both type-C
    "漢」！漢",       # close + ! (! is type C)
    "漢」漢",         # single close between CJK
    "漢（漢",         # single open between CJK
]

for t in TESTS:
    a = advances(t)
    print(f"{t}: {a}")

word.Quit()
