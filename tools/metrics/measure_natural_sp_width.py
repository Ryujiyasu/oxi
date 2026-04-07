"""Measure natural (unstretched) Latin space width adjacent to CJK by using
short paragraphs that don't reach the right margin (so no justification).
"""
import win32com.client
import os, sys, tempfile, zipfile, re, time

sys.stdout.reconfigure(encoding="utf-8")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False


def measure(text, font="游明朝", size=10.5):
    """Build short doc, measure per-char x positions."""
    doc = word.Documents.Add()
    time.sleep(0.1)
    ps = doc.PageSetup
    ps.PageWidth = 612.0; ps.PageHeight = 792.0
    ps.LeftMargin = 90.0; ps.RightMargin = 90.0
    ps.TopMargin = 72.0; ps.BottomMargin = 72.0
    rng = doc.Range()
    rng.InsertAfter(text)
    rng = doc.Range()
    rng.Font.Name = font
    rng.Font.Size = size
    doc.Paragraphs(1).Alignment = 0
    time.sleep(0.05)
    chars = doc.Range().Characters
    out = []
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci); ch = c.Text
            if ch in ("\r","\x07"): continue
            x = c.Information(5)
            out.append((ch, x))
        except: pass
    doc.Close(False)
    widths = [(out[i][0], round(out[i+1][1] - out[i][1], 4)) for i in range(len(out)-1)]
    return widths


def report(label, text, font, size):
    ws = measure(text, font, size)
    line = "  ".join(f"{ch!r}={w}" for ch, w in ws)
    print(f"{label:35s} {font:10s} {size}pt: {line}")


# All tests use very short text so no justification kicks in.
# Goal: isolate the natural width of SP in various CJK contexts.

print("=== 游明朝 10.5pt natural SP widths ===")
report("Latin only A B",      "A B",     "游明朝", 10.5)
report("Latin word A x B",    "AB CD",   "游明朝", 10.5)
report("CJK 1 + L 1",         "あ A",    "游明朝", 10.5)
report("L 1 + CJK 1",         "A あ",    "游明朝", 10.5)
report("CJK 1 + SP + L word", "あ Az",   "游明朝", 10.5)
report("L word + SP + CJK",   "Az あ",   "游明朝", 10.5)
report("L word + SP + CJK 1 + SP + L word", "Az あ Cd", "游明朝", 10.5)
report("Long L + SP + CJK + SP + L", "Microsoft の Azure", "游明朝", 10.5)
# What if we add many CJK before/after to stress test
report("L + SP + CJK CJK + SP + L", "Az あい Cd", "游明朝", 10.5)
report("CJK CJK + SP + L + SP + CJK CJK", "あい Az かき", "游明朝", 10.5)

print("\n=== Times New Roman 12pt natural ===")
report("Long L + SP + CJK + SP + L", "Microsoft の Azure", "Times New Roman", 12)
report("L word + SP + CJK", "Az あ", "Times New Roman", 12)

word.Quit()
