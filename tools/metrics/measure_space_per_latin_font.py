"""Verify if Latin font ownership affects 'space-adjacent-to-CJK' width.

Hypothesis: memory cjk_space_width_spec.md was measured with 游明朝 multi-name list
where Word may use eastAsia for the space too. With Times New Roman explicit Latin font,
the space stays natural Latin width.
"""
import win32com.client, time, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

def measure(text, font_name, size):
    doc = word.Documents.Add()
    time.sleep(0.15)
    rng = doc.Range()
    rng.InsertAfter(text)
    rng = doc.Range()
    rng.Font.Size = size
    rng.Font.Name = font_name
    time.sleep(0.05)
    chars = doc.Range().Characters
    xs = []
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci)
            ch = c.Text
            if ch in ("\r","\x07"):
                continue
            xs.append((ch, c.Information(5), c.Font.Name))
        except Exception:
            continue
    doc.Close(SaveChanges=False)
    return xs

# Compare 5 font families × multi-name vs explicit Latin
TESTS = [
    # (label, font_name, size, text)
    ("游明朝 10.5pt 'A あ'",      "游明朝",                    10.5, "A あ"),
    ("游明朝 10.5pt 'A あ B'",    "游明朝",                    10.5, "A あ B"),
    ("游明朝 12pt 'A あ B'",      "游明朝",                    12.0, "A あ B"),
    ("MS明朝 12pt 'A あ B'",      "ＭＳ 明朝",                  12.0, "A あ B"),
    ("multi 12pt 'A あ B'",       "MS明朝,Times New Roman",    12.0, "A あ B"),
    ("multi 12pt 't 日'",         "MS明朝,Times New Roman",    12.0, "t 日"),
    ("multi 12pt 'test. 日本'",   "MS明朝,Times New Roman",    12.0, "test. 日本"),
    ("multi 12pt '. 日'",         "MS明朝,Times New Roman",    12.0, ". 日"),
    ("multi 11pt '. 日'",         "MS明朝,Times New Roman",    11.0, ". 日"),
    ("游明朝 11pt '. 日'",         "游明朝",                    11.0, ". 日"),
]

for label, font, size, text in TESTS:
    xs = measure(text, font, size)
    print(f"\n[{label}] font={font!r} size={size}")
    for i in range(len(xs)-1):
        adv = round(xs[i+1][1] - xs[i][1], 3)
        print(f"  {xs[i][0]!r:>5} (font={xs[i][2]!r:32}) adv={adv}")
    if xs:
        print(f"  {xs[-1][0]!r:>5} (font={xs[-1][2]!r:32}) <last>")

word.Quit()
