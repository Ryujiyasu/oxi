"""Extended space measurement: many fonts, find the rule for CJK-adjacent space width.

Pattern: "A A" (Latin-only) vs "A あ" (CJK-adjacent)
Goal: derive a formula for the increment.
"""
import win32com.client
import time

def measure(text, font_name, size):
    word.Documents.Add()
    doc = word.ActiveDocument
    time.sleep(0.15)
    rng = doc.Range()
    rng.InsertAfter(text)
    rng = doc.Range()
    rng.Font.Name = font_name
    rng.Font.Size = size
    doc.Paragraphs(1).Alignment = 0
    time.sleep(0.05)

    chars = doc.Range().Characters
    out = []
    for i in range(1, chars.Count + 1):
        c = chars(i)
        ch = c.Text
        if ch in ("\r", "\x07"):
            continue
        out.append((ch, c.Information(5)))
    doc.Close(SaveChanges=False)

    widths = []
    for i in range(len(out) - 1):
        widths.append((out[i][0], round(out[i+1][1] - out[i][1], 4)))
    return widths


word = win32com.client.Dispatch("Word.Application")
word.Visible = False

fonts = [
    "Calibri",
    "Times New Roman",
    "Arial",
    "Cambria",
    "Century",
    "Yu Gothic",
    "Yu Mincho",
    "\u30e1\u30a4\u30ea\u30aa",  # Meiryo
    "MS Gothic",
    "MS Mincho",
    "MS PGothic",
    "MS PMincho",
]

print(f"{'font':<22} {'size':<5} {'L_sp':<6} {'C_sp':<6} {'incr':<6} {'A_w':<6} {'CJK_w':<6}")
print("-" * 65)
for font in fonts:
    for size in (10.5, 11):
        try:
            ll = measure("A A", font, size)         # widths: A_width, space_width
            lc = measure("A \u3042", font, size)    # widths: A_width, space_width
            cc = measure("\u3042\u3042", font, size)  # widths: CJK_width
            l_sp = ll[1][1]
            c_sp = lc[1][1]
            a_w = ll[0][1]
            cjk_w = cc[0][1]
            print(f"{font:<22} {size:<5} {l_sp:<6.2f} {c_sp:<6.2f} {c_sp-l_sp:<+6.2f} {a_w:<6.2f} {cjk_w:<6.2f}")
        except Exception as e:
            print(f"{font:<22} {size}: ERROR {e}")

word.Quit()
