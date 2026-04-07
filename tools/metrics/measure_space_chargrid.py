"""Test if document charGrid setting affects space widths.

Word's PageSetup.LayoutMode controls character grid:
  0 = wdLayoutModeDefault (no grid)
  1 = wdLayoutModeGrid
  2 = wdLayoutModeLineGrid
  3 = wdLayoutModeGenko (manuscript)
"""
import win32com.client
import time

def measure(text, font_name, size, layout_mode):
    word.Documents.Add()
    doc = word.ActiveDocument
    time.sleep(0.2)
    # Set layout mode
    try:
        doc.PageSetup.LayoutMode = layout_mode
    except Exception as e:
        return None, str(e)

    rng = doc.Range()
    rng.InsertAfter(text)
    rng = doc.Range()
    rng.Font.Name = font_name
    rng.Font.Size = size
    doc.Paragraphs(1).Alignment = 0
    time.sleep(0.1)

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
    return widths, None


word = win32com.client.Dispatch("Word.Application")
word.Visible = False

patterns = [
    ("L_L  ",  "A A"),
    ("LLL  ",  "A  A"),
    ("L_C  ",  "A \u3042"),
    ("L_C_L", "A \u3042 A"),
]

modes = [
    (0, "Default(no grid)"),
    (1, "Grid"),
    (2, "LineGrid"),
    (3, "Genko"),
]

for font in ("Yu Gothic", "Calibri"):
    print(f"\n=== {font} 10.5pt ===")
    for mode_id, mode_name in modes:
        print(f"\n--- LayoutMode={mode_id} ({mode_name}) ---")
        for name, text in patterns:
            widths, err = measure(text, font, 10.5, mode_id)
            if err:
                print(f"  {name}: ERROR {err}")
                continue
            wstr = " ".join(f"{w:.2f}" for _,w in widths)
            print(f"  {name}: {wstr}")

word.Quit()
