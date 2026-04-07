"""COM measurement V2: isolate space width drivers.

Tests:
  A) autoSpaceDE on/off across contexts
  B) different fonts (Meiryo, Yu Gothic, MS Gothic, Calibri)
  C) charGrid setting on/off
  D) check if Word collapses multiple spaces
"""
import win32com.client
import time

def measure(text, font_name="メイリオ", size=10.5, autospace_de=True):
    word.Documents.Add()
    doc = word.ActiveDocument
    time.sleep(0.2)
    rng = doc.Range()
    rng.InsertAfter(text)
    rng = doc.Range()
    rng.Font.Name = font_name
    rng.Font.Size = size
    para = doc.Paragraphs(1)
    para.Alignment = 0
    # ParagraphFormat properties
    try:
        para.Format.AutoAdjustRightIndent = False
    except: pass
    try:
        # autoSpaceDE = AutoSpaceLikeWord97? Actually it's via Format
        para.Format.AutoSpaceDE = autospace_de
        para.Format.AutoSpaceDN = autospace_de
    except Exception as e:
        pass
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
    return widths, [ch for ch,_ in out]


word = win32com.client.Dispatch("Word.Application")
word.Visible = False

patterns = [
    ("L_L  ",  "A A"),
    ("LLL  ",  "A  A"),    # 2 spaces
    ("CCC  ",  "\u3042 \u3042"),
    ("LCL  ",  "A\u3042A"),       # no space
    ("L_C  ",  "A \u3042"),
    ("C_L  ",  "\u3042 A"),
    ("L_C_L", "A \u3042 A"),
    ("C_L_C", "\u3042 A \u3042"),
]

for font in ("\u30e1\u30a4\u30ea\u30aa", "Yu Gothic", "MS Gothic", "Calibri"):
    print(f"\n=== Font: {font} 10.5pt ===")
    print(f"{'name':<8} {'autoDE':<6} {'sequence':<28} widths")
    for name, text in patterns:
        for autode in (True, False):
            try:
                widths, seq = measure(text, font_name=font, autospace_de=autode)
                seq_str = "".join(seq)
                wstr = " ".join(f"{w:.2f}" for _,w in widths)
                print(f"{name:<8} {str(autode):<6} {repr(seq_str):<28} {wstr}")
            except Exception as e:
                print(f"{name} {autode} ERROR: {e}")

word.Quit()
