"""COM measurement: ASCII space (U+0020) width in different contexts.

Hypothesis from session_040107:
- Latin context: 3.0pt
- CJK context (between CJK chars): 5.5-6.0pt
- This may be context-dependent justification or autoSpaceDE behavior.

Test patterns:
  P1: "A A"        — pure Latin
  P2: "あ あ"      — pure CJK with space
  P3: "A あ"       — Latin space CJK
  P4: "あ A"       — CJK space Latin
  P5: "あ A あ"    — CJK Latin CJK with spaces

Font: メイリオ 10.5pt (matches LOD_Handbook)
"""
import win32com.client
import time

def measure(text, font_name="メイリオ", size=10.5):
    word.Documents.Add()
    doc = word.ActiveDocument
    time.sleep(0.2)
    rng = doc.Range()
    rng.InsertAfter(text)
    rng = doc.Range()
    rng.Font.Name = font_name
    rng.Font.Size = size
    doc.Paragraphs(1).Alignment = 0  # left, no justify
    time.sleep(0.1)

    chars = doc.Range().Characters
    out = []
    for i in range(1, chars.Count + 1):
        c = chars(i)
        ch = c.Text
        if ch in ("\r", "\x07"):
            continue
        out.append((ch, c.Information(5)))  # X position in pt
    doc.Close(SaveChanges=False)

    # Compute widths between adjacent chars
    widths = []
    for i in range(len(out) - 1):
        ch = out[i][0]
        w = round(out[i + 1][1] - out[i][1], 4)
        widths.append((ch, w))
    return widths


word = win32com.client.Dispatch("Word.Application")
word.Visible = False

patterns = [
    ("P1 Latin",       "A A"),
    ("P2 CJK",         "\u3042 \u3042"),
    ("P3 L_C",         "A \u3042"),
    ("P4 C_L",         "\u3042 A"),
    ("P5 C_L_C",       "\u3042 A \u3042"),
    ("P6 L_C_L",       "A \u3042 A"),
    ("P7 multiL",      "A  A"),       # double space
    ("P8 CJK_punct",   "\u3042\u3000\u3042"),  # ideographic space
]

print(f"{'pattern':<14} {'sequence':<30} widths")
print("-" * 80)
for name, text in patterns:
    widths = measure(text)
    seq = " ".join(f"'{ch}'" for ch, _ in widths)
    wstr = " ".join(f"{w}" for _, w in widths)
    print(f"{name:<14} {seq:<30} {wstr}")

word.Quit()
