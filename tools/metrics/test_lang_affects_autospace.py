"""Test if w:lang.eastAsia=en-US suppresses CJK-adjacent space widening."""
import win32com.client, time, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

WD_JAPANESE = 1041
WD_ENGLISH_US = 1033

def measure(text, font, size, lang_eastasia=None):
    doc = word.Documents.Add()
    time.sleep(0.2)
    ps = doc.PageSetup
    ps.PageWidth = 612; ps.LeftMargin = 90; ps.RightMargin = 90
    rng = doc.Range()
    rng.InsertAfter(text)
    rng = doc.Range()
    rng.Font.Size = size
    rng.Font.Name = font
    if lang_eastasia is not None:
        rng.LanguageIDFarEast = lang_eastasia
        rng.LanguageID = lang_eastasia
    doc.Paragraphs(1).Alignment = 0
    time.sleep(0.1)
    chars = doc.Range().Characters
    xs = []
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci)
            ch = c.Text
            if ch in ("\r","\x07"):
                continue
            xs.append((ch, c.Information(5), c.Font.Name, c.LanguageID, c.LanguageIDFarEast))
        except Exception:
            continue
    doc.Close(SaveChanges=False)
    return xs

TEXTS = [
    "test. 日本",
    "Mは",
    "はM",
]

font = "MS明朝,Times New Roman"
for lang_label, lang in [("default(ja)", None), ("english", WD_ENGLISH_US), ("japanese", WD_JAPANESE)]:
    print(f"\n=== lang={lang_label} ===")
    for text in TEXTS:
        xs = measure(text, font, 12.0, lang)
        print(f"  text={text!r}")
        for i in range(len(xs)):
            adv = round(xs[i+1][1] - xs[i][1], 3) if i+1 < len(xs) else "(last)"
            print(f"    {i}: {xs[i][0]!r} font={xs[i][2]!r} lang={xs[i][3]} langFE={xs[i][4]} adv={adv}")

word.Quit()
