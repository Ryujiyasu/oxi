"""Bisect: when does ' ' before CJK get half-em vs Latin natural width?"""
import win32com.client, time, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

def measure_para(text, font, size, align=0, page_w=612, lm=90, rm=90):
    doc = word.Documents.Add()
    time.sleep(0.15)
    ps = doc.PageSetup
    ps.PageWidth = page_w
    ps.LeftMargin = lm
    ps.RightMargin = rm
    rng = doc.Range()
    rng.InsertAfter(text)
    rng = doc.Range()
    rng.Font.Size = size
    rng.Font.Name = font
    doc.Paragraphs(1).Alignment = align
    time.sleep(0.05)
    chars = doc.Range().Characters
    xs = []
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci)
            ch = c.Text
            if ch in ("\r","\x07"):
                continue
            xs.append((ch, c.Information(5), c.Information(6), c.Font.Name))
        except Exception:
            continue
    doc.Close(SaveChanges=False)
    return xs

# Test progression — find the threshold
TESTS = [
    "test. 日",
    "font test. 日",
    "Roman font test. 日",
    "New Roman font test. 日",
    "Times New Roman font test. 日",
    "is Times New Roman font test. 日",
    "This is Times New Roman font test. 日",
    "。This is Times New Roman font test. 日",
    "です。This is Times New Roman font test. 日",
    # Full jfmb-like
    "これはMS明朝フォントのテストです。This is Times New Roman font test. 日",
]

font = "MS明朝,Times New Roman"
size = 12.0
for text in TESTS:
    xs = measure_para(text, font, size)
    # Find ' ' before '日'
    sp_adv = None
    sp_font = None
    日_font = None
    for i in range(len(xs)-1):
        if xs[i][0] == ' ' and xs[i+1][0] == '日':
            sp_adv = round(xs[i+1][1] - xs[i][1], 3)
            sp_font = xs[i][3]
            日_font = xs[i+1][3]
            break
    print(f"len={len(xs):2d}  sp_adv={sp_adv}  sp_font={sp_font!r:30}  日_font={日_font!r:20}  text={text!r}")

word.Quit()
