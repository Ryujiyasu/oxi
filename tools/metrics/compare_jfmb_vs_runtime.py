"""Compare on-disk jfmb vs runtime-recreated. Find which setting matters."""
import win32com.client, time, os, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

TEXT = "これはMS明朝フォントのテストです。This is Times New Roman font test. 日本語と英語が混在している行です。"

def find_space_before_japan(doc):
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
    for i in range(len(xs)-1):
        if xs[i][0] == ' ' and xs[i+1][0] == '日':
            return round(xs[i+1][1] - xs[i][1], 3), xs[i][2], xs[i+1][2]
    return None, None, None

# 1. On-disk jfmb
path = os.path.abspath("pipeline_data/docx/japanese_font_mixing_baseline.docx")
doc = word.Documents.Open(path, ReadOnly=True)
time.sleep(0.4)
sp, sf, df = find_space_before_japan(doc)
compat = doc.CompatibilityMode
print(f"[on-disk jfmb] compat={compat} sp_adv={sp} sp_font={sf!r} 日_font={df!r}")
doc.Close(SaveChanges=False)

# 2. Runtime new doc, same text/font/page
doc = word.Documents.Add()
time.sleep(0.3)
ps = doc.PageSetup
ps.PageWidth = 612
ps.PageHeight = 792
ps.LeftMargin = 90
ps.RightMargin = 90
ps.TopMargin = 72
ps.BottomMargin = 72
rng = doc.Range()
rng.InsertAfter(TEXT)
rng = doc.Range()
rng.Font.Size = 12.0
rng.Font.Name = "MS明朝,Times New Roman"
doc.Paragraphs(1).Alignment = 0
time.sleep(0.2)
sp, sf, df = find_space_before_japan(doc)
compat = doc.CompatibilityMode
print(f"[runtime add]  compat={compat} sp_adv={sp} sp_font={sf!r} 日_font={df!r}")
doc.Close(SaveChanges=False)

# 3. Runtime with compat=14
doc = word.Documents.Add()
time.sleep(0.3)
doc.SetCompatibilityMode(14)
ps = doc.PageSetup
ps.PageWidth = 612
ps.LeftMargin = 90
ps.RightMargin = 90
rng = doc.Range()
rng.InsertAfter(TEXT)
rng = doc.Range()
rng.Font.Size = 12.0
rng.Font.Name = "MS明朝,Times New Roman"
doc.Paragraphs(1).Alignment = 0
time.sleep(0.2)
sp, sf, df = find_space_before_japan(doc)
print(f"[runtime cm14] compat={doc.CompatibilityMode} sp_adv={sp} sp_font={sf!r} 日_font={df!r}")
doc.Close(SaveChanges=False)

word.Quit()
