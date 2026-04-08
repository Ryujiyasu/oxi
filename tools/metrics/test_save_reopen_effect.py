"""Save runtime doc to disk, reopen, compare space adv."""
import win32com.client, time, os, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

TEXT = "これはMS明朝フォントのテストです。This is Times New Roman font test. 日本語と英語が混在している行です。"

def find_sp_adv(doc):
    chars = doc.Range().Characters
    xs = []
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci)
            ch = c.Text
            if ch in ("\r","\x07"):
                continue
            xs.append((ch, c.Information(5)))
        except Exception:
            continue
    for i in range(len(xs)-1):
        if xs[i][0] == ' ' and xs[i+1][0] == '日':
            return round(xs[i+1][1] - xs[i][1], 3)
    return None

# Create runtime doc, measure
doc = word.Documents.Add()
time.sleep(0.4)
ps = doc.PageSetup
ps.PageWidth = 612; ps.LeftMargin = 90; ps.RightMargin = 90
rng = doc.Range()
rng.InsertAfter(TEXT)
rng = doc.Range()
rng.Font.Size = 12.0
rng.Font.Name = "MS明朝,Times New Roman"
doc.Paragraphs(1).Alignment = 0
time.sleep(0.3)
sp1 = find_sp_adv(doc)
print(f"Runtime (before save): sp_adv = {sp1}")

# Save to disk as new file
out_path = "C:\\Users\\ryuji\\oxi-1\\tools\\metrics\\output\\test_runtime_jfmb.docx"
os.makedirs(os.path.dirname(out_path), exist_ok=True)
doc.SaveAs2(out_path, FileFormat=12)  # docx
time.sleep(0.4)
sp2 = find_sp_adv(doc)
print(f"Runtime (after save): sp_adv = {sp2}")
doc.Close(SaveChanges=False)

# Reopen the saved file
time.sleep(0.4)
doc2 = word.Documents.Open(out_path, ReadOnly=True)
time.sleep(0.4)
sp3 = find_sp_adv(doc2)
print(f"Reopened: sp_adv = {sp3}")
doc2.Close(SaveChanges=False)

word.Quit()
