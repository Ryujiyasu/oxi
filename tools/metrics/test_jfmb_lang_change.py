"""Open jfmb writable, change langFE to Japanese, re-measure space before 日."""
import win32com.client, time, os, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

WD_JAPANESE = 1041

path = os.path.abspath("pipeline_data/docx/japanese_font_mixing_baseline.docx")

def measure(doc, label):
    chars = doc.Range().Characters
    xs = []
    for ci in range(1, chars.Count + 1):
        try:
            c = chars(ci)
            ch = c.Text
            if ch in ("\r","\x07"):
                continue
            xs.append((ch, c.Information(5), c.LanguageID, c.LanguageIDFarEast))
        except Exception:
            continue
    # Find ' ' before '日'
    for i in range(len(xs)-1):
        if xs[i][0] == ' ' and xs[i+1][0] == '日':
            adv = round(xs[i+1][1] - xs[i][1], 3)
            print(f"{label}: ' '→'日' adv={adv}, lang={xs[i][2]} langFE={xs[i][3]}")
            return
    # Find 'は'→'M' for comparison
    for i in range(len(xs)-1):
        if xs[i][0] == 'は':
            adv = round(xs[i+1][1] - xs[i][1], 3)
            print(f"{label}: 'は' adv={adv}, lang={xs[i][2]} langFE={xs[i][3]}")
            break

# Original
doc = word.Documents.Open(path, ReadOnly=False)
time.sleep(2.0)
measure(doc, "ORIGINAL")
print(f"  Doc lang/langFE before: {doc.Range().LanguageID}/{doc.Range().LanguageIDFarEast}")

# Change langFE to Japanese
doc.Range().LanguageIDFarEast = WD_JAPANESE
doc.Range().LanguageID = WD_JAPANESE
time.sleep(0.3)
measure(doc, "AFTER langFE=1041")

# Find は too
chars = doc.Range().Characters
xs = [(chars(ci).Text, chars(ci).Information(5)) for ci in range(1, chars.Count+1)]
for i in range(len(xs)-1):
    if xs[i][0] == 'は':
        adv = round(xs[i+1][1] - xs[i][1], 3)
        print(f"  After: 'は' adv={adv}")
        break
for i in range(len(xs)-1):
    if xs[i][0] == ' ' and xs[i+1][0] == '日':
        adv = round(xs[i+1][1] - xs[i][1], 3)
        print(f"  After: ' '→'日' adv={adv}")
        break

doc.Close(SaveChanges=False)
word.Quit()
