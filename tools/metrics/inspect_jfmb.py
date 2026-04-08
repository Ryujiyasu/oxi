"""COM-measure japanese_font_mixing_baseline — multi-name font list."""
import win32com.client, time, os, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

path = os.path.abspath("pipeline_data/docx/japanese_font_mixing_baseline.docx")
doc = word.Documents.Open(path, ReadOnly=True)
time.sleep(0.4)

ps = doc.PageSetup
print(f"body width = {ps.PageWidth - ps.LeftMargin - ps.RightMargin:.2f}pt  (page={ps.PageWidth} L={ps.LeftMargin} R={ps.RightMargin})")

para = doc.Paragraphs(1)
chars = para.Range.Characters
rows = []
for ci in range(1, chars.Count + 1):
    try:
        c = chars(ci)
        ch = c.Text
        if ch in ("\r", "\x07"):
            continue
        rows.append({"i": ci, "ch": ch, "x": c.Information(5), "y": c.Information(6),
                     "font": c.Font.Name, "size": c.Font.Size,
                     "lang": c.LanguageID, "langFE": c.LanguageIDFarEast})
    except Exception:
        continue

lines = {}
for r in rows:
    lines.setdefault(round(r["y"], 1), []).append(r)

for y in sorted(lines.keys()):
    ln = lines[y]
    first_x = min(r["x"] for r in ln)
    last_x = max(r["x"] for r in ln)
    text = "".join(r["ch"] for r in ln)
    fonts = set((r["font"], r["size"]) for r in ln)
    print(f"\ny={y} chars={len(ln)} first_x={first_x:.2f} last_x={last_x:.2f} fonts={fonts}")
    print(f"  text: {text!r}")
    ln_sorted = sorted(ln, key=lambda r: r["x"])
    for i in range(len(ln_sorted)):
        ch = ln_sorted[i]["ch"]
        fn = ln_sorted[i]["font"]
        x = ln_sorted[i]["x"]
        if i + 1 < len(ln_sorted):
            adv = round(ln_sorted[i+1]["x"] - x, 3)
        else:
            adv = "(LAST)"
        print(f"    {i:3d}: {ch!r:>6} font={fn!r:25} lang={ln_sorted[i]['lang']} langFE={ln_sorted[i]['langFE']} x={x:.2f} adv={adv}")

doc.Close(SaveChanges=False)
word.Quit()
