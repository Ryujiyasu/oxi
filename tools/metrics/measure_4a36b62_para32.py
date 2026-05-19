"""Direct COM measurement of 4a36b62 paragraph 32 (the 備考 item 2) to verify
the S109c claim of Word fitting 58 chars on line 1."""
import json, os, sys, time
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = os.path.abspath("tools/golden-test/documents/docx/4a36b62555f2_kyodokenkyuyoushiki10.docx")
OUT = os.path.abspath("tools/metrics/4a36b62_para32_word.json")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False
doc = word.Documents.Open(DOC, ReadOnly=True)
time.sleep(0.5)

# Find paragraph that starts with "２" and is the 備考 item 2 (~98 chars)
target = None
for i in range(1, doc.Paragraphs.Count + 1):
    p = doc.Paragraphs(i)
    text = p.Range.Text
    if text.startswith("２") and "本報告書に記入された" in text and len(text) >= 90:
        target = i
        print(f"Found target at paragraph #{i}, text={text!r}")
        break

if target is None:
    print("Target paragraph not found")
    doc.Close(SaveChanges=False); word.Quit(); sys.exit(1)

para = doc.Paragraphs(target)
rng = para.Range
chars = rng.Characters
n = min(chars.Count, 200)
print(f"n_chars={chars.Count}")

per_char = []
for c in range(1, n + 1):
    ch = chars(c)
    try:
        x = ch.Information(5)
        y = ch.Information(6)
    except Exception:
        break
    per_char.append({"c": c, "x": round(x, 3), "y": round(y, 3), "ch": ch.Text})

# Group by y
from collections import defaultdict
lines = defaultdict(list)
for pc in per_char:
    lines[round(pc["y"], 2)].append(pc)
print(f"\nLines detected: {len(lines)}")
for yk in sorted(lines):
    lst = lines[yk]
    text = "".join(pc["ch"] for pc in lst)
    print(f"  y={yk:.2f}  N={len(lst):3d}  x=[{lst[0]['x']:.2f}..{lst[-1]['x']:.2f}]  end={lst[-1]['ch']!r}  text={text[:30]!r}...{text[-12:]!r}")

with open(OUT, "w", encoding="utf-8") as f:
    json.dump({"target_para": target, "chars": per_char}, f, ensure_ascii=False, indent=2)
print(f"\nSaved to {OUT}")

doc.Close(SaveChanges=False)
word.Quit()
