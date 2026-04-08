"""Verify hypothesis: autoSpaceDE behavior differs between same-run and separate-run boundaries."""
import win32com.client, time, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

def measure(doc):
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
    return xs

# Test 1: single multi-name run "Mは" 12pt
print("=== Test 1: single run multi-name 'Mは' 12pt ===")
doc = word.Documents.Add()
time.sleep(0.2)
rng = doc.Range()
rng.InsertAfter("Mは")
rng = doc.Range()
rng.Font.Size = 12.0
rng.Font.Name = "MS明朝,Times New Roman"
time.sleep(0.1)
xs = measure(doc)
for i in range(len(xs)):
    adv = round(xs[i+1][1] - xs[i][1], 3) if i+1 < len(xs) else "(last)"
    print(f"  {i}: {xs[i][0]!r} font={xs[i][2]!r} adv={adv}")
doc.Close(SaveChanges=False)

# Test 2: two separate runs "M" + "は" with explicit fonts
print("\n=== Test 2: two runs explicit fonts (M:TNR, は:ＭＳ明朝) 12pt ===")
doc = word.Documents.Add()
time.sleep(0.2)
# Insert M in Times New Roman
rng = doc.Range()
rng.InsertAfter("M")
rng = doc.Range()
rng.Font.Name = "Times New Roman"
rng.Font.Size = 12.0
# Append は in ＭＳ 明朝
rng = doc.Range()
rng.Collapse(0)  # collapse to end
rng.InsertAfter("は")
# Set font of just the last char
chars = doc.Range().Characters
last = chars(chars.Count)
last_minus = chars(chars.Count - 1)  # the は (since \r is last)
# Actually let's grab the は specifically
for ci in range(1, chars.Count+1):
    c = chars(ci)
    if c.Text == "は":
        c.Font.Name = "ＭＳ 明朝"
        c.Font.Size = 12.0
        break
time.sleep(0.1)
xs = measure(doc)
for i in range(len(xs)):
    adv = round(xs[i+1][1] - xs[i][1], 3) if i+1 < len(xs) else "(last)"
    print(f"  {i}: {xs[i][0]!r} font={xs[i][2]!r} adv={adv}")
doc.Close(SaveChanges=False)

# Test 3: opposite "はM" two runs
print("\n=== Test 3: two runs 'はM' (は:ＭＳ明朝, M:TNR) 12pt ===")
doc = word.Documents.Add()
time.sleep(0.2)
rng = doc.Range()
rng.InsertAfter("は")
rng = doc.Range()
rng.Font.Name = "ＭＳ 明朝"
rng.Font.Size = 12.0
rng = doc.Range()
rng.Collapse(0)
rng.InsertAfter("M")
chars = doc.Range().Characters
for ci in range(1, chars.Count+1):
    c = chars(ci)
    if c.Text == "M":
        c.Font.Name = "Times New Roman"
        c.Font.Size = 12.0
        break
time.sleep(0.1)
xs = measure(doc)
for i in range(len(xs)):
    adv = round(xs[i+1][1] - xs[i][1], 3) if i+1 < len(xs) else "(last)"
    print(f"  {i}: {xs[i][0]!r} font={xs[i][2]!r} adv={adv}")
doc.Close(SaveChanges=False)

# Test 4: ". 日" two runs (Latin period+space, then CJK)
print("\n=== Test 4: two runs '. 日' (.: TNR, : TNR, 日:ＭＳ明朝) 12pt ===")
doc = word.Documents.Add()
time.sleep(0.2)
rng = doc.Range()
rng.InsertAfter(". ")
rng = doc.Range()
rng.Font.Name = "Times New Roman"
rng.Font.Size = 12.0
rng = doc.Range()
rng.Collapse(0)
rng.InsertAfter("日")
chars = doc.Range().Characters
for ci in range(1, chars.Count+1):
    c = chars(ci)
    if c.Text == "日":
        c.Font.Name = "ＭＳ 明朝"
        c.Font.Size = 12.0
        break
time.sleep(0.1)
xs = measure(doc)
for i in range(len(xs)):
    adv = round(xs[i+1][1] - xs[i][1], 3) if i+1 < len(xs) else "(last)"
    print(f"  {i}: {xs[i][0]!r} font={xs[i][2]!r} adv={adv}")
doc.Close(SaveChanges=False)

word.Quit()
