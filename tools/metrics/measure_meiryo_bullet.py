"""COM measurement: Meiryo bullet (U+2022) width at 10.5/11/12pt.

Output values are in twips (matching com_tw_overrides.json scale).
"""
import win32com.client
import time

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

results = {}
for size in (10.5, 11, 12):
    doc = word.Documents.Add()
    time.sleep(0.3)
    # Use a delimiter so we can isolate the bullet width unambiguously.
    text = "A\u2022A"
    rng = doc.Range()
    rng.InsertAfter(text)
    rng = doc.Range()
    rng.Font.Name = "メイリオ"
    rng.Font.Size = size
    doc.Paragraphs(1).Alignment = 0
    time.sleep(0.1)

    chars = doc.Range().Characters
    xs = []
    for i in range(1, chars.Count + 1):
        c = chars(i)
        ch = c.Text
        if ch in ("\r", "\x07"):
            continue
        xs.append((ch, c.Information(5)))  # wdHorizontalPositionRelativeToPage (pt)

    # xs[0]=A, xs[1]=•, xs[2]=A
    bullet_width_pt = round(xs[2][1] - xs[1][1], 4)
    bullet_width_tw = round(bullet_width_pt * 20, 2)
    results[size] = (bullet_width_pt, bullet_width_tw)
    print(f"{size}pt: bullet width = {bullet_width_pt}pt = {bullet_width_tw}tw")
    doc.Close(SaveChanges=False)

word.Quit()
print("\nFor com_tw_overrides.json:")
for size, (pt, tw) in results.items():
    key = str(size).rstrip("0").rstrip(".") if size != 10.5 else "10.5"
    print(f'  "{key}": {{ ..., "8226": {tw} }}')
