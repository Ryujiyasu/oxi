"""
COM measurement: MS Mincho fullwidth character width at 10.5pt vs 10pt.
Question: Is fullwidth width == fontSize, or does GDI rounding make it 10pt at 10.5pt?

Method: Type fullwidth chars, measure X positions of consecutive chars.
Advance width = x[i+1] - x[i].
"""
import win32com.client, time, sys, os, json

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

word = win32com.client.gencache.EnsureDispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False
time.sleep(1)

doc = word.Documents.Add()
time.sleep(1)

sel = word.Selection

# Set page to A4, narrow margins for enough space
doc.PageSetup.PageWidth = 595.3  # A4
doc.PageSetup.PageHeight = 841.9
doc.PageSetup.LeftMargin = 72    # 1 inch
doc.PageSetup.RightMargin = 72

test_sizes = [10.5, 10.0, 11.0, 9.0]
test_chars = "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほ"  # 30 chars
test_fonts = ["MS 明朝", "ＭＳ 明朝"]

results = {}

for font_name in test_fonts:
    for size in test_sizes:
        sel.Font.Name = font_name
        sel.Font.Size = size
        sel.TypeText(test_chars)
        sel.TypeParagraph()

# Measure
p_idx = 0
for font_name in test_fonts:
    for size in test_sizes:
        p_idx += 1
        p = doc.Paragraphs(p_idx)
        text = p.Range.Text.rstrip('\r\n')
        actual_size = p.Range.Font.Size

        # Measure X positions of first 10 chars
        positions = []
        for i in range(min(11, len(text))):
            r = doc.Range(p.Range.Start + i, p.Range.Start + i + 1)
            x = r.Information(14)  # wdHorizontalPositionRelativeToPage
            positions.append(round(x, 4))

        # Calculate advance widths
        advances = []
        for i in range(len(positions) - 1):
            adv = round(positions[i+1] - positions[i], 4)
            advances.append(adv)

        key = f"{font_name}_{size}pt"
        results[key] = {
            "font": font_name,
            "requested_size": size,
            "actual_size": actual_size,
            "x_positions": positions,
            "advances": advances,
            "avg_advance": round(sum(advances) / len(advances), 4) if advances else 0,
        }

        print(f"{key}: actual={actual_size}pt, advances={advances[:5]}..., avg={results[key]['avg_advance']}")

# Also measure with explicit document (0e7a) context
print("\n--- Direct width measurement via Information(14) ---")
for key, data in results.items():
    advs = data["advances"]
    if advs:
        min_a = min(advs)
        max_a = max(advs)
        print(f"{key}: min={min_a}, max={max_a}, avg={data['avg_advance']}, all_same={min_a == max_a}")

# Save results
out_path = os.path.join(os.path.dirname(__file__), 'output', 'ms_mincho_fullwidth.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nSaved to {out_path}")

doc.Close(SaveChanges=False)
word.Quit()
