"""COM-measure Word's text Y position for exact/atLeast line spacing.

For each (font, font_size, line_height_pt) combination:
  - Create doc, top margin exactly 1 inch (72pt)
  - Single para with "A漢" and exact/atLeast line spacing
  - Read Information(6) for first char, second char, second-line first char
  - Compute offset from top_margin

Hypotheses to discriminate:
  H_TOP    : text Y = top_margin                       (offset = 0)
  H_CENTER : text Y = top_margin + (lh - natural)/2    (offset = (lh-natural)/2)
  H_BOTTOM : text Y = top_margin + (lh - natural)      (offset = lh - natural)
  H_BASELINE_DESC : text Y = top_margin + lh - descent (offset = lh - descent)
"""
import win32com.client
import time
import sys
import json

sys.stdout.reconfigure(encoding='utf-8')

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

TOP_MARGIN_PT = 72.0  # 1 inch

# wdLineSpaceExactly = 4, wdLineSpaceAtLeast = 3
WD_EXACTLY = 4
WD_AT_LEAST = 3
# Information constants
WD_HORIZ_POS_REL_TEXT_BOUND = 1
WD_VERT_POS_REL_PAGE = 6

def measure_one(font_name, font_size, line_height_pt, rule):
    doc = word.Documents.Add()
    time.sleep(0.15)
    doc.PageSetup.TopMargin = TOP_MARGIN_PT
    doc.PageSetup.LeftMargin = 72.0
    doc.PageSetup.RightMargin = 72.0
    doc.PageSetup.BottomMargin = 72.0

    is_cjk = any(ord(c) > 0x2000 for c in font_name)
    text = "A漢B字\nL2first" if is_cjk else "ABCDE\nL2first"
    rng = doc.Range()
    rng.InsertAfter(text)
    rng = doc.Range()
    rng.Font.Name = font_name
    if is_cjk:
        rng.Font.NameFarEast = font_name
    rng.Font.Size = font_size
    rng.ParagraphFormat.LineSpacingRule = rule
    rng.ParagraphFormat.LineSpacing = line_height_pt
    rng.ParagraphFormat.SpaceBefore = 0
    rng.ParagraphFormat.SpaceAfter = 0
    time.sleep(0.15)

    chars = doc.Range().Characters
    samples = []
    for ci in range(1, min(8, chars.Count + 1)):
        try:
            c = chars(ci)
            ch = c.Text
            if ch in ('\r','\x07','\n'):
                continue
            cy = c.Information(WD_VERT_POS_REL_PAGE)
            cx = c.Information(WD_HORIZ_POS_REL_TEXT_BOUND)
            samples.append({'ch': ch, 'x': round(cx,3), 'y': round(cy,3)})
        except Exception as e:
            samples.append({'err': str(e)})

    doc.Close(SaveChanges=False)
    return samples


CASES = [
    # (font, size, lh, rule, label)
    ("Calibri",   11.0, 14.0, WD_EXACTLY,  "Calibri 11 exact 14"),
    ("Calibri",   11.0, 17.0, WD_EXACTLY,  "Calibri 11 exact 17"),
    ("Calibri",   11.0, 20.0, WD_EXACTLY,  "Calibri 11 exact 20"),
    ("Calibri",   11.0, 24.0, WD_EXACTLY,  "Calibri 11 exact 24"),
    ("Calibri",   11.0, 30.0, WD_EXACTLY,  "Calibri 11 exact 30"),

    ("ＭＳ 明朝", 10.5, 17.0, WD_EXACTLY,  "MS Mincho 10.5 exact 17"),
    ("ＭＳ 明朝", 10.5, 14.0, WD_EXACTLY,  "MS Mincho 10.5 exact 14"),
    ("ＭＳ 明朝", 10.5, 20.0, WD_EXACTLY,  "MS Mincho 10.5 exact 20"),
    ("ＭＳ 明朝", 10.5, 25.0, WD_EXACTLY,  "MS Mincho 10.5 exact 25"),
    ("ＭＳ 明朝", 12.0, 20.0, WD_EXACTLY,  "MS Mincho 12 exact 20"),

    ("ＭＳ ゴシック", 10.5, 17.0, WD_EXACTLY, "MS Gothic 10.5 exact 17"),
    ("ＭＳ ゴシック", 10.5, 25.0, WD_EXACTLY, "MS Gothic 10.5 exact 25"),

    ("Calibri",   11.0, 14.0, WD_AT_LEAST, "Calibri 11 atLeast 14"),
    ("Calibri",   11.0, 20.0, WD_AT_LEAST, "Calibri 11 atLeast 20"),
    ("ＭＳ 明朝", 10.5, 17.0, WD_AT_LEAST, "MS Mincho 10.5 atLeast 17"),
    ("ＭＳ 明朝", 10.5, 25.0, WD_AT_LEAST, "MS Mincho 10.5 atLeast 25"),
]

results = []
for font, fs, lh, rule, label in CASES:
    try:
        samples = measure_one(font, fs, lh, rule)
        # Find first L1 char and second-line first char
        l1y = None
        l2y = None
        if samples:
            l1y = samples[0].get('y')
            # second line: find a y > l1y + 5
            for s in samples[1:]:
                y = s.get('y')
                if y is not None and l1y is not None and y - l1y > 5:
                    l2y = y
                    break
        offset_from_top = (l1y - TOP_MARGIN_PT) if l1y is not None else None
        actual_lh = (l2y - l1y) if (l1y is not None and l2y is not None) else None
        results.append({
            'label': label, 'font': font, 'size': fs, 'lh': lh,
            'rule': 'exact' if rule == WD_EXACTLY else 'atLeast',
            'l1y': l1y, 'l2y': l2y,
            'offset_from_top': round(offset_from_top, 3) if offset_from_top is not None else None,
            'actual_lh': round(actual_lh, 3) if actual_lh is not None else None,
            'samples': samples,
        })
        print(f"  {label:35s} l1y={l1y} offset={offset_from_top} actual_lh={actual_lh}")
    except Exception as e:
        print(f"  {label}: ERR {e}")
        results.append({'label': label, 'err': str(e)})

word.Quit()

with open('tools/metrics/output/text_y_offset_exact.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)

print("\nSaved to tools/metrics/output/text_y_offset_exact.json")
