"""COM-verify trHeight hRule semantics.

Test cases:
- hRule unspecified (default): how does Word actually behave?
- hRule="atLeast"
- hRule="exact"
- Various trHeight values vs natural content height
"""
import win32com.client, time, os, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

WD_ROW_HEIGHT_AUTO = 0  # wdRowHeightAuto
WD_ROW_HEIGHT_AT_LEAST = 1  # wdRowHeightAtLeast
WD_ROW_HEIGHT_EXACTLY = 2  # wdRowHeightExactly

def measure_table(content, rule, height_pt):
    doc = word.Documents.Add()
    time.sleep(0.2)
    ps = doc.PageSetup
    ps.PageWidth = 612; ps.LeftMargin = 90; ps.RightMargin = 90
    ps.TopMargin = 72; ps.BottomMargin = 72
    rng = doc.Range()
    table = doc.Tables.Add(rng, 2, 1)
    table.Cell(1, 1).Range.Text = content
    table.Rows(1).HeightRule = rule
    if rule != WD_ROW_HEIGHT_AUTO:
        table.Rows(1).Height = height_pt
    table.Rows(2).HeightRule = WD_ROW_HEIGHT_AUTO
    time.sleep(0.2)
    try:
        y1 = table.Rows(1).Cells(1).Range.Information(6)
        y2 = table.Rows(2).Cells(1).Range.Information(6)
    except Exception:
        doc.Close(SaveChanges=False)
        return None, None
    doc.Close(SaveChanges=False)
    return y1, y2

def short_text():
    return "A"  # 1 char, ~13.8pt single line

def tall_text():
    return "A\nB\nC\nD"  # 4 lines

print("=== trHeight rule semantics ===")
print(f"{'rule':<10} {'requested':<10} {'content':<8} {'r1_y':<7} {'r2_y':<7} {'r1_height':<10}")

for rule_label, rule in [("auto", WD_ROW_HEIGHT_AUTO), ("atLeast", WD_ROW_HEIGHT_AT_LEAST), ("exact", WD_ROW_HEIGHT_EXACTLY)]:
    for height in [10.0, 20.0, 50.0, 100.0]:
        for label, txt in [("short(1l)", short_text()), ("tall(4l)", tall_text())]:
            y1, y2 = measure_table(txt, rule, height)
            if y1 is not None:
                row1_h = y2 - y1
                print(f"{rule_label:<10} {height:<10} {label:<10} {y1:<7.2f} {y2:<7.2f} {row1_h:<10.2f}")

word.Quit()
