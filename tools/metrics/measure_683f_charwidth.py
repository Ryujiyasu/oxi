"""Measure character X positions in 683f table cell to find colon width."""
import win32com.client, time, sys, os
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = os.path.abspath("tools/golden-test/documents/docx/683ffcab86e2_20230331_resources_open_data_contract_addon_00.docx")

word = win32com.client.Dispatch("Word.Application")
word.Visible = True
word.DisplayAlerts = False
doc = word.Documents.Open(DOC, ReadOnly=True)
time.sleep(1)

# Table 1, Cell 1 — "解説：本条項は総則等に..."
t = doc.Tables(1)
cell = t.Cell(1, 1)
rng = cell.Range
chars = rng.Characters
n = chars.Count

# wdHorizontalPositionRelativeToPage = 5
WD_X = 5
WD_Y = 6

print(f"Cell chars: {n}")
print(f"First 20 chars X positions:")
prev_x = None
for i in range(1, min(21, n + 1)):
    c = chars(i)
    ch = c.Text
    if ch in ('\r', '\x07'):
        print(f"  [{i}] '\\r' or '\\x07'")
        continue
    x = c.Information(WD_X)
    y = c.Information(WD_Y)
    width = x - prev_x if prev_x is not None else 0
    print(f"  [{i}] '{ch}' (U+{ord(ch):04X}) x={x:.2f} y={y:.2f} width={width:.2f}pt")
    prev_x = x

doc.Close(SaveChanges=False)
word.Quit()
