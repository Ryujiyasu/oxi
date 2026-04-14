"""Measure actual Word row heights for 0e7a contract sample table 1."""
import win32com.client, time, sys, os
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOC = os.path.abspath("tools/golden-test/documents/docx/0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx")

word = win32com.client.Dispatch("Word.Application")
word.Visible = True
word.DisplayAlerts = False
doc = word.Documents.Open(DOC, ReadOnly=True)
time.sleep(1)

t = doc.Tables(1)
print(f"Table 1: {t.Rows.Count} rows x {t.Columns.Count} cols")
prev_y = None
for ri in range(1, min(8, t.Rows.Count + 1)):
    r = t.Rows(ri)
    height_val = r.Height
    rule = r.HeightRule  # 0=auto, 1=atLeast, 2=exact
    rule_name = {0:'auto', 1:'atLeast', 2:'exact'}.get(rule, f'?{rule}')
    # Get actual rendered Y of the row's first cell
    first_cell = r.Cells(1)
    first_para = first_cell.Range.Paragraphs(1)
    y = first_para.Range.Information(6)  # vertical position
    txt = first_cell.Range.Text[:20]
    delta = y - prev_y if prev_y is not None else 0
    print(f"  row{ri}: HeightRule={rule_name} val={height_val} actual_y={y:.2f} delta={delta:.2f}  text={txt!r}")
    prev_y = y

doc.Close(SaveChanges=False)
word.Quit()
