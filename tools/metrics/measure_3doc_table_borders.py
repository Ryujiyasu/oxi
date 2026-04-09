"""COM-measure row heights for 3 docs with 1-row tables.
Goal: confirm/refute the hypothesis that outer top+bottom borders
add to first/last row height.
"""
import win32com.client, time, sys, os
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCS = [
    "tools/golden-test/documents/docx/683ffcab86e2_20230331_resources_open_data_contract_addon_00.docx",
    "tools/golden-test/documents/docx/4a36b62555f2_kyodokenkyuyoushiki10.docx",
    "tools/golden-test/documents/docx/e201249db062_tokumei_08_05.docx",
]

WD_Y_PAGE = 6
WD_IN_TABLE = 12

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

def measure(doc_path):
    doc = word.Documents.Open(os.path.abspath(doc_path), ReadOnly=True)
    time.sleep(0.5)
    print(f"\n========== {os.path.basename(doc_path)} ==========")
    print(f"Total tables: {doc.Tables.Count}")
    for ti in range(1, doc.Tables.Count + 1):
        t = doc.Tables(ti)
        rows = t.Rows.Count
        if rows != 1:
            continue
        # Find paragraph index of table top and the para after
        # Table 1 starts at some para; the next para after the table tells us the table bottom
        try:
            tbl_top_y = t.Range.Information(WD_Y_PAGE)
        except: continue
        # Find the next paragraph not in this table
        # iterate doc paragraphs
        next_y = None
        for pi in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(pi)
            try:
                py = p.Range.Information(WD_Y_PAGE)
                in_tbl = p.Range.Information(WD_IN_TABLE)
            except: continue
            if py > tbl_top_y + 0.5 and not in_tbl:
                next_y = py
                break
        cell_text = t.Cell(1,1).Range.Text[:30].replace('\r','\\r').replace('\x07','\\BEL')
        print(f"\n  Table {ti}: 1 row")
        print(f"    Top Y: {tbl_top_y:.2f}")
        print(f"    Next-para Y: {next_y}")
        if next_y is not None:
            print(f"    Inferred table height: {next_y - tbl_top_y:.2f}pt")
        # borders
        for bi, name in [(-1,"Top"),(-2,"Left"),(-3,"Bot"),(-4,"Right"),(-5,"InsH"),(-6,"InsV")]:
            try:
                b = t.Borders(bi)
                if b.LineStyle != 0:
                    print(f"    {name}: sz={b.LineWidth/8:.1f}pt style={b.LineStyle}")
            except: pass
        # cell content
        print(f"    cell text: '{cell_text}'")
        # font / line height of cell
        try:
            cr = t.Cell(1,1).Range
            print(f"    font: {cr.Font.Name} sz={cr.Font.Size}")
            # lineSpacing
            f = cr.ParagraphFormat
            print(f"    LineSpacing={f.LineSpacing} Rule={f.LineSpacingRule}")
            print(f"    SpaceBefore={f.SpaceBefore} SpaceAfter={f.SpaceAfter}")
        except Exception as e:
            print(f"    err: {e}")
    doc.Close(SaveChanges=False)

for d in DOCS:
    if os.path.exists(d):
        measure(d)
    else:
        print(f"NOT FOUND: {d}")

word.Quit()
