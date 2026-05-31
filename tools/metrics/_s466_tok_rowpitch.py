"""tokumei row-height mechanism: Word vs Oxi per-row top Y for the table on a
given page. If Oxi pitches are uniformly LARGER => CJK cell rows over-snapped
(too tall), pushing content down (matches scout +9px Oxi-too-low). Page-relative
via Information(6) wdVerticalPositionRelativeToPage + R30 collapsed start."""
import json
import win32com.client as win32
VPOS=6;PAGEINFO=3
DOCX=r"C:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\d4d126dfe1d9_tokumei_08_01-3.docx"
DUMP="C:/tmp/tok_d4/layout.json"
TARGET_PAGE=4

# Oxi: per (row_idx) min-y for the table cells on page 4
d=json.load(open(DUMP,encoding="utf-8"))
orow={}
for pg in d["pages"]:
    if pg["page"]!=TARGET_PAGE: continue
    for el in pg["elements"]:
        if el["type"]!="text": continue
        r=el.get("cell_row_idx")
        if r is None: continue
        orow[r]=min(orow.get(r,1e9),el["y"])
oxi_rows=sorted(orow.items())
print("Oxi page-4 row tops (row_idx: y):")
for r,y in oxi_rows: print(f"  r{r}: {y:.2f}")
oxi_pitch=[round(oxi_rows[i+1][1]-oxi_rows[i][1],2) for i in range(len(oxi_rows)-1)]
print("Oxi row pitches:",oxi_pitch)

# Word: find the table, measure each row's first cell top Y, keep those on page 4
word=win32.gencache.EnsureDispatch("Word.Application");word.Visible=False
doc=word.Documents.Open(DOCX,ReadOnly=True)
wrows=[]
try:
    for ti in range(1,doc.Tables.Count+1):
        t=doc.Tables(ti)
        for r in range(1,t.Rows.Count+1):
            try:
                c=t.Cell(r,1).Range
            except: continue
            st=doc.Range(c.Start,c.Start)
            pg=st.Information(PAGEINFO)
            if pg==TARGET_PAGE:
                wrows.append(round(st.Information(VPOS),2))
        # also handle multi-table; but tokumei likely 1 big table
finally:
    doc.Close(False);word.Quit()
wrows=sorted(set(wrows))
print("\nWord page-4 row tops:",wrows)
wpitch=[round(wrows[i+1]-wrows[i],2) for i in range(len(wrows)-1)]
print("Word row pitches:",wpitch)
if wpitch and oxi_pitch:
    import statistics
    print(f"\nWord pitch median {statistics.median(wpitch):.2f}  Oxi pitch median {statistics.median(oxi_pitch):.2f}")
    print(f"=> Oxi - Word (per-row): {statistics.median(oxi_pitch)-statistics.median(wpitch):+.2f}pt")
