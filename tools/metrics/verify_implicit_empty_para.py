"""COM-verify what Word does when body has no <w:p> elements (only sectPr).

Test:
1. Open header_page_number_01.docx — body XML has only sectPr
2. Read Documents object's Paragraphs collection
3. Confirm Word reports >=1 paragraph and where it's positioned
4. Compare against a doc with explicit empty <w:p>
"""
import win32com.client, time, os, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

for doc_name in ["header_page_number_01.docx", "footer_complex_01.docx"]:
    path = os.path.abspath(f"pipeline_data/docx/{doc_name}")
    print(f"\n=== {doc_name} ===")
    doc = word.Documents.Open(path, ReadOnly=True)
    time.sleep(0.4)
    print(f"  Paragraphs.Count = {doc.Paragraphs.Count}")
    print(f"  StoryRanges count = {doc.StoryRanges.Count}")
    print(f"  Range.End = {doc.Range().End}")
    print(f"  Range.Text = {doc.Range().Text!r}")
    if doc.Paragraphs.Count >= 1:
        p1 = doc.Paragraphs(1)
        print(f"  P1.Range.Text = {p1.Range.Text!r}")
        print(f"  P1.Range.Start = {p1.Range.Start}")
        print(f"  P1.Range.End = {p1.Range.End}")
        try:
            y = p1.Range.Information(6)
            x = p1.Range.Information(5)
            print(f"  P1 position: x={x}, y={y}")
        except Exception as e:
            print(f"  P1 position read err: {e}")
        print(f"  P1.Format.SpaceBefore = {p1.SpaceBefore}")
        print(f"  P1.Format.LineSpacing = {p1.LineSpacing}")
    doc.Close(SaveChanges=False)

word.Quit()
