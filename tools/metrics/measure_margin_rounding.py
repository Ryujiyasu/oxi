"""COM: Measure actual page margins vs XML twip values.

Check how Word rounds margin values from twips to points.
"""
import win32com.client
import os, time, json

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

docs = [
    "e3c545fac7a7_LOD_Handbook.docx",
    "b837808d0555_20240705_resources_data_guideline_02.docx",
    "459f05f1e877_kyodokenkyuyoushiki01.docx",
    "0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx",
    "3a4f9fbe1a83_001620506.docx",
]

for docname in docs:
    path = os.path.abspath(f"tools/golden-test/documents/docx/{docname}")
    if not os.path.exists(path):
        continue

    doc = word.Documents.Open(path, ReadOnly=True)
    time.sleep(0.5)

    ps = doc.Sections(1).PageSetup
    # COM reports in points
    print(f"\n=== {docname} ===")
    print(f"  TopMargin:    {ps.TopMargin:.2f}pt ({ps.TopMargin*20:.0f}tw)")
    print(f"  BottomMargin: {ps.BottomMargin:.2f}pt ({ps.BottomMargin*20:.0f}tw)")
    print(f"  LeftMargin:   {ps.LeftMargin:.2f}pt ({ps.LeftMargin*20:.0f}tw)")
    print(f"  RightMargin:  {ps.RightMargin:.2f}pt ({ps.RightMargin*20:.0f}tw)")

    # First paragraph Y position
    if doc.Paragraphs.Count > 0:
        p1 = doc.Paragraphs(1)
        y1 = p1.Range.Information(6)
        x1 = p1.Range.Information(5)
        print(f"  P1 y={y1:.2f} x={x1:.2f}")
        print(f"  y vs TopMargin: diff={y1 - ps.TopMargin:.2f}pt")

    doc.Close(SaveChanges=False)

# Test with synthetic documents at various margins
print(f"\n=== Synthetic margin tests ===")
test_margins_tw = [1134, 1077, 1021, 851, 720, 1418, 1440, 567, 284]
for mtw in test_margins_tw:
    doc = word.Documents.Add()
    time.sleep(0.3)
    ps = doc.Sections(1).PageSetup
    ps.TopMargin = mtw / 20.0
    time.sleep(0.1)

    # Insert text
    rng = doc.Range()
    rng.InsertAfter("Test")
    rng.Font.Name = "ＭＳ 明朝"
    rng.Font.Size = 10.5
    time.sleep(0.1)

    actual_top = ps.TopMargin
    p1_y = doc.Paragraphs(1).Range.Information(6)

    xml_pt = mtw / 20.0
    print(f"  {mtw}tw = {xml_pt:.2f}pt -> COM TopMargin={actual_top:.2f}pt, P1 y={p1_y:.2f}pt, diff={p1_y - xml_pt:.2f}pt")

    doc.Close(SaveChanges=False)

word.Quit()
