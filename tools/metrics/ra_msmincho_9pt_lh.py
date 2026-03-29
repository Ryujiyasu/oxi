#!/usr/bin/env python3
"""COM measurement: MS Mincho 9pt line height (no grid, no spacing)."""
import win32com.client
import os, time, json

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    # Create a new document with MS Mincho 9pt, no grid, no spacing
    doc = word.Documents.Add()
    time.sleep(1)

    # Remove grid
    sec = doc.Sections(1)
    pf = sec.PageSetup
    # A4
    pf.PageWidth = 595.3
    pf.PageHeight = 841.9
    pf.TopMargin = 56.7   # 1134tw / 20
    pf.BottomMargin = 56.7
    pf.LeftMargin = 53.85  # 1077tw / 20
    pf.RightMargin = 53.85

    # Set layout mode to no grid
    doc.Sections(1).PageSetup.LayoutMode = 0  # wdLayoutModeDefault (no grid)

    # Set default paragraph: no spacing
    rng = doc.Range()
    rng.Font.Name = "ＭＳ 明朝"
    rng.Font.Size = 9
    rng.ParagraphFormat.SpaceAfter = 0
    rng.ParagraphFormat.SpaceBefore = 0
    rng.ParagraphFormat.LineSpacingRule = 0  # wdLineSpaceSingle

    # Insert 10 identical single-line paragraphs
    for i in range(10):
        rng = doc.Range()
        rng.InsertAfter(f"テスト行{i+1}あいうえお\r")

    time.sleep(1)
    doc.Repaginate()
    time.sleep(1)

    print("=== MS Mincho 9pt, no grid, no spacing ===")
    results = []
    for i in range(1, min(12, doc.Paragraphs.Count + 1)):
        p = doc.Paragraphs(i)
        rng = p.Range
        y = rng.Information(6)  # wdVerticalPositionRelativeToPage
        font = rng.Font.Name
        size = rng.Font.Size
        text = rng.Text[:30].replace('\r', '')
        results.append({"idx": i, "y": round(y, 4), "font": font, "size": size})
        print(f"  P{i}: y={y:.4f}pt, font={font}, size={size}")

    print("\n=== Gaps ===")
    for i in range(1, len(results)):
        gap = results[i]["y"] - results[i-1]["y"]
        print(f"  P{results[i-1]['idx']}→P{results[i]['idx']}: gap={gap:.4f}pt")

    # Also test with Arial Unicode MS
    doc2 = word.Documents.Add()
    time.sleep(1)
    doc2.Sections(1).PageSetup.LayoutMode = 0
    pf2 = doc2.Sections(1).PageSetup
    pf2.PageWidth = 595.3
    pf2.PageHeight = 841.9
    pf2.TopMargin = 56.7
    pf2.BottomMargin = 56.7

    rng2 = doc2.Range()
    rng2.Font.Name = "Arial Unicode MS"
    rng2.Font.Size = 9
    rng2.ParagraphFormat.SpaceAfter = 0
    rng2.ParagraphFormat.SpaceBefore = 0
    rng2.ParagraphFormat.LineSpacingRule = 0

    for i in range(10):
        rng2 = doc2.Range()
        rng2.InsertAfter(f"テスト行{i+1}あいうえお\r")

    time.sleep(1)
    doc2.Repaginate()
    time.sleep(1)

    print("\n=== Arial Unicode MS 9pt, no grid, no spacing ===")
    results2 = []
    for i in range(1, min(12, doc2.Paragraphs.Count + 1)):
        p = doc2.Paragraphs(i)
        rng = p.Range
        y = rng.Information(6)
        font = rng.Font.Name
        size = rng.Font.Size
        results2.append({"idx": i, "y": round(y, 4), "font": font, "size": size})
        print(f"  P{i}: y={y:.4f}pt, font={font}, size={size}")

    print("\n=== Gaps (Arial Unicode MS) ===")
    for i in range(1, len(results2)):
        gap = results2[i]["y"] - results2[i-1]["y"]
        print(f"  P{results2[i-1]['idx']}→P{results2[i]['idx']}: gap={gap:.4f}pt")

    # Also test MS Mincho 9pt WITH grid (linePitch=18pt = 360tw, typical default)
    doc3 = word.Documents.Add()
    time.sleep(1)
    pf3 = doc3.Sections(1).PageSetup
    pf3.PageWidth = 595.3
    pf3.PageHeight = 841.9
    pf3.TopMargin = 56.7
    pf3.BottomMargin = 56.7
    # Grid mode
    doc3.Sections(1).PageSetup.LayoutMode = 1  # wdLayoutModeLineGrid

    rng3 = doc3.Range()
    rng3.Font.Name = "ＭＳ 明朝"
    rng3.Font.Size = 9
    rng3.ParagraphFormat.SpaceAfter = 0
    rng3.ParagraphFormat.SpaceBefore = 0
    rng3.ParagraphFormat.LineSpacingRule = 0

    for i in range(10):
        rng3 = doc3.Range()
        rng3.InsertAfter(f"テスト行{i+1}あいうえお\r")

    time.sleep(1)
    doc3.Repaginate()
    time.sleep(1)

    print(f"\n=== MS Mincho 9pt WITH grid (linePitch={pf3.LinePitch}pt) ===")
    results3 = []
    for i in range(1, min(12, doc3.Paragraphs.Count + 1)):
        p = doc3.Paragraphs(i)
        rng = p.Range
        y = rng.Information(6)
        font = rng.Font.Name
        size = rng.Font.Size
        results3.append({"idx": i, "y": round(y, 4), "font": font, "size": size})
        print(f"  P{i}: y={y:.4f}pt, font={font}, size={size}")

    print("\n=== Gaps (with grid) ===")
    for i in range(1, len(results3)):
        gap = results3[i]["y"] - results3[i-1]["y"]
        print(f"  P{results3[i-1]['idx']}→P{results3[i]['idx']}: gap={gap:.4f}pt")

    doc.Close(False)
    doc2.Close(False)
    doc3.Close(False)
    word.Quit()

if __name__ == "__main__":
    main()
