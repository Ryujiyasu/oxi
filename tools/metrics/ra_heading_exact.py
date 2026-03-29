#!/usr/bin/env python3
"""COM: Exact measurement of heading + body paragraph gap in the actual document."""
import win32com.client
import time

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    # Test 1: Replicate exact heading format
    doc = word.Documents.Add()
    time.sleep(0.5)
    doc.Sections(1).PageSetup.LayoutMode = 0  # no grid
    pf = doc.Sections(1).PageSetup
    pf.TopMargin = 56.7
    pf.BottomMargin = 56.7

    # Set default paragraph font to MS Mincho 10.5pt (like docDefaults)
    style = doc.Styles(-1)  # wdStyleNormal
    style.Font.Name = "ＭＳ 明朝"
    style.Font.Size = 10.5
    style.ParagraphFormat.SpaceAfter = 0
    style.ParagraphFormat.SpaceBefore = 0

    # Insert heading (Arial Unicode MS Bold 10pt) + body (MS Mincho 9pt)
    rng = doc.Range()
    rng.InsertAfter("総則\r")  # heading
    rng = doc.Paragraphs(1).Range
    rng.Font.Name = "Arial Unicode MS"
    rng.Font.Size = 10
    rng.Font.Bold = True
    rng.ParagraphFormat.SpaceAfter = 0
    rng.ParagraphFormat.SpaceBefore = 0

    # Body paragraphs
    for i in range(5):
        rng = doc.Range()
        rng.InsertAfter(f"第{i+1}条　委託者は、この契約書に基づき、別紙仕様書に付属する設計図書、図面、質問回答書等がある場合にはこれらの書面を含む。\r")
    # Set body font
    for i in range(2, 7):
        rng = doc.Paragraphs(i).Range
        rng.Font.Name = "ＭＳ 明朝"
        rng.Font.Size = 9
        rng.ParagraphFormat.SpaceAfter = 0
        rng.ParagraphFormat.SpaceBefore = 0

    time.sleep(1)
    doc.Repaginate()
    time.sleep(0.5)

    print("=== Test 1: Heading + Body (replicated document settings) ===")
    for i in range(1, 7):
        p = doc.Paragraphs(i)
        y = p.Range.Information(6)
        font = p.Range.Font.Name
        size = p.Range.Font.Size
        text = p.Range.Text[:40].replace('\r', '')
        print(f"  P{i}: y={y:.4f}, font={font}, size={size}, bold={p.Range.Font.Bold}, \"{text}\"")

    print("\n=== Gaps ===")
    prev_y = None
    for i in range(1, 7):
        y = doc.Paragraphs(i).Range.Information(6)
        if prev_y is not None:
            print(f"  P{i-1}->P{i}: gap={y-prev_y:.4f}")
        prev_y = y

    # Test 2: Pure CJK text with Arial 10pt bold (what Oxi would use as fallback)
    doc2 = word.Documents.Add()
    time.sleep(0.5)
    doc2.Sections(1).PageSetup.LayoutMode = 0
    style2 = doc2.Styles(-1)  # wdStyleNormal
    style2.Font.Name = "ＭＳ 明朝"
    style2.Font.Size = 10.5
    style2.ParagraphFormat.SpaceAfter = 0
    style2.ParagraphFormat.SpaceBefore = 0

    rng = doc2.Range()
    rng.InsertAfter("総則\r")
    rng = doc2.Paragraphs(1).Range
    rng.Font.Name = "Arial"  # Oxi fallback
    rng.Font.Size = 10
    rng.Font.Bold = True
    rng.ParagraphFormat.SpaceAfter = 0
    rng.ParagraphFormat.SpaceBefore = 0

    for i in range(5):
        rng = doc2.Range()
        rng.InsertAfter(f"第{i+1}条　テスト文。\r")
    for i in range(2, 7):
        rng = doc2.Paragraphs(i).Range
        rng.Font.Name = "ＭＳ 明朝"
        rng.Font.Size = 9
        rng.ParagraphFormat.SpaceAfter = 0
        rng.ParagraphFormat.SpaceBefore = 0

    time.sleep(1)
    doc2.Repaginate()
    time.sleep(0.5)

    print("\n=== Test 2: Arial 10pt bold heading + MS Mincho 9pt body ===")
    for i in range(1, 7):
        p = doc2.Paragraphs(i)
        y = p.Range.Information(6)
        font = p.Range.Font.Name
        size = p.Range.Font.Size
        text = p.Range.Text[:40].replace('\r', '')
        print(f"  P{i}: y={y:.4f}, font={font}, size={size}, bold={p.Range.Font.Bold}, \"{text}\"")

    print("\n=== Gaps ===")
    prev_y = None
    for i in range(1, 7):
        y = doc2.Paragraphs(i).Range.Information(6)
        if prev_y is not None:
            print(f"  P{i-1}->P{i}: gap={y-prev_y:.4f}")
        prev_y = y

    doc.Close(False)
    doc2.Close(False)
    word.Quit()

if __name__ == "__main__":
    main()
