#!/usr/bin/env python3
"""COM: Measure heading line height for Arial Unicode MS vs Arial at 10pt."""
import win32com.client
import time

def measure_font(word, font_name, font_size):
    doc = word.Documents.Add()
    time.sleep(0.5)
    doc.Sections(1).PageSetup.LayoutMode = 0  # no grid

    rng = doc.Range()
    rng.Font.Name = font_name
    rng.Font.Size = font_size
    rng.ParagraphFormat.SpaceAfter = 0
    rng.ParagraphFormat.SpaceBefore = 0
    rng.ParagraphFormat.LineSpacingRule = 0

    for i in range(5):
        rng = doc.Range()
        rng.InsertAfter(f"Test line {i+1} テスト\r")

    time.sleep(0.5)
    doc.Repaginate()
    time.sleep(0.5)

    gaps = []
    prev_y = None
    for i in range(1, 6):
        p = doc.Paragraphs(i)
        y = p.Range.Information(6)
        actual_font = p.Range.Font.Name
        if prev_y is not None:
            gaps.append(y - prev_y)
        prev_y = y

    avg_gap = sum(gaps) / len(gaps) if gaps else 0
    doc.Close(False)
    return avg_gap, gaps, actual_font

def main():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    for font_name in ["Arial Unicode MS", "Arial", "ＭＳ 明朝"]:
        for font_size in [9.0, 10.0, 10.5]:
            avg, gaps, actual = measure_font(word, font_name, font_size)
            print(f"{font_name:20s} {font_size:5.1f}pt: avg_gap={avg:.4f}pt, gaps={[f'{g:.1f}' for g in gaps]}, actual_font={actual}")

    word.Quit()

if __name__ == "__main__":
    main()
