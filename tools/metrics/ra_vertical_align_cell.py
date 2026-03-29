"""
Ra: テーブルセル垂直配置(vAlign)の精度確定
- top / center / bottom の正確なY座標
- 複数行テキストのcenter計算
- セル高さとテキスト高さの関係
"""
import win32com.client, json, os

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

results = []


def test_valign():
    wdoc = word.Documents.Add()
    try:
        ps = wdoc.Sections(1).PageSetup
        ps.LeftMargin = 72

        wdoc.Content.Text = ""
        rng = wdoc.Range(0, 0)
        tbl = wdoc.Tables.Add(rng, 3, 3)
        tbl.Borders.Enable = True

        # Set fixed row height
        for r in range(1, 4):
            tbl.Rows(r).Height = 60  # 60pt tall
            tbl.Rows(r).HeightRule = 1  # wdRowHeightExactly

        # Set vertical alignment: row1=top, row2=center, row3=bottom
        valigns = [0, 1, 3]  # wdCellAlignVerticalTop, Center, Bottom
        labels = ["top", "center", "bottom"]

        for r in range(1, 4):
            for c in range(1, 4):
                cell = tbl.Cell(r, c)
                cell.VerticalAlignment = valigns[r-1]
                if c == 1:
                    cell.Range.Text = f"Single line ({labels[r-1]})"
                elif c == 2:
                    cell.Range.Text = f"Line 1\rLine 2 ({labels[r-1]})"
                else:
                    cell.Range.Text = f"Line 1\rLine 2\rLine 3 ({labels[r-1]})"
                cell.Range.Font.Name = "Calibri"
                cell.Range.Font.Size = 11

        wdoc.Repaginate()

        data = {"scenario": "valign", "row_height": 60, "cells": []}
        for r in range(1, 4):
            for c in range(1, 4):
                cell = tbl.Cell(r, c)
                # Get first paragraph Y
                first_para = cell.Range.Paragraphs(1)
                y1 = first_para.Range.Information(6)
                # Get last paragraph Y
                last_para = cell.Range.Paragraphs(cell.Range.Paragraphs.Count)
                y_last = last_para.Range.Information(6)

                data["cells"].append({
                    "row": r, "col": c,
                    "valign": labels[r-1],
                    "first_y": round(y1, 4),
                    "last_y": round(y_last, 4),
                    "para_count": cell.Range.Paragraphs.Count,
                })

        return data
    finally:
        wdoc.Close(False)


def test_valign_with_varying_heights():
    """vAlign center with different row heights."""
    wdoc = word.Documents.Add()
    try:
        ps = wdoc.Sections(1).PageSetup
        ps.LeftMargin = 72

        wdoc.Content.Text = ""
        rng = wdoc.Range(0, 0)
        tbl = wdoc.Tables.Add(rng, 4, 1)
        tbl.Borders.Enable = True

        heights = [40, 60, 80, 100]
        for r in range(1, 5):
            tbl.Rows(r).Height = heights[r-1]
            tbl.Rows(r).HeightRule = 1
            cell = tbl.Cell(r, 1)
            cell.VerticalAlignment = 1  # center
            cell.Range.Text = f"Center (h={heights[r-1]}pt)"
            cell.Range.Font.Name = "Calibri"
            cell.Range.Font.Size = 11

        wdoc.Repaginate()

        data = {"scenario": "valign_heights", "rows": []}
        for r in range(1, 5):
            cell = tbl.Cell(r, 1)
            para = cell.Range.Paragraphs(1)
            y = para.Range.Information(6)
            # Row top position
            row_top = tbl.Rows(r).Range.Information(6) if r == 1 else data["rows"][-1]["row_bottom"]
            if r == 1:
                row_top = round(tbl.Cell(1,1).Range.Information(6) - 20, 4)  # approximate

            data["rows"].append({
                "row": r,
                "height": heights[r-1],
                "text_y": round(y, 4),
                "row_bottom": round(y + heights[r-1], 4) if r == 1 else 0,
            })

        return data
    finally:
        wdoc.Close(False)


try:
    d1 = test_valign()
    results.append(d1)
    print("=== valign (row_height=60pt) ===")
    for c in d1["cells"]:
        print(f"  R{c['row']}C{c['col']} ({c['valign']}): first_y={c['first_y']}, "
              f"last_y={c['last_y']}, paras={c['para_count']}")

    d2 = test_valign_with_varying_heights()
    results.append(d2)
    print(f"\n=== valign center with varying heights ===")
    for r in d2["rows"]:
        print(f"  Row {r['row']} (h={r['height']}pt): text_y={r['text_y']}")

finally:
    word.Quit()

out_path = os.path.join(os.path.dirname(__file__), 'output', 'ra_vertical_align.json')
os.makedirs(os.path.dirname(out_path), exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print(f"\nResults saved to {out_path}")
