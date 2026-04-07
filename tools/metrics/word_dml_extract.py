"""Extract Word DML (layout positions) via COM API and cache as JSON.

Usage:
  python word_dml_extract.py <docx_path>           # single file
  python word_dml_extract.py <docx_dir> [--all]     # batch all docx in dir

Output: pipeline_data/word_dml/<doc_id>.json
"""
import win32com.client
import json
import sys
import os
from pathlib import Path

CACHE_DIR = os.path.join(os.path.dirname(__file__), "..", "..", "pipeline_data", "word_dml")


def extract_dml(docx_path: str) -> dict:
    """Extract layout positions from Word via COM."""
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False  # Suppress "Save changes?" dialogs

    doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)

    result = {
        "file": os.path.basename(docx_path),
        "pages": doc.ComputeStatistics(2),  # wdStatisticPages
        "page_setup": {
            "page_width": doc.PageSetup.PageWidth,
            "page_height": doc.PageSetup.PageHeight,
            "margin_top": doc.PageSetup.TopMargin,
            "margin_bottom": doc.PageSetup.BottomMargin,
            "margin_left": doc.PageSetup.LeftMargin,
            "margin_right": doc.PageSetup.RightMargin,
        },
        "paragraphs": [],
        "tables": [],
    }

    # Extract paragraph positions
    for pi in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(pi)
        rng = p.Range
        try:
            page = rng.Information(1)  # wdActiveEndPageNumber
            y = rng.Information(6)     # wdVerticalPositionRelativeToPage
            x = rng.Information(5)     # wdHorizontalPositionRelativeToPage
        except:
            continue

        align = p.Alignment  # 0=left, 1=center, 2=right, 3=justify
        ls = p.Format.LineSpacing
        sa = p.Format.SpaceAfter
        sb = p.Format.SpaceBefore

        # Get font info from first character
        font_name = ""
        font_size = 0
        try:
            font_name = rng.Font.Name
            font_size = rng.Font.Size
        except:
            pass

        # Count lines and get per-line info
        chars = rng.Characters
        n_chars = chars.Count
        prev_y = None
        lines = []
        line_start_x = None
        line_char_count = 0

        for ci in range(1, n_chars + 1):
            try:
                c = chars(ci)
                ch = c.Text
                # Filter: \r=para mark, \x07=cell mark, \x0b=soft line break (<w:br/>)
                if ch == '\r' or ch == '\x07' or ch == '\x0b':
                    continue
                cy = c.Information(6)
                cx = c.Information(5)

                if prev_y is None or abs(cy - prev_y) > 0.5:
                    if prev_y is not None:
                        lines.append({
                            "y": round(prev_y, 2),
                            "x": round(line_start_x, 2),
                            "chars": line_char_count,
                        })
                    line_start_x = cx
                    line_char_count = 0
                    prev_y = cy

                line_char_count += 1
            except:
                continue

        if prev_y is not None and line_char_count > 0:
            lines.append({
                "y": round(prev_y, 2),
                "x": round(line_start_x, 2),
                "chars": line_char_count,
            })

        text_preview = rng.Text[:60].strip().replace('\r', '').replace('\x07', '')

        result["paragraphs"].append({
            "index": pi,
            "page": page,
            "x": round(x, 2),
            "y": round(y, 2),
            "align": align,
            "line_spacing": round(ls, 2),
            "space_after": round(sa, 2),
            "space_before": round(sb, 2),
            "font": font_name,
            "font_size": font_size,
            "lines": lines,
            "text": text_preview,
        })

    # Extract table positions
    for ti in range(1, doc.Tables.Count + 1):
        t = doc.Tables(ti)
        table_data = {
            "index": ti,
            "rows": t.Rows.Count,
            "cols": t.Columns.Count,
            "row_data": [],
        }

        for ri in range(1, t.Rows.Count + 1):
            row_info = {"cells": []}
            for ci in range(1, t.Columns.Count + 1):
                try:
                    cell = t.Cell(ri, ci)
                    cy = cell.Range.Information(6)
                    cx = cell.Range.Information(5)
                    cw = cell.Width
                    text = cell.Range.Text[:30].strip().replace('\r', '').replace('\x07', '')
                    row_info["cells"].append({
                        "row": ri, "col": ci,
                        "x": round(cx, 2), "y": round(cy, 2),
                        "width": round(cw, 2),
                        "text": text,
                    })
                except:
                    pass

            if row_info["cells"]:
                # Row Y = first cell Y
                row_info["y"] = row_info["cells"][0]["y"]
                table_data["row_data"].append(row_info)

        result["tables"].append(table_data)

    doc.Close(False)
    word.Quit()
    return result


def main():
    if len(sys.argv) < 2:
        print("Usage: python word_dml_extract.py <docx_path_or_dir> [--all]")
        sys.exit(1)

    target = sys.argv[1]
    os.makedirs(CACHE_DIR, exist_ok=True)

    if os.path.isdir(target):
        docx_files = sorted(Path(target).glob("*.docx"))
        for f in docx_files:
            doc_id = f.stem
            cache_path = os.path.join(CACHE_DIR, f"{doc_id}.json")
            if os.path.exists(cache_path) and "--force" not in sys.argv:
                print(f"  [cached] {doc_id}")
                continue
            print(f"  Extracting {doc_id}...")
            try:
                data = extract_dml(str(f))
                with open(cache_path, "w", encoding="utf-8") as out:
                    json.dump(data, out, ensure_ascii=False, indent=2)
                print(f"    {len(data['paragraphs'])} paras, {len(data['tables'])} tables")
            except Exception as e:
                print(f"    [ERROR] {e}")
    else:
        doc_id = Path(target).stem
        cache_path = os.path.join(CACHE_DIR, f"{doc_id}.json")
        print(f"Extracting {doc_id}...")
        data = extract_dml(target)
        with open(cache_path, "w", encoding="utf-8") as out:
            json.dump(data, out, ensure_ascii=False, indent=2)
        print(f"  {len(data['paragraphs'])} paras, {len(data['tables'])} tables -> {cache_path}")


if __name__ == "__main__":
    main()
