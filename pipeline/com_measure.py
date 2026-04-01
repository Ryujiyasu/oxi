"""COM API measurement for Word rendering specification derivation.

Given a docx file and a target area (page, region), measures Word's actual
rendering behavior via COM API to establish ground truth values.
"""

import os
import json
import sys
import win32com.client
from pathlib import Path
from .config import DATA_DIR

MEASUREMENTS_DIR = os.path.join(DATA_DIR, "com_measurements")


def measure_line_heights(docx_path: str) -> list[dict]:
    """Measure actual line heights for every paragraph via COM."""
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    results = []

    try:
        doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
        try:
            for i, para in enumerate(doc.Paragraphs):
                rng = para.Range
                # wdFirstCharacterLineNumber = 10
                line_num = rng.Information(10)
                # wdVerticalPositionRelativeToPage = 6
                y_pos = rng.Information(6)  # points from top of page
                # wdActiveEndPageNumber = 3
                page = rng.Information(3)

                # Line spacing
                fmt = para.Format
                line_spacing = fmt.LineSpacing
                line_spacing_rule = fmt.LineSpacingRule
                space_before = fmt.SpaceBefore
                space_after = fmt.SpaceAfter

                # Font info from first run
                font_name = ""
                font_size = 0
                if para.Range.Characters.Count > 0:
                    font = para.Range.Characters(1).Font
                    font_name = font.Name
                    font_size = font.Size

                results.append({
                    "para_index": i,
                    "page": page,
                    "y_position_pt": round(y_pos, 2),
                    "line_number": line_num,
                    "line_spacing_pt": round(line_spacing, 2),
                    "line_spacing_rule": line_spacing_rule,
                    "space_before_pt": round(space_before, 2),
                    "space_after_pt": round(space_after, 2),
                    "font_name": font_name,
                    "font_size_pt": round(font_size, 2),
                    "text_preview": rng.Text[:50].strip(),
                })
        finally:
            doc.Close(SaveChanges=False)
    finally:
        word.Quit()

    return results


def measure_table_widths(docx_path: str) -> list[dict]:
    """Measure actual table column widths via COM."""
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    results = []

    try:
        doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
        try:
            for t, table in enumerate(doc.Tables):
                table_result = {
                    "table_index": t,
                    "rows": table.Rows.Count,
                    "cols": table.Columns.Count,
                    "columns": [],
                    "row_heights": [],
                }
                try:
                    for c in range(1, table.Columns.Count + 1):
                        table_result["columns"].append({
                            "index": c,
                            "width_pt": round(table.Columns(c).Width, 2),
                        })
                except Exception:
                    pass  # Merged cells can cause column access errors

                try:
                    for r in range(1, table.Rows.Count + 1):
                        table_result["row_heights"].append({
                            "index": r,
                            "height_pt": round(table.Rows(r).Height, 2),
                            "height_rule": table.Rows(r).HeightRule,
                        })
                except Exception:
                    pass

                results.append(table_result)
        finally:
            doc.Close(SaveChanges=False)
    finally:
        word.Quit()

    return results


def measure_page_breaks(docx_path: str) -> list[dict]:
    """Measure which paragraph starts each page."""
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    results = []

    try:
        doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
        try:
            current_page = 0
            for i, para in enumerate(doc.Paragraphs):
                page = para.Range.Information(3)  # wdActiveEndPageNumber
                if page != current_page:
                    y_pos = para.Range.Information(6)
                    results.append({
                        "page": page,
                        "first_para_index": i,
                        "y_position_pt": round(y_pos, 2),
                        "text_preview": para.Range.Text[:50].strip(),
                    })
                    current_page = page
        finally:
            doc.Close(SaveChanges=False)
    finally:
        word.Quit()

    return results


def measure_document(docx_path: str) -> dict:
    """Full measurement of a document."""
    doc_id = Path(docx_path).stem
    Path(MEASUREMENTS_DIR).mkdir(parents=True, exist_ok=True)

    print(f"  COM: measuring {doc_id}...")
    result = {
        "doc_id": doc_id,
        "docx_path": docx_path,
    }

    try:
        result["line_heights"] = measure_line_heights(docx_path)
        print(f"    line_heights: {len(result['line_heights'])} paragraphs")
    except Exception as e:
        print(f"    line_heights: FAILED ({e})")
        result["line_heights"] = []

    try:
        result["table_widths"] = measure_table_widths(docx_path)
        print(f"    table_widths: {len(result['table_widths'])} tables")
    except Exception as e:
        print(f"    table_widths: FAILED ({e})")
        result["table_widths"] = []

    try:
        result["page_breaks"] = measure_page_breaks(docx_path)
        print(f"    page_breaks: {len(result['page_breaks'])} pages")
    except Exception as e:
        print(f"    page_breaks: FAILED ({e})")
        result["page_breaks"] = []

    # Save
    out_path = os.path.join(MEASUREMENTS_DIR, f"{doc_id}.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"  COM: saved {out_path}")
    return result


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python -m pipeline.com_measure <docx_path>")
        sys.exit(1)
    measure_document(sys.argv[1])
