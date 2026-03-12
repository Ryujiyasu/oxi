"""
Extract exact layout positions from a .docx file using Word COM API.

Usage: python word_layout_extract.py <docx_path>

Outputs TSV with columns:
  para_idx, page, y_pt, font_name, font_size, space_before, space_after,
  line_spacing, line_spacing_rule, line_count, text_preview
"""

import win32com.client
import sys
import os

# Force UTF-8 output
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# Word constants
wdVerticalPositionRelativeToPage = 6
wdHorizontalPositionRelativeToPage = 5
wdActiveEndPageNumber = 3
wdFirstCharacterLineNumber = 10
wdLineSpacingSingle = 0
wdLineSpacing1pt5 = 1
wdLineSpaceDouble = 2
wdLineSpaceAtLeast = 3
wdLineSpaceExactly = 4
wdLineSpaceMultiple = 5

LINE_RULE_NAMES = {
    0: "Single",
    1: "1.5",
    2: "Double",
    3: "AtLeast",
    4: "Exactly",
    5: "Multiple",
}


def extract_layout(docx_path):
    docx_path = os.path.abspath(docx_path)
    if not os.path.exists(docx_path):
        print(f"ERROR: file not found: {docx_path}", file=sys.stderr)
        sys.exit(1)

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False

    try:
        doc = word.Documents.Open(docx_path, ReadOnly=True)

        # Print header
        print("idx\tpage\ty_pt\tx_pt\tfont_name\tfont_size\tspace_before\tspace_after\tline_spacing\tline_rule\ttext")

        para_count = doc.Paragraphs.Count
        for i in range(1, para_count + 1):  # COM is 1-indexed
            para = doc.Paragraphs(i)
            rng = para.Range

            try:
                top = rng.Information(wdVerticalPositionRelativeToPage)
                left = rng.Information(wdHorizontalPositionRelativeToPage)
                page = rng.Information(wdActiveEndPageNumber)
            except Exception:
                top = -1
                left = -1
                page = -1

            font_name = rng.Font.Name or ""
            font_size = rng.Font.Size

            sb = para.SpaceBefore
            sa = para.SpaceAfter
            ls = para.LineSpacing
            lr = para.LineSpacingRule

            lr_name = LINE_RULE_NAMES.get(int(lr), str(lr))

            text_preview = rng.Text[:50].replace("\r", "\\r").replace("\n", "\\n").replace("\t", "\\t").strip()

            print(f"{i-1}\t{page}\t{top:.2f}\t{left:.2f}\t{font_name}\t{font_size:.1f}\t{sb:.2f}\t{sa:.2f}\t{ls:.2f}\t{lr_name}\t{text_preview}")

        doc.Close(False)
    finally:
        word.Quit()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python word_layout_extract.py <docx_path>", file=sys.stderr)
        sys.exit(1)
    extract_layout(sys.argv[1])
