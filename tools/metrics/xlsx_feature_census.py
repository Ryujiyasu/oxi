# -*- coding: utf-8 -*-
"""What the corpus asks for, feature by feature, against what the renderer draws.

A gate can only find what the corpus exposes, and only where somebody thinks
to look. This lists every part, element and attribute in the 285 workbooks
that changes what a sheet looks like, with how many workbooks ask for it — so
the ones nothing draws yet are visible without having to notice them first.

    python tools\\metrics\\xlsx_feature_census.py
"""
import collections
import re
import sys
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

CORPUS = Path(__file__).resolve().parents[2] / "tools" / "golden-test" / "documents" / "xlsx"

# (name, where to look, pattern, what the renderer does with it today)
LOOKS = [
    # The cell table
    ("merged cells", "sheet", r"<mergeCell ", "drawn"),
    ("wrapped text", "styles", r'wrapText="1"', "drawn"),
    ("shrink to fit", "styles", r'shrinkToFit="1"', "drawn"),
    ("indent", "styles", r'indent="[1-9]', "drawn"),
    ("stacked text (rotation 255)", "styles", r'textRotation="255"', "drawn"),
    ("turned text (other rotations)", "styles", r'textRotation="(?!255)\d', "NOT DRAWN"),
    ("distributed alignment", "styles", r'horizontal="distributed"', "drawn"),
    ("centre across the selection", "styles", r'horizontal="centerContinuous"', "drawn"),
    ("justified alignment", "styles", r'horizontal="justify"', "as distributed"),
    ("raised or lowered runs", "shared", r"<vertAlign ", "drawn, row height NOT"),
    ("phonetic guides (ruby)", "shared", r"<rPh ", "text only, guides NOT DRAWN"),
    ("rich text runs", "shared", r"</r><r>", "drawn"),
    ("solid fills", "styles", r'patternType="solid"', "drawn"),
    ("pattern fills (not solid)", "styles",
     r'patternType="(?!solid|none)[a-zA-Z0-9]+"', "NOT DRAWN"),
    ("gradient fills", "styles", r"<gradientFill", "NOT DRAWN"),
    ("diagonal borders", "styles", r"<diagonal (?!/)", "NOT DRAWN"),
    ("number formats", "styles", r"<numFmt ", "drawn"),
    # What hangs over the grid
    ("drawings", "sheet", r"<drawing ", "pictures, lines, boxes, diamonds"),
    ("charts", "parts", r"xl/charts/chart\d+\.xml", "NOT DRAWN"),
    ("pictures", "parts", r"xl/media/", "drawn"),
    ("notes (comments)", "parts", r"xl/comments\d+\.xml", "drawn when pinned open"),
    ("threaded comments", "parts", r"xl/threadedComments/", "NOT DRAWN"),
    ("form controls", "parts", r"xl/ctrlProps/", "NOT DRAWN"),
    ("OLE objects", "parts", r"xl/embeddings/", "NOT DRAWN"),
    ("slicers", "parts", r"xl/slicers/", "NOT DRAWN"),
    ("pivot tables", "parts", r"xl/pivotTables/", "NOT DRAWN"),
    ("tables", "parts", r"xl/tables/", "drawn"),
    # The sheet's own settings
    ("conditional formatting", "sheet", r"<conditionalFormatting", "NOT DRAWN"),
    ("data bars", "sheet", r"<dataBar", "NOT DRAWN"),
    ("colour scales", "sheet", r"<colorScale", "NOT DRAWN"),
    ("icon sets", "sheet", r"<iconSet", "NOT DRAWN"),
    ("sparklines", "sheet", r"<x14:sparklineGroup", "NOT DRAWN"),
    ("data validation", "sheet", r"<dataValidation", "no arrow drawn (Excel draws none either)"),
    ("hyperlinks", "sheet", r"<hyperlink ", "styled by the cell's own format"),
    ("auto filter", "sheet", r"<autoFilter", "buttons drawn"),
    ("frozen panes", "sheet", r"<pane ", "no effect on a range picture"),
    ("gridlines shown", "sheet", r'showGridLines="0"', "none drawn either way"),
    ("right to left", "sheet", r'rightToLeft="1"', "NOT DRAWN"),
    ("hidden rows", "sheet", r'<row [^>]*hidden="1"', "drawn as zero height"),
    ("hidden columns", "sheet", r'<col [^>]*hidden="1"', "drawn as zero width"),
    ("outline groups", "sheet", r'outlineLevel="[1-9]', "no gutter drawn"),
    ("cell errors", "sheet", r't="e"', "drawn as the error text"),
]


def main():
    books = collections.Counter()
    uses = collections.Counter()
    total = 0
    for path in sorted(CORPUS.glob("*.xlsx")):
        try:
            zipped = zipfile.ZipFile(path)
        except Exception:
            continue
        total += 1
        names = zipped.namelist()
        bodies = {}

        def body(kind):
            if kind not in bodies:
                if kind == "parts":
                    bodies[kind] = "\n".join(names)
                elif kind == "styles":
                    bodies[kind] = read("xl/styles.xml")
                elif kind == "shared":
                    bodies[kind] = read("xl/sharedStrings.xml")
                else:
                    bodies[kind] = "\n".join(
                        read(name) for name in names
                        if name.startswith("xl/worksheets/sheet")
                    )
            return bodies[kind]

        def read(name):
            try:
                return zipped.read(name).decode("utf-8", "replace")
            except Exception:
                return ""

        for name, where, pattern, _ in LOOKS:
            found = len(re.findall(pattern, body(where)))
            if found:
                books[name] += 1
                uses[name] += found

    print(f"{total} workbooks\n")
    print(f"{'feature':<34}{'books':>6}{'uses':>8}   what the renderer does")
    for name, _where, _pattern, doing in LOOKS:
        if not books[name]:
            continue
        print(f"{name:<34}{books[name]:>6}{uses[name]:>8}   {doing}")
    print("\nnot found in the corpus at all:")
    for name, _where, _pattern, _doing in LOOKS:
        if not books[name]:
            print(f"  {name}")


if __name__ == "__main__":
    main()
