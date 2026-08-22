# -*- coding: utf-8 -*-
"""What the corpus asks for, feature by feature, against what the renderer draws.

The `drawn sheet` column is how many workbooks carry the feature on the
sheet the gate compares — conditional formatting is in 16 workbooks and
on the drawn sheet of 2, which is the number that decides what to build.

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
    # `textRotation="0"` is no rotation at all and says nothing about the
    # sheet; only 1..180 turn the text.
    ("turned text (1 to 180 degrees)", "styles", r'textRotation="(?!0"|255)\d', "NOT DRAWN"),
    ("distributed alignment", "styles", r'horizontal="distributed"', "drawn"),
    ("centre across the selection", "styles", r'horizontal="centerContinuous"', "drawn"),
    ("justified alignment", "styles", r'horizontal="justify"', "as distributed"),
    ("raised or lowered runs", "shared", r"<vertAlign ", "drawn, row height NOT"),
    ("phonetic guides (ruby)", "shared", r"<rPh ", "text only, guides NOT DRAWN"),
    ("rich text runs", "shared", r"</r><r>", "drawn"),
    ("solid fills", "styles", r'patternType="solid"', "drawn"),
    # Counted by hand: Excel writes `<patternFill patternType="gray125"/>` as
    # fill 1 of every workbook whether or not a cell asks for it, so counting
    # the element says all 285 when the honest number is 2.
    ("pattern fills (not solid)", "patterned cells", None, "NOT DRAWN"),
    ("gradient fills", "styles", r"<gradientFill", "NOT DRAWN"),
    ("diagonal borders", "styles", r"<diagonal (?!/)", "drawn"),
    ("number formats", "styles", r"<numFmt ", "drawn"),
    # What hangs over the grid
    ("drawings", "sheet", r"<drawing ", "pictures, lines, boxes, diamonds"),
    ("charts", "parts", r"xl/charts/chart\d+\.xml", "line charts drawn"),
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


def patterned_cells(names, read):
    """How many cells wear a fill that is neither solid nor plain.

    Follows the chain a cell actually walks — cell `s=` to `cellXfs` entry to
    `fillId` to the fill's `patternType` — because a fill that is declared and
    never named paints nothing.
    """
    styles = read("xl/styles.xml")
    listed = re.search(r"<fills.*?</fills>", styles, re.S)
    if not listed:
        return 0
    patterned = {
        index
        for index, fill in enumerate(
            re.findall(r"<fill>.*?</fill>|<fill/>", listed.group(), re.S))
        if (kind := re.search(r'patternType="([^"]+)"', fill))
        and kind.group(1) not in ("solid", "none")
    }
    if not patterned:
        return 0

    body = re.search(r"<cellXfs.*?</cellXfs>", styles, re.S)
    if not body:
        return 0
    wanted = {
        str(index)
        for index, xf in enumerate(
            re.findall(r"<xf[^>]*/>|<xf[^>]*>.*?</xf>", body.group(), re.S))
        if (fill_id := re.search(r'fillId="(\d+)"', xf))
        and int(fill_id.group(1)) in patterned
    }
    if not wanted:
        return 0

    return sum(
        style in wanted
        for name in names if name.startswith("xl/worksheets/sheet")
        for style in re.findall(r'<c [^>]*s="(\d+)"', read(name))
    )


def main():
    books = collections.Counter()
    uses = collections.Counter()
    # How many carry it on the sheet the gate actually draws.
    drawn = collections.Counter()
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
                elif kind == "sheet":
                    bodies[kind] = "\n".join(
                        read(name) for name in names
                        if name.startswith("xl/worksheets/sheet")
                    )
                else:
                    # What the gate sees is one sheet, and a feature on the
                    # ninth sheet of a workbook is one nothing measures.
                    bodies[kind] = read("xl/worksheets/sheet1.xml")
            return bodies[kind]

        def read(name):
            try:
                return zipped.read(name).decode("utf-8", "replace")
            except Exception:
                return ""

        for name, where, pattern, _ in LOOKS:
            if where == "patterned cells":
                found = patterned_cells(names, read)
                if found:
                    books[name] += 1
                    uses[name] += found
                continue
            found = len(re.findall(pattern, body(where)))
            if found:
                books[name] += 1
                uses[name] += found
            if where == "sheet" and re.search(pattern, body("drawn")):
                drawn[name] += 1

    print(f"{total} workbooks\n")
    print(f"{'feature':<34}{'books':>6}{'uses':>8}{'drawn sheet':>12}"
          f"   what the renderer does")
    for name, where, _pattern, doing in LOOKS:
        if not books[name]:
            continue
        seen = str(drawn[name]) if where == "sheet" else ""
        print(f"{name:<34}{books[name]:>6}{uses[name]:>8}{seen:>12}   {doing}")
    print("\nnot found in the corpus at all:")
    for name, _where, _pattern, _doing in LOOKS:
        if not books[name]:
            print(f"  {name}")


if __name__ == "__main__":
    main()
