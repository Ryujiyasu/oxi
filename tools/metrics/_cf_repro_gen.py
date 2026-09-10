# -*- coding: utf-8 -*-
"""Writes workbooks whose conditional rules certainly fire.

The 285-workbook corpus turned out to be a poor place to ask whether Oxi
applies a conditional rule the way Excel does: 3288 of its 3342 rule ranges are
`containsText` looking for a tick mark on a blank government form, so the
correct answer for almost all of them is "nothing is formatted". A test that
can only be passed by doing nothing is not a test.

So this authors the other half: one sheet per rule kind, with values chosen so
that some cells match and some do not, and a distinct fill on each rule so a
wrong rule cannot accidentally produce a right colour. Excel is then asked what
it shows through `Range.DisplayFormat`, and Oxi through
`examples/_conditional_dump`.

    python tools/metrics/_cf_repro_gen.py tests/fixtures/conditional

Written by hand rather than through a library so the file states exactly the
rules under test and nothing else.
"""
import sys
import zipfile
from pathlib import Path

# Three looks, far enough apart that no two can be confused in a screenshot or
# a colour comparison: red fill, green fill, bold blue text on yellow.
# A dxf states its parts in one order and no other -- font, numFmt, fill,
# alignment, border, protection. Excel refuses to open a file that puts the fill
# first, which is how the first draft of this generator was caught.
DXFS = """<dxfs count="5">
<dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>
<dxf><font><color rgb="FF006100"/></font><fill><patternFill><bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf>
<dxf><font><b/><color rgb="FF0000FF"/></font><fill><patternFill><bgColor rgb="FFFFEB9C"/></patternFill></fill></dxf>
<dxf><font><b/></font></dxf>
<dxf><fill><patternFill><bgColor rgb="FFBDD7EE"/></patternFill></fill></dxf>
</dxfs>"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
""" + DXFS + "</styleSheet>"

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

BOOK_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

WORKBOOK = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"""


def column(index):
    name = ""
    n = index
    while True:
        name = chr(65 + n % 26) + name
        n = n // 26 - 1
        if n < 0:
            return name


def cell(ref, value):
    """A cell holding either a number or an inline string."""
    if isinstance(value, (int, float)):
        return '<c r="%s"><v>%s</v></c>' % (ref, value)
    text = str(value).replace("&", "&amp;").replace("<", "&lt;")
    return '<c r="%s" t="inlineStr"><is><t>%s</t></is></c>' % (ref, text)


def sheet(values, rules):
    """`values` is a list of rows; `rules` a list of (sqref, cfRule xml)."""
    rows = []
    for at, line in enumerate(values):
        cells = "".join(cell("%s%d" % (column(col), at + 1), held)
                        for col, held in enumerate(line) if held is not None)
        rows.append('<row r="%d">%s</row>' % (at + 1, cells))
    blocks = "".join('<conditionalFormatting sqref="%s">%s</conditionalFormatting>' % pair
                     for pair in rules)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetData>%s</sheetData>%s</worksheet>' % ("".join(rows), blocks))


def write(path, values, rules):
    body = sheet(values, rules)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("xl/workbook.xml", WORKBOOK)
        z.writestr("xl/_rels/workbook.xml.rels", BOOK_RELS)
        z.writestr("xl/styles.xml", STYLES)
        z.writestr("xl/worksheets/sheet1.xml", body)
    print("wrote %s" % path)


def rule(kind, dxf, priority, **rest):
    bits = ['type="%s"' % kind, 'dxfId="%d"' % dxf, 'priority="%d"' % priority]
    for name in ("operator", "text", "rank", "bottom", "percent", "stopIfTrue"):
        if name in rest:
            bits.append('%s="%s"' % (name, rest.pop(name)))
    formulas = "".join("<formula>%s</formula>" % f for f in rest.pop("formulas", []))
    assert not rest, rest
    return "<cfRule %s>%s</cfRule>" % (" ".join(bits), formulas)


BOOKS = {
    # A number against a bound, the commonest rule of all.
    "cell_is.xlsx": (
        [["value"], [10], [50], [75], [100], [3], [None], ["50"]],
        [("A2:A8", rule("cellIs", 0, 1, operator="greaterThan", formulas=["50"])
                  + rule("cellIs", 1, 2, operator="between", formulas=["5", "20"]))],
    ),
    # Text tests, including the case a blank cell answers.
    "text_rules.xlsx": (
        [["word"], ["apple"], ["Pineapple"], ["banana"], ["APPLESAUCE"], [None], ["cape"]],
        [("A2:A7", rule("containsText", 0, 1, operator="containsText", text="apple",
                        formulas=['NOT(ISERROR(SEARCH("apple",A2)))'])
                   + rule("beginsWith", 1, 2, operator="beginsWith", text="ba",
                          formulas=['LEFT(A2,2)="ba"'])
                   + rule("endsWith", 2, 3, operator="endsWith", text="pe",
                          formulas=['RIGHT(A2,2)="pe"']))],
    ),
    # Blank and not-blank over the same range, so every cell must be caught by
    # exactly one of them.
    "blanks.xlsx": (
        [["state"], ["here"], [None], [""], [7], [None]],
        [("A2:A6", rule("containsBlanks", 0, 1, formulas=['LEN(TRIM(A2))=0'])
                   + rule("notContainsBlanks", 1, 2, formulas=['LEN(TRIM(A2))>0']))],
    ),
    # Counting rules: what fires depends on the whole range, not one cell.
    "duplicates.xlsx": (
        [["name"], ["red"], ["blue"], ["red"], ["green"], ["blue"], ["red"]],
        [("A2:A7", rule("duplicateValues", 0, 1)),
         ("A2:A7", rule("uniqueValues", 1, 2))],
    ),
    # A formula written for the top-left cell and moved to each of the others,
    # including one that reaches out to a fixed cell.
    "expression.xlsx": (
        [["n", "flag"], [4, 10], [11, 10], [40, 10], [9, 10]],
        [("A2:A5", rule("expression", 0, 1, formulas=["A2>$B$2"])),
         ("B2:B5", rule("expression", 2, 2, formulas=["MOD(ROW(),2)=0"]))],
    ),
    # Two rules over one range, the higher-precedence one setting only weight
    # and the lower only fill. Excel does not pick a winner: it lays them over
    # each other and each rule contributes what the one above it left alone.
    "layered.xlsx": (
        [["n"], [1], [2], [3], [4], [5], [6]],
        [("A2:A7", rule("cellIs", 3, 1, operator="greaterThan", formulas=["3"])
                   + rule("cellIs", 4, 2, operator="greaterThan", formulas=["1"]))],
    ),
    # The same shape, but the first rule stops the ones under it.
    "stop_if_true.xlsx": (
        [["n"], [1], [2], [3], [4], [5], [6]],
        [("A2:A7", rule("cellIs", 0, 1, operator="greaterThan", formulas=["4"],
                        stopIfTrue="1")
                   + rule("cellIs", 1, 2, operator="greaterThan", formulas=["1"]))],
    ),
    # The rank rules, which Oxi does not claim to run: this is here so the
    # comparison states the gap rather than leaving it undiscovered.
    "top10.xlsx": (
        [["score", "share"], [12, 5], [55, 90], [3, 40], [98, 70], [40, 20], [77, 60]],
        [("A2:A7", rule("top10", 0, 1, rank="2")
                   + rule("top10", 1, 2, rank="2", bottom="1")),
         # A share rather than a count: 30% of six numbers. Whether Excel
         # rounds that share up, down, or to the nearest is not something to
         # guess at, so it is measured.
         ("B2:B7", rule("top10", 2, 3, rank="30", percent="1"))],
    ),
}


def main():
    where = Path(sys.argv[1] if len(sys.argv) > 1 else "tests/fixtures/conditional")
    where.mkdir(parents=True, exist_ok=True)
    for name, (values, rules) in BOOKS.items():
        write(where / name, values, rules)
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
