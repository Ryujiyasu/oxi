# -*- coding: utf-8 -*-
"""Two-column probes with a TABLE inside the run.

The plain probes in `_col_split_gen.py` pinned where Word breaks a run of
text lines. They say nothing about a run holding a table, and that turns out
to be the shape this engine gets wrong: it keeps a table out of the balance
entirely, either pinning it above the split or refusing to balance at all,
while Word counts a table's rows as lines of the column and will break the
table across the column boundary.

    python tools/metrics/_col_table_gen.py [out-dir]

Sweeps the table's SIZE and its POSITION among the text lines, because those
are the two things a rule could key on and one of them alone cannot tell them
apart. Measure with `_col_split_lines.py` (Word, line by line) and
`_col_table_sweep.py` (this engine).
"""
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _col_split_gen as g  # noqa: E402

LINES = 8
ROWS = [1, 2, 3, 5]
AT = [0, 2, 4, 6, 8]

PGSZ = ('<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" '
        'w:header="0" w:footer="0" w:gutter="0"/>')


def cell(text: str) -> str:
    return ('<w:tc><w:tcPr><w:tcW w:w="2300" w:type="dxa"/></w:tcPr>'
            + g.para(text, 20) + '</w:tc>')


def table(rows: int) -> str:
    """A plain bordered two-column table, one line per cell."""
    trs = "".join('<w:tr>' + cell(f"r{i + 1}c1") + cell(f"r{i + 1}c2") + '</w:tr>'
                  for i in range(rows))
    borders = "".join(f'<w:{side} w:val="single" w:sz="4" w:color="000000"/>'
                      for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
    return ('<w:tbl><w:tblPr><w:tblW w:w="4600" w:type="dxa"/>'
            f'<w:tblBorders>{borders}</w:tblBorders></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="2300"/><w:gridCol w:w="2300"/></w:tblGrid>'
            + trs + '</w:tbl>')


def document(rows: int, at: int) -> str:
    parts = []
    for i in range(LINES):
        if i == at:
            parts.append(table(rows))
        parts.append(g.para(f"L{i + 1}", 20))
    if at >= LINES:
        parts.append(table(rows))
    body = "".join(parts)
    body += ('<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/>'
             f'<w:cols w:num="2" w:space="360" w:equalWidth="1"/>{PGSZ}'
             '</w:sectPr></w:pPr></w:p>')
    body += g.para("after", 20)
    body += (f'<w:sectPr><w:type w:val="continuous"/><w:cols w:num="1" w:space="360"/>'
             f'{PGSZ}</w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{body}</w:body></w:document>')


def write(path: Path, rows: int, at: int) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", g.CONTENT_TYPES)
        z.writestr("_rels/.rels", g.ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", g.DOC_RELS)
        z.writestr("word/settings.xml", g.settings(15))
        z.writestr("word/document.xml", document(rows, at))


def main() -> int:
    where = Path(sys.argv[1] if len(sys.argv) > 1 else "tests/fixtures/column_split")
    where.mkdir(parents=True, exist_ok=True)
    made = 0
    for rows in ROWS:
        for at in AT:
            write(where / f"tab_{rows}rows_at{at}_compat15.docx", rows, at)
            made += 1
    print(f"{made} documents in {where}")
    print(f"each is {LINES} lines with a table of 1-5 rows dropped in among them")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
