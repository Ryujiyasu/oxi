# -*- coding: utf-8 -*-
"""Where does Word break a two-column run?

The engine decides this with a rule read off Word: ten equal lines split five
and five, but a final line tall enough to need two grid slots pushes the split
to six. That rule passes every document it was written for and gets a page of
two-column exercises wrong on every page, so the condition it keys on is not
the condition Word keys on.

This asks Word directly. Each document is one two-column section holding N
lines, every line the same height except the last, which is swept. Word then
says where it put the break, and the boundary between "five left" and "six
left" can be read off rather than guessed.

    python tools/metrics/_col_split_gen.py <out-dir>

Writes one .docx per (line count, last-line size, compat mode). Measure them
with `_col_split_word.py`.
"""
import sys
import zipfile
from pathlib import Path

# Half-points, because that is how a run states its size. The base line is 10pt.
BASE_HALF = 20
# Between 10pt and 11pt Word changes its mind, so the sweep is finest there.
LAST_HALF = [20, 21, 22, 23, 24]
COUNTS = [6, 8, 10, 12, 14]
COMPAT = [15]

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""


def settings(compat: int) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:compat><w:compatSetting w:name="compatibilityMode" '
        f'w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/></w:compat>'
        '</w:settings>')


def para(text: str, half: int) -> str:
    """One line, at a stated size, with nothing else on it."""
    return (
        '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" '
        'w:lineRule="auto"/><w:rPr>'
        f'<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="{half}"/>'
        '</w:rPr></w:pPr><w:r><w:rPr>'
        f'<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="{half}"/>'
        f'</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p>')


def document(count: int, last_half: int, tall_at: int = -1) -> str:
    """`tall_at` is which line carries `last_half`; -1 means the last one.

    Sweeping the position as well as the size is the only way to tell what
    "tall" is measured against — the line before it, the average, or the grid.
    A rule that only ever saw a tall LAST line cannot distinguish those.
    """
    lines = []
    where = count - 1 if tall_at < 0 else tall_at
    for i in range(count):
        half = last_half if i == where else BASE_HALF
        lines.append(para(f"L{i + 1}", half))
    # One two-column section, evenly spaced, on a page big enough that the
    # whole run fits on one page — the split is then purely the balancer's.
    # A two-column run that is BALANCED, not filled: Word only balances a
    # column run when the section ends, so the two-column part is its own
    # section closed by a continuous break, with a one-column section after it.
    # Without that, Word fills the left column to the page bottom and the
    # question being asked here never arises.
    two_col = (
        '<w:p><w:pPr><w:sectPr>'
        '<w:type w:val="continuous"/>'
        '<w:cols w:num="2" w:space="360" w:equalWidth="1"/>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" '
        'w:header="0" w:footer="0" w:gutter="0"/>'
        '</w:sectPr></w:pPr></w:p>')
    # The one-column section needs real content: the engine only notices a
    # column run ending when something flows into the next one, so a section
    # that holds nothing leaves the two-column run unbalanced and the question
    # unanswered. Word balances either way, which is what made the empty
    # version look like an engine bug rather than a probe bug.
    after = para("after", BASE_HALF)
    body = "".join(lines) + two_col + after + (
        '<w:sectPr><w:type w:val="continuous"/>'
        '<w:cols w:num="1" w:space="360"/>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" '
        'w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{body}</w:body></w:document>')


def prefixed(count: int, prefix_lines: int, compat: int) -> str:
    """The same run, but with full-width lines above it on the same page.

    A column run rarely starts at the top of a page — a heading or a lead
    paragraph usually sits above it, and the engine takes a different arm when
    it does. Whether the height Word halves is the RUN or what is left of the
    PAGE cannot be read off a run that starts at the top, because there the
    two are the same.
    """
    before = "".join(para(f"P{i + 1}", BASE_HALF) for i in range(prefix_lines))
    one_col = (
        '<w:p><w:pPr><w:sectPr>'
        '<w:type w:val="continuous"/>'
        '<w:cols w:num="1" w:space="360"/>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" '
        'w:header="0" w:footer="0" w:gutter="0"/>'
        '</w:sectPr></w:pPr></w:p>')
    body = document(count, BASE_HALF)
    at = body.index("<w:body>") + len("<w:body>")
    return body[:at] + before + one_col + body[at:]


def write(path: Path, count: int, last_half: int, compat: int, tall_at: int = -1,
          prefix_lines: int = 0) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/settings.xml", settings(compat))
        z.writestr("word/document.xml",
                   prefixed(count, prefix_lines, compat) if prefix_lines
                   else document(count, last_half, tall_at))


def main() -> int:
    where = Path(sys.argv[1] if len(sys.argv) > 1 else "tests/fixtures/column_split")
    where.mkdir(parents=True, exist_ok=True)
    made = 0
    for compat in COMPAT:
        for count in COUNTS:
            for half in LAST_HALF:
                name = f"col_{count}lines_last{half}hp_compat{compat}.docx"
                write(where / name, count, half, compat)
                made += 1
    # The same tall line, moved: if Word is looking at the last line
    # specifically, only the last position changes the split.
    for count in (8, 10, 12):
        for half in (24,):
            for at in range(count):
                name = f"pos_{count}lines_tall{half}hp_at{at}_compat15.docx"
                write(where / name, count, half, 15, at)
                made += 1
    # A run that does NOT start at the top of the page.
    for count in (8, 10):
        for lines in (1, 2, 3, 5):
            write(where / f"pre_{count}lines_prefix{lines}_compat15.docx",
                  count, BASE_HALF, 15, prefix_lines=lines)
            made += 1
    print(f"{made} documents in {where}")
    print("each is one two-column section: N lines at 10pt, the last one swept")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
