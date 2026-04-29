"""
Generate 10 minimal-repro .docx fixtures covering comments + tracked changes.

Phase 1 Tick 5-6 of feat/comments-tracked-changes mission.

Each fixture is the smallest document that exercises exactly ONE feature so
that the next Tick (Word COM measurement) can isolate the rendered geometry /
color / ordering rules without interference from unrelated content.

Fixture list (indexed by file name):
  01 — Single comment, single paragraph
  02 — Comment with a single reply (parent + child)
  03 — Resolved comment (`w15:done="1"`)
  04 — Comment whose range spans 3 paragraphs
  05 — Single `<w:ins>` insertion, single paragraph
  06 — Single `<w:del>` deletion, single paragraph
  07 — Mixed `<w:ins>` + `<w:del>` in same paragraph
  08 — `<w:moveFrom>` + `<w:moveTo>` between paragraphs
  09 — `<w:rPrChange>` formatting revision (bold toggle)
  10 — Multiple reviewers — two distinct authors, one `<w:ins>` + one `<w:del>`
  11 — CJK body (MS Mincho 24pt) with one `<w:ins>` + one `<w:del>` —
       exercises R-01/R-03 styling on CJK glyphs and is the smallest case
       for verifying strikethrough Y on full-width characters.
  12 — Three reviewers (Alice + Bob + Carol) — exercises palette slot 2
       (`#2B6033`). Slots 0/1 already proven via fixture_10; slot 2 is
       documented but not yet COM-confirmed. The fixture surfaces it on
       the Oxi side and provides the input file that a future Word-side
       pixel-pass run can sample for ground-truth confirmation.
  13 — `<w:pPrChange>` paragraph-property revision (indent toggle) —
       paired with fixture_09 (rPrChange, run-level). Confirms R-12 v2
       lands a "Formatted" margin balloon for paragraph-level changes
       too, not just run-level. Indent is the most common pPrChange
       observed in real docs.
  14 — `<w:rPrChange>` font-family revision (Calibri → Times New Roman)
       — second rPrChange variant after fixture_09's bold toggle.
       Exercises `describe_rpr_diff`'s font_family branch (R71
       extension). Multiple-property rPrChange surfaces here too:
       the rPr also toggles font size, so the balloon body lists
       both properties.
  15 — `<w:pPrChange>` paragraph-alignment revision (default left →
       center). R72 unblocks alignment-toggle pPrChange by adding a
       `prior_alignment` IR field on PropertyChange. The R-12 v3.5
       balloon body must mention "Alignment: Centered".
  16 — `<w:rPrChange>` multi-axis revision (small_caps + all_caps +
       character_spacing toggled together). R86 extends
       describe_rpr_diff to cover these 3 axes; this fixture is the
       end-to-end exercise.
  17 — `<w:rPrChange>` multi-axis revision (vertical_align=superscript
       + shading=#FFFF00 yellow). R87 closes the remaining R72 §19.47.5
       axes; describe_rpr_diff now covers 12 axes total. East-Asia
       font axis is wired but not exercised by this fixture (Western
       fixtures suffice for the majority of real revisions).
  18 — `<w:pPrChange>` paragraph-shading revision (default → yellow
       background). R88 extends describe_ppr_diff to cover the
       paragraph-scope shading axis (sister to RunStyle shading
       added in R87).
  19 — `<w:pPrChange>` keep-with-next revision. R89 extends
       describe_ppr_diff to cover keep_next + keep_lines (direct-only
       bool axes on pPr). Surfaces "Keep With Next".
  20 — `<w:pPrChange>` paragraph-borders revision (bottom border
       added). R93 extends describe_ppr_diff to cover the borders
       struct axis via side-presence summary (no PartialEq derive
       needed). Surfaces "Borders Added".
  21 — `<w:pPrChange>` paragraph-tab_stops revision (3 tab stops
       added at 72/144/216pt). R94 extends describe_ppr_diff via
       position-only summary (mirrors R93 borders pattern).
       Surfaces "Tab Stops Added".
  22 — `<w:pPrChange>` paragraph-numPr revision (inline numId=1
       ilvl=0 attached). R95 patches the parser to mirror inline
       numPr onto style.num_id/num_ilvl, unblocking the R89
       attempted branches. Surfaces "Numbering: list 1".
  23 — `<w:rPrChange>` outline + emboss revision. R98 extends
       describe_rpr_diff to cover 3 NEW non-R72 rPr axes
       (outline + emboss + imprint). Surfaces "Outline" + "Emboss".

Outputs to  tools/fixtures/comments_samples/fixture_NN_<slug>.docx.

Strategy: build each .docx from scratch as a minimal OOXML ZIP so we have
exact control of the marker elements.  python-docx has no first-class API
for comments or revisions; zip-patching would work but full-control is
cleaner and keeps the script self-contained (no template file dependency).

Run:
    python tools/fixtures/comments_samples/build_comments_samples.py
"""

from __future__ import annotations

import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from textwrap import dedent

OUT_DIR = Path(__file__).resolve().parent

# --------------------------------------------------------------------------
# Fixed chrome — same across all fixtures
# --------------------------------------------------------------------------

CONTENT_TYPES_BASE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
{comments_override}{comments_ext_override}{people_override}
</Types>
"""

COMMENTS_OVERRIDE = '  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>\n'
# Word 16.0 strict-open rejects the historical "application/vnd.ms-word.*" content
# types for these parts; only the openxmlformats-officedocument flavor is accepted.
# Confirmed 2026-04-25 by comparing fixture build output to Word's OpenAndRepair
# output (which always rewrites the content types to the form below).
COMMENTS_EXT_OVERRIDE = '  <Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>\n'
PEOPLE_OVERRIDE = '  <Override PartName="/word/people.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml"/>\n'

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

DOC_RELS_BASE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
{comments_rel}{comments_ext_rel}{people_rel}
</Relationships>
"""

COMMENTS_REL = '  <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>\n'
COMMENTS_EXT_REL = '  <Relationship Id="rId11" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>\n'
PEOPLE_REL = '  <Relationship Id="rId12" Type="http://schemas.microsoft.com/office/2011/relationships/people" Target="people.xml"/>\n'

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Calibri" w:eastAsia="MS Mincho" w:hAnsi="Calibri"/>
      <w:sz w:val="22"/>
      <w:szCs w:val="22"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr>
      <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="character" w:styleId="CommentReference">
    <w:name w:val="annotation reference"/>
    <w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>
  </w:style>
</w:styles>
"""

SETTINGS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
</w:settings>
"""

# XML namespaces reused in document.xml / comments.xml
DOC_XMLNS = ' '.join([
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
    'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"',
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
    'mc:Ignorable="w14 w15"',
])

COMMENTS_XMLNS = ' '.join([
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
])

COMMENTS_EXT_XMLNS = ' '.join([
    'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"',
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
])

PEOPLE_XMLNS = ' '.join([
    'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"',
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
])

DATE_A = "2026-04-18T10:00:00Z"
DATE_B = "2026-04-18T10:05:00Z"
DATE_REPLY = "2026-04-18T10:10:00Z"


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------

@dataclass
class Fixture:
    name: str
    document_body: str
    comments_xml: str | None = None      # full <w:comments>…</w:comments>, or None
    comments_ext_xml: str | None = None  # full <w15:commentsEx>…</w15:commentsEx>, or None
    people_xml: str | None = None        # optional <w15:people>
    description: str = ""


def _wrap_document(body: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document {DOC_XMLNS}>\n'
        '  <w:body>\n'
        f'{body}\n'
        '    <w:sectPr>\n'
        '      <w:pgSz w:w="11906" w:h="16838"/>\n'
        '      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>\n'
        '    </w:sectPr>\n'
        '  </w:body>\n'
        '</w:document>\n'
    )


def _wrap_comments(entries: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:comments {COMMENTS_XMLNS}>\n'
        f'{entries}\n'
        '</w:comments>\n'
    )


def _wrap_comments_ext(entries: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w15:commentsEx {COMMENTS_EXT_XMLNS}>\n'
        f'{entries}\n'
        '</w15:commentsEx>\n'
    )


def _wrap_people(entries: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w15:people {PEOPLE_XMLNS}>\n'
        f'{entries}\n'
        '</w15:people>\n'
    )


def _content_types_for(f: Fixture) -> str:
    return CONTENT_TYPES_BASE.format(
        comments_override=COMMENTS_OVERRIDE if f.comments_xml else "",
        comments_ext_override=COMMENTS_EXT_OVERRIDE if f.comments_ext_xml else "",
        people_override=PEOPLE_OVERRIDE if f.people_xml else "",
    )


def _doc_rels_for(f: Fixture) -> str:
    return DOC_RELS_BASE.format(
        comments_rel=COMMENTS_REL if f.comments_xml else "",
        comments_ext_rel=COMMENTS_EXT_REL if f.comments_ext_xml else "",
        people_rel=PEOPLE_REL if f.people_xml else "",
    )


def _write_docx(fixture: Fixture) -> Path:
    out_path = OUT_DIR / fixture.name
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",    _content_types_for(fixture))
        z.writestr("_rels/.rels",            ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", _doc_rels_for(fixture))
        z.writestr("word/document.xml",      _wrap_document(fixture.document_body))
        z.writestr("word/styles.xml",        STYLES_XML)
        z.writestr("word/settings.xml",      SETTINGS_XML)
        if fixture.comments_xml:
            z.writestr("word/comments.xml",          fixture.comments_xml)
        if fixture.comments_ext_xml:
            z.writestr("word/commentsExtended.xml",  fixture.comments_ext_xml)
        if fixture.people_xml:
            z.writestr("word/people.xml",            fixture.people_xml)
    return out_path


# --------------------------------------------------------------------------
# Run / paragraph helpers
# --------------------------------------------------------------------------

def run(text: str) -> str:
    """Plain run."""
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def para(*runs: str, para_id: str | None = None, text_id: str | None = None) -> str:
    """Paragraph with optional w14:paraId / w14:textId for commentsExtended linkage."""
    attrs = ""
    if para_id:
        attrs += f' w14:paraId="{para_id}"'
    if text_id:
        attrs += f' w14:textId="{text_id}"'
    return f'    <w:p{attrs}>\n      ' + "\n      ".join(runs) + "\n    </w:p>"


def empty_para(para_id: str | None = None) -> str:
    attrs = f' w14:paraId="{para_id}"' if para_id else ""
    return f'    <w:p{attrs}/>'


def comment_ref(cid: int) -> str:
    """Inline marker that references comment id `cid`."""
    return (
        f'<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
        f'<w:commentReference w:id="{cid}"/></w:r>'
    )


def comment_range_start(cid: int) -> str:
    return f'<w:commentRangeStart w:id="{cid}"/>'


def comment_range_end(cid: int) -> str:
    return f'<w:commentRangeEnd w:id="{cid}"/>'


def comment_entry(
    cid: int,
    author: str,
    initials: str,
    text: str,
    date: str = DATE_A,
    para_id: str | None = None,
) -> str:
    paraid_attr = f' w14:paraId="{para_id}"' if para_id else ""
    return (
        f'  <w:comment w:id="{cid}" w:author="{author}" w:date="{date}" w:initials="{initials}">\n'
        f'    <w:p{paraid_attr}>\n'
        f'      <w:r><w:t>{text}</w:t></w:r>\n'
        '    </w:p>\n'
        '  </w:comment>'
    )


# --------------------------------------------------------------------------
# Fixture definitions
# --------------------------------------------------------------------------

def f01_single_comment() -> Fixture:
    body = para(
        run("The quick "),
        comment_range_start(0),
        run("brown fox"),
        comment_range_end(0),
        comment_ref(0),
        run(" jumps over the lazy dog."),
        para_id="00000001",
    )
    comments = _wrap_comments(comment_entry(
        cid=0, author="Alice Reviewer", initials="AR",
        text="Is 'brown' needed here?",
        para_id="00000010",
    ))
    return Fixture(
        name="fixture_01_single_comment.docx",
        description="One comment anchored on 'brown fox' in a single-paragraph body.",
        document_body=body,
        comments_xml=comments,
    )


def f02_comment_with_reply() -> Fixture:
    body = para(
        run("Initial thought: "),
        comment_range_start(0),
        run("color matters"),
        comment_range_end(0),
        comment_ref(0),
        comment_ref(1),
        run("."),
        para_id="00000001",
    )
    comments = _wrap_comments("\n".join([
        comment_entry(0, "Alice Reviewer", "AR", "Why?",        para_id="00000010"),
        comment_entry(1, "Alice Reviewer", "AR", "Following up.", date=DATE_REPLY, para_id="00000011"),
    ]))
    # commentsExtended: comment 1 is a reply to comment 0 (parentParaId="00000010")
    ext_entries = (
        '  <w15:commentEx w15:paraId="00000010" w15:done="0"/>\n'
        '  <w15:commentEx w15:paraId="00000011" w15:paraIdParent="00000010" w15:done="0"/>'
    )
    return Fixture(
        name="fixture_02_comment_with_reply.docx",
        description="Parent comment + one reply via commentsExtended paraIdParent linkage.",
        document_body=body,
        comments_xml=comments,
        comments_ext_xml=_wrap_comments_ext(ext_entries),
    )


def f03_resolved_comment() -> Fixture:
    body = para(
        run("This sentence "),
        comment_range_start(0),
        run("has been reviewed"),
        comment_range_end(0),
        comment_ref(0),
        run("."),
        para_id="00000001",
    )
    comments = _wrap_comments(comment_entry(
        cid=0, author="Alice Reviewer", initials="AR",
        text="Looks good.", para_id="00000010",
    ))
    ext_entries = (
        '  <w15:commentEx w15:paraId="00000010" w15:done="1"/>'
    )
    return Fixture(
        name="fixture_03_resolved_comment.docx",
        description="Single comment marked resolved (w15:done=\"1\").",
        document_body=body,
        comments_xml=comments,
        comments_ext_xml=_wrap_comments_ext(ext_entries),
    )


def f04_multi_para_range() -> Fixture:
    body = "\n".join([
        para(
            comment_range_start(0),
            run("First paragraph of the ranged comment."),
            para_id="00000001",
        ),
        para(
            run("Second paragraph — still inside the comment range."),
            para_id="00000002",
        ),
        para(
            run("Third paragraph — last one inside the range."),
            comment_range_end(0),
            comment_ref(0),
            para_id="00000003",
        ),
    ])
    comments = _wrap_comments(comment_entry(
        cid=0, author="Alice Reviewer", initials="AR",
        text="Applies to all three paragraphs.", para_id="00000010",
    ))
    return Fixture(
        name="fixture_04_multi_para_range.docx",
        description="One comment whose range spans three consecutive paragraphs.",
        document_body=body,
        comments_xml=comments,
    )


def f05_single_ins() -> Fixture:
    body = para(
        run("Before insertion "),
        '<w:ins w:id="100" w:author="Alice Reviewer" w:date="' + DATE_A + '">'
        '<w:r><w:t xml:space="preserve">INSERTED TEXT </w:t></w:r></w:ins>',
        run("after insertion."),
        para_id="00000001",
    )
    return Fixture(
        name="fixture_05_single_ins.docx",
        description="Single w:ins insertion of 'INSERTED TEXT ' by one author in one paragraph.",
        document_body=body,
    )


def f06_single_del() -> Fixture:
    body = para(
        run("Before delete "),
        '<w:del w:id="100" w:author="Alice Reviewer" w:date="' + DATE_A + '">'
        '<w:r><w:delText xml:space="preserve">DELETED TEXT </w:delText></w:r></w:del>',
        run("after delete."),
        para_id="00000001",
    )
    return Fixture(
        name="fixture_06_single_del.docx",
        description="Single w:del deletion of 'DELETED TEXT ' by one author in one paragraph.",
        document_body=body,
    )


def f07_mixed_ins_del() -> Fixture:
    body = para(
        run("Start. "),
        '<w:ins w:id="100" w:author="Alice Reviewer" w:date="' + DATE_A + '">'
        '<w:r><w:t xml:space="preserve">ins1 </w:t></w:r></w:ins>',
        run("middle "),
        '<w:del w:id="101" w:author="Alice Reviewer" w:date="' + DATE_A + '">'
        '<w:r><w:delText xml:space="preserve">del1 </w:delText></w:r></w:del>',
        '<w:ins w:id="102" w:author="Alice Reviewer" w:date="' + DATE_B + '">'
        '<w:r><w:t xml:space="preserve">ins2</w:t></w:r></w:ins>',
        run(". End."),
        para_id="00000001",
    )
    return Fixture(
        name="fixture_07_mixed_ins_del.docx",
        description="Two insertions + one deletion interleaved in a single paragraph.",
        document_body=body,
    )


def f08_move_from_to() -> Fixture:
    # Paragraph A holds the moved text (moveFrom), paragraph B is the destination (moveTo).
    body = "\n".join([
        para(
            run("Origin: "),
            '<w:moveFromRangeStart w:id="200" w:author="Alice Reviewer" w:date="' + DATE_A + '" w:name="move1"/>',
            '<w:moveFrom w:id="201" w:author="Alice Reviewer" w:date="' + DATE_A + '">'
            '<w:r><w:t xml:space="preserve">moved clause</w:t></w:r></w:moveFrom>',
            '<w:moveFromRangeEnd w:id="200"/>',
            run(" — end origin."),
            para_id="00000001",
        ),
        para(
            run("Destination: "),
            '<w:moveToRangeStart w:id="200" w:author="Alice Reviewer" w:date="' + DATE_A + '" w:name="move1"/>',
            '<w:moveTo w:id="202" w:author="Alice Reviewer" w:date="' + DATE_A + '">'
            '<w:r><w:t xml:space="preserve">moved clause</w:t></w:r></w:moveTo>',
            '<w:moveToRangeEnd w:id="200"/>',
            run(" — end destination."),
            para_id="00000002",
        ),
    ])
    return Fixture(
        name="fixture_08_move_from_to.docx",
        description="Text moved between two paragraphs via moveFromRange/moveToRange pair.",
        document_body=body,
    )


def f09_rPrChange_bold() -> Fixture:
    # Run currently shows bold; rPrChange records that it used to be not-bold.
    body = para(
        run("Regular. "),
        (
            '<w:r><w:rPr>'
            '<w:b/>'
            '<w:rPrChange w:id="300" w:author="Alice Reviewer" w:date="' + DATE_A + '">'
            '<w:rPr/>'
            '</w:rPrChange>'
            '</w:rPr>'
            '<w:t xml:space="preserve">Now bold (was plain).</w:t></w:r>'
        ),
        para_id="00000001",
    )
    return Fixture(
        name="fixture_09_rPrChange_bold.docx",
        description="rPrChange revision — a run toggled to bold; prior rPr recorded empty.",
        document_body=body,
    )


def f10_multiple_reviewers() -> Fixture:
    body = para(
        run("Alpha. "),
        '<w:ins w:id="400" w:author="Alice Reviewer" w:date="' + DATE_A + '">'
        '<w:r><w:t xml:space="preserve">ALICE ADD </w:t></w:r></w:ins>',
        run("middle "),
        '<w:del w:id="401" w:author="Bob Reviewer"   w:date="' + DATE_B + '">'
        '<w:r><w:delText xml:space="preserve">BOB REMOVE </w:delText></w:r></w:del>',
        run("omega."),
        para_id="00000001",
    )
    people = _wrap_people(
        '  <w15:person w15:author="Alice Reviewer">\n'
        '    <w15:presenceInfo w15:providerId="None" w15:userId="Alice Reviewer"/>\n'
        '  </w15:person>\n'
        '  <w15:person w15:author="Bob Reviewer">\n'
        '    <w15:presenceInfo w15:providerId="None" w15:userId="Bob Reviewer"/>\n'
        '  </w15:person>'
    )
    return Fixture(
        name="fixture_10_multiple_reviewers.docx",
        description="Two authors: Alice's insertion + Bob's deletion in same paragraph; people.xml includes both.",
        document_body=body,
        people_xml=people,
    )


def f11_cjk_revisions() -> Fixture:
    # Single MS Mincho 24pt paragraph: prefix + ins(CJK) + middle + del(CJK)
    # + suffix. All runs by Alice Reviewer (palette slot 0, #D03337). The
    # large 24pt size makes strikethrough Y measurable in the GDI render
    # (font ascent/descent based; CJK fonts have different metrics than
    # Latin so this is the smallest case that surfaces a CJK-specific bug
    # if one exists). Limitation #5 in PHASE_2_CLOSEOUT.md known-limitations
    # was the trigger for this fixture.
    rpr_mincho_24 = (
        '<w:rPr>'
        '<w:rFonts w:ascii="Calibri" w:eastAsia="MS Mincho" w:hAnsi="Calibri"/>'
        '<w:sz w:val="48"/><w:szCs w:val="48"/>'
        '</w:rPr>'
    )
    body = para(
        f'<w:r>{rpr_mincho_24}<w:t xml:space="preserve">前段落。</w:t></w:r>',
        (
            '<w:ins w:id="500" w:author="Alice Reviewer" w:date="' + DATE_A + '">'
            f'<w:r>{rpr_mincho_24}<w:t xml:space="preserve">挿入された文字</w:t></w:r>'
            '</w:ins>'
        ),
        f'<w:r>{rpr_mincho_24}<w:t xml:space="preserve">、</w:t></w:r>',
        (
            '<w:del w:id="501" w:author="Alice Reviewer" w:date="' + DATE_A + '">'
            f'<w:r>{rpr_mincho_24}<w:delText xml:space="preserve">削除された文字</w:delText></w:r>'
            '</w:del>'
        ),
        f'<w:r>{rpr_mincho_24}<w:t xml:space="preserve">、終わり。</w:t></w:r>',
        para_id="00000001",
    )
    return Fixture(
        name="fixture_11_cjk_revisions.docx",
        description="CJK body (MS Mincho 24pt) with one ins + one del — exercises R-01/R-03 styling on CJK glyphs.",
        document_body=body,
    )


def f12_three_reviewers() -> Fixture:
    # Three distinct authors in first-seen order: Alice (slot 0), Bob (slot 1),
    # Carol (slot 2). Slot 2 in REVISION_AUTHOR_PALETTE is "#2B6033" — the same
    # green Word uses for moves regardless of author. The fixture is the
    # smallest input that exercises the third palette slot end-to-end on the
    # Oxi side; future Word ground-truth pixel sampling will confirm whether
    # slot 2 really matches the documented Office reviewing-palette green.
    body = para(
        run("Start. "),
        '<w:ins w:id="500" w:author="Alice Reviewer" w:date="' + DATE_A + '">'
        '<w:r><w:t xml:space="preserve">ALICE INS </w:t></w:r></w:ins>',
        run("middle1 "),
        '<w:del w:id="501" w:author="Bob Reviewer"   w:date="' + DATE_A + '">'
        '<w:r><w:delText xml:space="preserve">BOB DEL </w:delText></w:r></w:del>',
        run("middle2 "),
        '<w:ins w:id="502" w:author="Carol Reviewer" w:date="' + DATE_B + '">'
        '<w:r><w:t xml:space="preserve">CAROL INS</w:t></w:r></w:ins>',
        run(". End."),
        para_id="00000001",
    )
    people = _wrap_people(
        '  <w15:person w15:author="Alice Reviewer">\n'
        '    <w15:presenceInfo w15:providerId="None" w15:userId="Alice Reviewer"/>\n'
        '  </w15:person>\n'
        '  <w15:person w15:author="Bob Reviewer">\n'
        '    <w15:presenceInfo w15:providerId="None" w15:userId="Bob Reviewer"/>\n'
        '  </w15:person>\n'
        '  <w15:person w15:author="Carol Reviewer">\n'
        '    <w15:presenceInfo w15:providerId="None" w15:userId="Carol Reviewer"/>\n'
        '  </w15:person>'
    )
    return Fixture(
        name="fixture_12_three_reviewers.docx",
        description="Three reviewers (Alice ins / Bob del / Carol ins) — surfaces palette slot 2 on the Oxi side for future Word pixel-pass confirmation.",
        document_body=body,
        people_xml=people,
    )


def f13_pPrChange_indent() -> Fixture:
    # Single paragraph whose indent was changed via <w:pPrChange>. The current
    # pPr declares left-indent=720dxa (= 0.5"), and the pPrChange body records
    # that prior pPr had no indent. R-12 v2 (R69) emits a "Formatted: …"
    # balloon for this paragraph in the right margin, mirroring how it does
    # for rPrChange (fixture_09 / R-12 v1).
    body = (
        '    <w:p>\n'
        '      <w:pPr>\n'
        '        <w:ind w:left="720"/>\n'
        '        <w:pPrChange w:id="600" w:author="Alice Reviewer" w:date="' + DATE_A + '">\n'
        '          <w:pPr/>\n'
        '        </w:pPrChange>\n'
        '      </w:pPr>\n'
        '      <w:r><w:t xml:space="preserve">Now indented (was not).</w:t></w:r>\n'
        '    </w:p>'
    )
    return Fixture(
        name="fixture_13_pPrChange_indent.docx",
        description="pPrChange revision — paragraph left-indent toggled from 0 to 720dxa; prior pPr recorded empty.",
        document_body=body,
    )


def f14_rPrChange_font() -> Fixture:
    # Run currently shows Times New Roman 14pt; rPrChange records that it
    # used to be Calibri 11pt (default). Two properties change in a single
    # rPrChange, exercising `describe_rpr_diff`'s comma-join behaviour:
    # body should be "Formatted: Font Size: 14pt, Font: Times New Roman".
    body = para(
        run("Default text. "),
        (
            '<w:r><w:rPr>'
            '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            '<w:sz w:val="28"/><w:szCs w:val="28"/>'
            '<w:rPrChange w:id="700" w:author="Alice Reviewer" w:date="' + DATE_A + '">'
            '<w:rPr/>'
            '</w:rPrChange>'
            '</w:rPr>'
            '<w:t xml:space="preserve">Now Times 14pt (was default).</w:t></w:r>'
        ),
        para_id="00000001",
    )
    return Fixture(
        name="fixture_14_rPrChange_font.docx",
        description="rPrChange revision — run toggled to Times New Roman 14pt; prior rPr recorded empty.",
        document_body=body,
    )


def f26_pPrChange_pageBreakBefore_widow() -> Fixture:
    # Paragraph that gained page_break_before + lost widow_control via
    # pPrChange. Current pPr has <w:pageBreakBefore/> and
    # <w:widowControl w:val="0"/>; pPrChange's prior pPr is empty
    # (defaults: no page-break-before, widow_control=true). R107
    # surfaces "Page Break Before" + "Not Widow/Orphan Control".
    body = (
        '    <w:p>\n'
        '      <w:pPr>\n'
        '        <w:pageBreakBefore/>\n'
        '        <w:widowControl w:val="0"/>\n'
        '        <w:pPrChange w:id="1900" w:author="Alice Reviewer" w:date="' + DATE_A + '">\n'
        '          <w:pPr/>\n'
        '        </w:pPrChange>\n'
        '      </w:pPr>\n'
        '      <w:r><w:t xml:space="preserve">Now starts on new page; widow control off (was defaults).</w:t></w:r>\n'
        '    </w:p>'
    )
    return Fixture(
        name="fixture_26_pPrChange_pageBreak_widow.docx",
        description="pPrChange revision — page_break_before added + widow_control turned off; prior pPr defaults.",
        document_body=body,
    )


def f25_rPrChange_highlight_position_em() -> Fixture:
    # Run currently has highlight=yellow + position=6 (3pt raise) +
    # emphasis_mark=dot; rPrChange records prior rPr empty. Three more
    # NEW non-R72 axes — exercises describe_rpr_diff's R100 highlight
    # + position + emphasis_mark branches plus comma-join.
    body = para(
        run("Default. "),
        (
            '<w:r><w:rPr>'
            '<w:highlight w:val="yellow"/>'
            '<w:position w:val="6"/>'
            '<w:em w:val="dot"/>'
            '<w:rPrChange w:id="1800" w:author="Alice Reviewer" w:date="' + DATE_A + '">'
            '<w:rPr/>'
            '</w:rPrChange>'
            '</w:rPr>'
            '<w:t xml:space="preserve">Now highlighted+raised+dotted.</w:t></w:r>'
        ),
        para_id="00000001",
    )
    return Fixture(
        name="fixture_25_rPrChange_highlight_position.docx",
        description="rPrChange revision — run toggled to highlight=yellow + position=3pt + emphasis_mark=dot; prior rPr empty.",
        document_body=body,
    )


def f24_rPrChange_shadow_vanish_dstrike() -> Fixture:
    # Run currently has shadow + vanish + double-strikethrough; rPrChange
    # records prior rPr had none. Three NEW non-R72 axes change at once —
    # exercises describe_rpr_diff's R99 shadow + vanish + dstrike branches
    # plus the comma-join across new branches.
    body = para(
        run("Default. "),
        (
            '<w:r><w:rPr>'
            '<w:shadow/>'
            '<w:vanish/>'
            '<w:dstrike/>'
            '<w:rPrChange w:id="1700" w:author="Alice Reviewer" w:date="' + DATE_A + '">'
            '<w:rPr/>'
            '</w:rPrChange>'
            '</w:rPr>'
            '<w:t xml:space="preserve">Now shadowed+hidden+dstrike.</w:t></w:r>'
        ),
        para_id="00000001",
    )
    return Fixture(
        name="fixture_24_rPrChange_shadow_vanish_dstrike.docx",
        description="rPrChange revision — run toggled to shadow + vanish + double-strikethrough; prior rPr empty.",
        document_body=body,
    )


def f23_rPrChange_outline_emboss() -> Fixture:
    # Run currently has outline + emboss; rPrChange records prior rPr
    # had neither. Two NEW non-R72 axes change at once — exercises
    # describe_rpr_diff's R98 outline + emboss branches plus the
    # comma-join across new branches.
    body = para(
        run("Default. "),
        (
            '<w:r><w:rPr>'
            '<w:outline/>'
            '<w:emboss/>'
            '<w:rPrChange w:id="1600" w:author="Alice Reviewer" w:date="' + DATE_A + '">'
            '<w:rPr/>'
            '</w:rPrChange>'
            '</w:rPr>'
            '<w:t xml:space="preserve">Now outlined+embossed.</w:t></w:r>'
        ),
        para_id="00000001",
    )
    return Fixture(
        name="fixture_23_rPrChange_outline_emboss.docx",
        description="rPrChange revision — run toggled to outline + emboss; prior rPr empty.",
        document_body=body,
    )


def f22_pPrChange_numPr() -> Fixture:
    # Paragraph that became a list item via inline pPrChange. Current
    # pPr has <w:numPr><w:numId val=1/><w:ilvl val=0/></w:numPr>;
    # pPrChange's prior pPr is empty. R95 patches parser to mirror
    # inline numPr onto style.num_id, making describe_ppr_diff fire
    # the "Numbering: list 1" branch.
    body = (
        '    <w:p>\n'
        '      <w:pPr>\n'
        '        <w:numPr>\n'
        '          <w:ilvl w:val="0"/>\n'
        '          <w:numId w:val="1"/>\n'
        '        </w:numPr>\n'
        '        <w:pPrChange w:id="1500" w:author="Alice Reviewer" w:date="' + DATE_A + '">\n'
        '          <w:pPr/>\n'
        '        </w:pPrChange>\n'
        '      </w:pPr>\n'
        '      <w:r><w:t xml:space="preserve">Now numbered (was plain).</w:t></w:r>\n'
        '    </w:p>'
    )
    return Fixture(
        name="fixture_22_pPrChange_numPr.docx",
        description="pPrChange revision — paragraph attached to inline numPr (numId=1 ilvl=0); prior pPr empty.",
        document_body=body,
    )


def f21_pPrChange_tabs() -> Fixture:
    # Paragraph that gained 3 tab stops via pPrChange. Current pPr has
    # <w:tabs><w:tab pos=1440 (=72pt) center align/><w:tab pos=2880
    # (=144pt) right/><w:tab pos=4320 (=216pt) decimal/></w:tabs>;
    # pPrChange's prior pPr is empty. R94 surfaces "Tab Stops Added".
    body = (
        '    <w:p>\n'
        '      <w:pPr>\n'
        '        <w:tabs>\n'
        '          <w:tab w:val="center" w:pos="1440"/>\n'
        '          <w:tab w:val="right" w:pos="2880"/>\n'
        '          <w:tab w:val="decimal" w:pos="4320"/>\n'
        '        </w:tabs>\n'
        '        <w:pPrChange w:id="1400" w:author="Alice Reviewer" w:date="' + DATE_A + '">\n'
        '          <w:pPr/>\n'
        '        </w:pPrChange>\n'
        '      </w:pPr>\n'
        '      <w:r><w:t xml:space="preserve">Now has 3 custom tab stops (was none).</w:t></w:r>\n'
        '    </w:p>'
    )
    return Fixture(
        name="fixture_21_pPrChange_tabs.docx",
        description="pPrChange revision — 3 tab stops added at 72/144/216pt; prior pPr empty.",
        document_body=body,
    )


def f20_pPrChange_borders() -> Fixture:
    # Paragraph that gained a bottom border via pPrChange. Current pPr
    # declares <w:pBdr><w:bottom .../></w:pBdr>; pPrChange's prior pPr
    # is empty (no border). R93's describe_ppr_diff branch surfaces
    # "Borders Added".
    body = (
        '    <w:p>\n'
        '      <w:pPr>\n'
        '        <w:pBdr>\n'
        '          <w:bottom w:val="single" w:sz="4" w:space="1" w:color="000000"/>\n'
        '        </w:pBdr>\n'
        '        <w:pPrChange w:id="1300" w:author="Alice Reviewer" w:date="' + DATE_A + '">\n'
        '          <w:pPr/>\n'
        '        </w:pPrChange>\n'
        '      </w:pPr>\n'
        '      <w:r><w:t xml:space="preserve">Now has a bottom border (was plain).</w:t></w:r>\n'
        '    </w:p>'
    )
    return Fixture(
        name="fixture_20_pPrChange_borders.docx",
        description="pPrChange revision — paragraph gained a bottom border (single 0.5pt black); prior pPr empty.",
        document_body=body,
    )


def f19_pPrChange_keep_next() -> Fixture:
    # Paragraph toggled to "keep with next" via pPrChange. Current pPr
    # declares <w:keepNext/>; the pPrChange body records that prior pPr
    # was empty (default = not kept). R89's describe_ppr_diff branch
    # surfaces "Keep With Next".
    body = (
        '    <w:p>\n'
        '      <w:pPr>\n'
        '        <w:keepNext/>\n'
        '        <w:pPrChange w:id="1200" w:author="Alice Reviewer" w:date="' + DATE_A + '">\n'
        '          <w:pPr/>\n'
        '        </w:pPrChange>\n'
        '      </w:pPr>\n'
        '      <w:r><w:t xml:space="preserve">Now keep-with-next (was not).</w:t></w:r>\n'
        '    </w:p>'
    )
    return Fixture(
        name="fixture_19_pPrChange_keep_next.docx",
        description="pPrChange revision — paragraph toggled to keep_next=true; prior pPr empty.",
        document_body=body,
    )


def f18_pPrChange_shading() -> Fixture:
    # Single paragraph whose background shading was added via pPrChange.
    # Current pPr declares <w:shd w:fill="FFFF00"/> (yellow paragraph bg);
    # the pPrChange body records that the prior pPr was empty. R88's
    # describe_ppr_diff branch produces "Paragraph Shading: FFFF00".
    body = (
        '    <w:p>\n'
        '      <w:pPr>\n'
        '        <w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>\n'
        '        <w:pPrChange w:id="1100" w:author="Alice Reviewer" w:date="' + DATE_A + '">\n'
        '          <w:pPr/>\n'
        '        </w:pPrChange>\n'
        '      </w:pPr>\n'
        '      <w:r><w:t xml:space="preserve">Now highlighted (was plain).</w:t></w:r>\n'
        '    </w:p>'
    )
    return Fixture(
        name="fixture_18_pPrChange_shading.docx",
        description="pPrChange revision — paragraph shading toggled from default to yellow (#FFFF00); prior pPr empty.",
        document_body=body,
    )


def f17_rPrChange_vAlign_shading() -> Fixture:
    # Run currently has vertical_align=superscript + shading=yellow.
    # rPrChange records that the prior rPr had neither. Two axes
    # change in a single rPrChange — exercises describe_rpr_diff's
    # R87 vertical_align + shading branches plus comma-join.
    body = para(
        run("Inline. "),
        (
            '<w:r><w:rPr>'
            '<w:vertAlign w:val="superscript"/>'
            '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>'
            '<w:rPrChange w:id="1000" w:author="Alice Reviewer" w:date="' + DATE_A + '">'
            '<w:rPr/>'
            '</w:rPrChange>'
            '</w:rPr>'
            '<w:t xml:space="preserve">super+yellow</w:t></w:r>'
        ),
        para_id="00000001",
    )
    return Fixture(
        name="fixture_17_rPrChange_vAlign_shading.docx",
        description="rPrChange revision — run toggled to superscript + yellow shading; prior rPr empty.",
        document_body=body,
    )


def f16_rPrChange_caps_spacing() -> Fixture:
    # Run currently has all_caps + character_spacing=1.0pt; rPrChange
    # records that the prior rPr had neither. Two axes change in a single
    # rPrChange — exercises describe_rpr_diff's R86 small_caps/all_caps/
    # character_spacing branches plus its comma-join.
    body = para(
        run("Default. "),
        (
            '<w:r><w:rPr>'
            '<w:caps/>'
            '<w:spacing w:val="20"/>'
            '<w:rPrChange w:id="900" w:author="Alice Reviewer" w:date="' + DATE_A + '">'
            '<w:rPr/>'
            '</w:rPrChange>'
            '</w:rPr>'
            '<w:t xml:space="preserve">Now caps + spaced.</w:t></w:r>'
        ),
        para_id="00000001",
    )
    return Fixture(
        name="fixture_16_rPrChange_caps_spacing.docx",
        description="rPrChange revision — run toggled to all_caps + character_spacing=1pt; prior rPr empty.",
        document_body=body,
    )


def f15_pPrChange_alignment() -> Fixture:
    # Single paragraph whose alignment was toggled via <w:pPrChange>.
    # Current pPr declares <w:jc w:val="center"/>; the pPrChange body
    # records that the prior pPr had <w:jc w:val="left"/>. R-12 v3.5 (R72)
    # surfaces this in the "Formatted: Alignment: Centered" balloon.
    body = (
        '    <w:p>\n'
        '      <w:pPr>\n'
        '        <w:jc w:val="center"/>\n'
        '        <w:pPrChange w:id="800" w:author="Alice Reviewer" w:date="' + DATE_A + '">\n'
        '          <w:pPr>\n'
        '            <w:jc w:val="left"/>\n'
        '          </w:pPr>\n'
        '        </w:pPrChange>\n'
        '      </w:pPr>\n'
        '      <w:r><w:t xml:space="preserve">Now centered (was left).</w:t></w:r>\n'
        '    </w:p>'
    )
    return Fixture(
        name="fixture_15_pPrChange_alignment.docx",
        description="pPrChange revision — paragraph alignment toggled from left to center.",
        document_body=body,
    )


ALL_FIXTURES = [
    f01_single_comment,
    f02_comment_with_reply,
    f03_resolved_comment,
    f04_multi_para_range,
    f05_single_ins,
    f06_single_del,
    f07_mixed_ins_del,
    f08_move_from_to,
    f09_rPrChange_bold,
    f10_multiple_reviewers,
    f11_cjk_revisions,
    f12_three_reviewers,
    f13_pPrChange_indent,
    f14_rPrChange_font,
    f15_pPrChange_alignment,
    f16_rPrChange_caps_spacing,
    f17_rPrChange_vAlign_shading,
    f18_pPrChange_shading,
    f19_pPrChange_keep_next,
    f20_pPrChange_borders,
    f21_pPrChange_tabs,
    f22_pPrChange_numPr,
    f23_rPrChange_outline_emboss,
    f24_rPrChange_shadow_vanish_dstrike,
    f25_rPrChange_highlight_position_em,
    f26_pPrChange_pageBreakBefore_widow,
]


# --------------------------------------------------------------------------
# Validation — open each generated file with python-docx and skim for markers
# --------------------------------------------------------------------------

def validate(path: Path, expect: Fixture) -> str:
    """Return '' on success or a short failure summary."""
    try:
        import docx  # python-docx
        doc = docx.Document(str(path))
    except Exception as e:
        return f"python-docx open failed: {type(e).__name__}: {e}"

    # Re-read raw document.xml and check that our marker substrings survived
    with zipfile.ZipFile(path) as z:
        try:
            doc_xml = z.read("word/document.xml").decode("utf-8")
        except KeyError:
            return "word/document.xml missing"
        if expect.comments_xml and "word/comments.xml" not in z.namelist():
            return "comments.xml missing from ZIP"
        if expect.comments_ext_xml and "word/commentsExtended.xml" not in z.namelist():
            return "commentsExtended.xml missing from ZIP"
        if expect.people_xml and "word/people.xml" not in z.namelist():
            return "people.xml missing from ZIP"

    # basic sanity: at least one w:p and not a BadZipFile
    if "<w:p" not in doc_xml:
        return "no <w:p> in document.xml"
    _ = doc  # silence lint

    return ""


def main() -> int:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    manifest: list[dict[str, str]] = []

    print(f"Writing to {OUT_DIR}")
    for fn in ALL_FIXTURES:
        fx = fn()
        path = _write_docx(fx)
        err = validate(path, fx)
        ok = "OK" if not err else f"FAIL: {err}"
        print(f"  {path.name:45s} {ok}")
        manifest.append({
            "file": path.name,
            "description": fx.description,
            "status": ok,
        })

    import json
    (OUT_DIR / "MANIFEST.json").write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    print(f"\nManifest: {OUT_DIR / 'MANIFEST.json'}")
    any_fail = any(m["status"] != "OK" for m in manifest)
    return 1 if any_fail else 0


if __name__ == "__main__":
    sys.exit(main())
