"""Author minimal `<w:p><w:pPr>...</w:pPr></w:p>` repro fixtures for
S313 paragraph-properties coverage deepening.

`parse_paragraph_properties` (parser/ooxml.rs:1746) is the largest
single parser function in oxidocs and the foundation of every layout
decision. Several sub-features have dedicated suites already:
  - <w:pBdr>      → paragraph_borders_integration   (S302)
  - <w:tabs>      → tabstops_integration            (S286)
  - <w:numPr>     → numbering_integration           (S272)
  - <w:sectPr>    → section_integration             (S279)
  - <w:pPrChange> → property_change_integration     (S306)
But the BREADTH of inline pPr — jc, ind, spacing, the four boolean
toggles (wordWrap / autoSpaceDE / autoSpaceDN / snapToGrid), keeps,
widowControl — is not pinned at the integration layer.

This file covers:
  - jc aliases: "left|start" → Left, "right|end" → Right, "both" →
    Justify, "distribute" → Distribute. All four enum variants reached
    from a single fixture.
  - ind twip-priority (CLAUDE.md 2026-04-10): when BOTH `left` (twip)
    and `leftChars` are present, twip wins; `leftChars` is NOT stored
    on style.indent_left_chars. Single-fixture pin.
  - ind hanging → negative indent_first_line. Twip-priority is one
    case of polarity; sign-flip on hanging is another.
  - spacing modes: exact, auto, and the COM-confirmed Word quirk
    (S114 2026-05-15) where line="-240" with NO lineRule (or auto)
    is treated as wdLineSpaceExactly |val|/20 pt.
  - widowControl has_explicit tracking: a missing widowControl ≠
    `<w:widowControl/>` (presence) ≠ `<w:widowControl w:val="0"/>`
    (explicit off). The has_explicit_widow_control flag separates
    these three states so inheritance can decide.
  - Four boolean toggles all default to TRUE and flip to FALSE on
    `w:val="0"`. wordWrap=0 is the discriminator at the heart of
    S301's discriminator fix.

Outputs to ``tools/fixtures/paragraph_properties_samples/``.
"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures",
                       "paragraph_properties_samples")

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""

SETTINGS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat>
</w:settings>"""

# Intentionally minimal styles.xml: no pPrDefault. Without a pPrDefault
# the inline pPr is the SOLE source of paragraph formatting (the
# defaults applied are from ParagraphStyle::default() in Rust, NOT
# from any docx-side default). This isolates the test from style-
# inheritance behavior.
STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
</w:styles>"""

DOC_HEAD = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
"""

DOC_TAIL = (
    '<w:sectPr>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
    'w:header="851" w:footer="992" w:gutter="0"/>'
    '</w:sectPr>'
    '\n</w:body>\n</w:document>'
)


def _para(ppr_xml: str, text: str) -> str:
    ppr = f'<w:pPr>{ppr_xml}</w:pPr>' if ppr_xml else ''
    run = (
        f'<w:r><w:t xml:space="preserve">{escape(text)}'
        f'</w:t></w:r>'
    )
    return f'<w:p>{ppr}{run}</w:p>'


def write_docx(path: str, body_xml: str) -> None:
    full = DOC_HEAD + body_xml + DOC_TAIL
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", full)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/settings.xml", SETTINGS_XML)
    print(f"  wrote {path}")


def main() -> None:
    print(f"Writing fixtures to {OUT_DIR}/")

    # v1_pp_jc_aliases: 5 paragraphs covering each jc → Alignment
    # match arm at parser/ooxml.rs:1998-2004. "start"/"end" are the
    # bidi-friendly aliases for "left"/"right"; "both" means Justify
    # in OOXML (counter-intuitive name).
    body = (
        _para('<w:jc w:val="left"/>', "jc-left") +
        _para('<w:jc w:val="start"/>', "jc-start-alias") +
        _para('<w:jc w:val="right"/>', "jc-right") +
        _para('<w:jc w:val="end"/>', "jc-end-alias") +
        _para('<w:jc w:val="center"/>', "jc-center") +
        _para('<w:jc w:val="both"/>', "jc-both-equals-justify") +
        _para('<w:jc w:val="distribute"/>', "jc-distribute")
    )
    write_docx(os.path.join(OUT_DIR, "v1_pp_jc_aliases.docx"), body)

    # v1_pp_ind_twip_priority: pin the CLAUDE.md 2026-04-10 rule.
    # Paragraph 1: BOTH `left="720"` AND `leftChars="200"`. Parser
    # at ooxml.rs:2175-2180 detects has_twip_left=true and SUPPRESSES
    # storing leftChars (indent_left_chars stays None despite the
    # attribute being present). Twip is pre-computed by Word so it
    # wins over the character-based fallback.
    # Paragraph 2: ONLY `leftChars="200"` (no twip). Now leftChars
    # IS stored — there is no authoritative twip to defer to.
    # Paragraph 3: `hanging="120"` only — indent_first_line becomes
    # the NEGATIVE of val/20 (parser/ooxml.rs:2158).
    # Paragraph 4: `hangingChars="100"` only — indent_first_line_chars
    # is the NEGATIVE of hc (parser/ooxml.rs:2186).
    body = (
        _para(
            '<w:ind w:left="720" w:leftChars="200"/>',
            "twip-wins-over-chars",
        ) +
        _para(
            '<w:ind w:leftChars="200"/>',
            "chars-alone-stored",
        ) +
        _para(
            '<w:ind w:hanging="120"/>',
            "hanging-becomes-negative",
        ) +
        _para(
            '<w:ind w:hangingChars="100"/>',
            "hanging-chars-also-negative",
        )
    )
    write_docx(os.path.join(OUT_DIR, "v1_pp_ind_twip_priority.docx"), body)

    # v1_pp_spacing_line_modes: pin the three line-spacing modes plus
    # the negative-line Word quirk.
    # Paragraph 1: line=480 lineRule=auto → line_spacing=2.0 (auto =
    #   val/240, NOT val/20). 480/240 = 2.0 multiplier.
    # Paragraph 2: line=240 lineRule=exact → line_spacing=12pt
    #   (exact = val/20), line_spacing_rule="exact".
    # Paragraph 3: line=-240 (NO lineRule) → COM-confirmed 2026-05-15
    #   d4d126 Word quirk: treated as wdLineSpaceExactly with |val|/20
    #   = 12pt, line_spacing_rule materialized to "exact" (parser
    #   FABRICATES the rule because Word does too).
    # Paragraph 4: before=120 after=240 → 6pt / 12pt (twips/20).
    body = (
        _para(
            '<w:spacing w:line="480" w:lineRule="auto"/>',
            "line-auto-2x",
        ) +
        _para(
            '<w:spacing w:line="240" w:lineRule="exact"/>',
            "line-exact-12pt",
        ) +
        _para(
            '<w:spacing w:line="-240"/>',
            "line-negative-equals-exact",
        ) +
        _para(
            '<w:spacing w:before="120" w:after="240"/>',
            "before-after-twips",
        )
    )
    write_docx(os.path.join(OUT_DIR, "v1_pp_spacing_line_modes.docx"), body)

    # v1_pp_widow_control_explicit: the three states of widowControl
    # that has_explicit_widow_control tracks.
    # Paragraph 1: NO <w:widowControl> at all → widow_control=true
    #   (default), has_explicit_widow_control=false. Inheritance
    #   should still apply.
    # Paragraph 2: <w:widowControl/> with no val → widow_control=true,
    #   has_explicit_widow_control=true. Explicitly affirmed.
    # Paragraph 3: <w:widowControl w:val="0"/> → widow_control=false,
    #   has_explicit_widow_control=true. Explicitly OFF.
    # The middle state matters for OXI_FORCE_WIDOW (CLAUDE.md): an
    # env-flag that decides whether to FORCE widow_control on
    # paragraphs that didn't explicitly opt out.
    body = (
        _para("", "no-widow-attr") +
        _para('<w:widowControl/>', "explicit-on") +
        _para('<w:widowControl w:val="0"/>', "explicit-off")
    )
    write_docx(os.path.join(OUT_DIR, "v1_pp_widow_control_explicit.docx"), body)

    # v1_pp_word_wrap_autospace: pin the FOUR boolean toggles that all
    # default TRUE and flip to FALSE on val="0":
    #   - wordWrap   → word_wrap (S301 discriminator)
    #   - autoSpaceDE → auto_space_de (East-Asian / Latin spacing)
    #   - autoSpaceDN → auto_space_dn (East-Asian / numeral spacing)
    #   - snapToGrid → snap_to_grid (grid alignment)
    # Paragraph 1: all four flipped off in one pPr.
    # Paragraph 2: bare paragraph → all four stay at the Rust-side
    #   default (true).
    body = (
        _para(
            (
                '<w:wordWrap w:val="0"/>'
                '<w:autoSpaceDE w:val="0"/>'
                '<w:autoSpaceDN w:val="0"/>'
                '<w:snapToGrid w:val="0"/>'
            ),
            "four-off",
        ) +
        _para("", "four-default-on")
    )
    write_docx(os.path.join(OUT_DIR, "v1_pp_word_wrap_autospace.docx"), body)

    print("Done.")


if __name__ == "__main__":
    main()
