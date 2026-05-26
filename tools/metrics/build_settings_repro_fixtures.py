"""Author minimal `word/settings.xml` repro fixtures for S315 —
settings parser coverage.

The settings parsers at parser/ooxml.rs:555-686 extract:
  - adjust_line_height_in_table  (substring search, line 560)
  - compat_mode                  (compatSetting w:name="compatibilityMode", line 564)
  - compress_punctuation         (characterSpacingControl, line 598)
  - compat bool flags            (any <w:flag/> inside <w:compat>, line 628)
  - default_tab_stop             (defaultTabStop, line 657)

These five parsers all share the same "read settings.xml → return
field" pattern but each has its own non-obvious branches:
  - compat_mode defaults to 15 (Word 2013+) on read-error AND on
    absent compatSetting (NOT to 14 — the parser explicitly favors
    modern Word).
  - compress_punctuation: TWO distinct val strings count as true
    ("compressPunctuation" AND "compressPunctuationAndJapaneseKana").
    "doNotCompress" is FALSE, same as absent — but the explicit
    "doNotCompress" pinning catches a future refactor that changes
    the default.
  - parse_compat_bool_flag tri-state: <flag/> (presence-no-val) → TRUE,
    <flag w:val="0"/> → FALSE, absent → FALSE. The middle case is
    DISTINCT from absent at the source level but indistinguishable
    at the IR level (both → false). Pinning val=0 SUPPRESSES catches
    a regression that flipped the polarity.
  - adjust_line_height_in_table uses SUBSTRING SEARCH (line 560) —
    NOT real XML parsing. <w:adjustLineHeightInTable w:val="0"/>
    would STILL return true because the substring is present. This
    is a deliberate shortcut; the test pins it so a future "proper
    parse" refactor is aware.

Outputs to ``tools/fixtures/settings_samples/``.
"""
import os
import zipfile

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "fixtures",
                       "settings_samples")

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

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="ＭＳ 明朝" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault/>
</w:docDefaults>
</w:styles>"""

DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t xml:space="preserve">body</w:t></w:r></w:p>
<w:sectPr>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="851" w:footer="992" w:gutter="0"/>
</w:sectPr>
</w:body>
</w:document>"""


def write_docx(path: str, settings_xml: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", DOCUMENT_XML)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/settings.xml", settings_xml)
    print(f"  wrote {path}")


def main() -> None:
    print(f"Writing fixtures to {OUT_DIR}/")

    # v1_settings_all_features_on: every settings parser hit with a
    # non-default value in one pass.
    #   - compatibilityMode=14 → compat_mode=14
    #   - characterSpacingControl="compressPunctuation" → compress_punctuation=true
    #   - <w:doNotExpandShiftReturn/> → do_not_expand_shift_return=true
    #   - <w:balanceSingleByteDoubleByteWidth/> → balance_single_byte_double_byte_width=true
    #   - <w:adjustLineHeightInTable/> → adjust_line_height_in_table=true
    #   - <w:defaultTabStop w:val="708"/> → default_tab_stop=Some(35.4)
    settings = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:defaultTabStop w:val="708"/>
<w:characterSpacingControl w:val="compressPunctuation"/>
<w:compat>
<w:doNotExpandShiftReturn/>
<w:balanceSingleByteDoubleByteWidth/>
<w:adjustLineHeightInTable/>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>
</w:compat>
</w:settings>"""
    write_docx(
        os.path.join(OUT_DIR, "v1_settings_all_features_on.docx"),
        settings,
    )

    # v1_settings_minimal_defaults: empty <w:settings> → all defaults.
    #   - compat_mode=15 (Word 2013+, NOT 14)
    #   - compress_punctuation=false
    #   - do_not_expand_shift_return=false
    #   - balance_single_byte_double_byte_width=false
    #   - adjust_line_height_in_table=false
    #   - default_tab_stop=None
    settings = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
</w:settings>"""
    write_docx(
        os.path.join(OUT_DIR, "v1_settings_minimal_defaults.docx"),
        settings,
    )

    # v1_settings_yakumono_kana_variant: characterSpacingControl=
    #   "compressPunctuationAndJapaneseKana" → compress_punctuation=true.
    # The parser at line 611-612 accepts TWO distinct val strings as
    # true. Pinning the kana variant (different from the main fixture's
    # "compressPunctuation") catches a regression that drops one.
    settings = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:characterSpacingControl w:val="compressPunctuationAndJapaneseKana"/>
</w:settings>"""
    write_docx(
        os.path.join(OUT_DIR, "v1_settings_yakumono_kana_variant.docx"),
        settings,
    )

    # v1_settings_yakumono_donotcompress: explicit "doNotCompress" →
    #   compress_punctuation=false. SAME end state as absent, but the
    #   explicit suppression catches a future regression that flips
    #   the default.
    settings = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:characterSpacingControl w:val="doNotCompress"/>
</w:settings>"""
    write_docx(
        os.path.join(OUT_DIR, "v1_settings_yakumono_donotcompress.docx"),
        settings,
    )

    # v1_settings_compat_flag_val_zero_suppresses: presence-with-no-val
    # → TRUE, val=0 → FALSE. Pins the polarity branch in
    # parse_compat_bool_flag (parser/ooxml.rs:640-645).
    #   <w:doNotExpandShiftReturn w:val="0"/>            → false
    #   <w:balanceSingleByteDoubleByteWidth w:val="false"/> → false
    # (Both "0" and "false" trigger the suppression branch.)
    settings = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat>
<w:doNotExpandShiftReturn w:val="0"/>
<w:balanceSingleByteDoubleByteWidth w:val="false"/>
</w:compat>
</w:settings>"""
    write_docx(
        os.path.join(
            OUT_DIR,
            "v1_settings_compat_flag_val_zero_suppresses.docx",
        ),
        settings,
    )

    print("Done.")


if __name__ == "__main__":
    main()
