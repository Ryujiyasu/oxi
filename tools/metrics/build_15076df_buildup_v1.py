"""Session 112 — build-up bisection from S111-3 minimal repro
(which renders '．' = 9.75pt) toward 15076df085f5 actual (which renders
'．' = 6.0pt at L12).

Each variant swaps ONE 15076df part into the minimal base. If any single
variant transitions '．' from 9.75pt -> 6.0pt, that file is the trigger.
If none do, conjunction is required and pair-variants must be tried next.

Variants:
  v1_base               : minimal (S111-3 compressPunc base, expected 9.75pt)
  v2_settings           : v1 with 15076df word/settings.xml swapped in
  v3_styles             : v1 with 15076df word/styles.xml swapped in
  v4_theme              : v1 + theme1.xml (referenced via styles theme refs)
  v5_fontTable          : v1 + fontTable.xml (font substitution metadata)
  v6_webSettings        : v1 + webSettings.xml
  v7_all                : v1 with ALL above swapped/added simultaneously

Run COM measurement after generation (next script).
"""
import os
import sys
import io
import shutil
import zipfile

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUT_DIR = os.path.normpath(os.path.join(REPO, "tools/metrics/15076df_buildup/variants"))
SRC_DIR = os.path.normpath(os.path.join(REPO, "tools/metrics/15076df_buildup/extracted"))
os.makedirs(OUT_DIR, exist_ok=True)


def read(rel):
    p = os.path.join(SRC_DIR, rel)
    with open(p, 'rb') as f:
        return f.read()


# ----- Minimal base parts (S111-3 compressPunc repro) -----
MIN_SETTINGS = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    b'<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
    b'<w:compat>\n'
    b'<w:balanceSingleByteDoubleByteWidth/>\n'
    b'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>\n'
    b'</w:compat>\n'
    b'<w:characterSpacingControl w:val="compressPunctuation"/>\n'
    b'</w:settings>\n'
)

MIN_STYLES = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    b'<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
    b'<w:docDefaults><w:rPrDefault><w:rPr>'
    b'<w:rFonts w:ascii="\xef\xbc\xad\xef\xbc\xb3 \xe6\x98\x8e\xe6\x9c\x9d" w:hAnsi="\xef\xbc\xad\xef\xbc\xb3 \xe6\x98\x8e\xe6\x9c\x9d" w:eastAsia="\xef\xbc\xad\xef\xbc\xb3 \xe6\x98\x8e\xe6\x9c\x9d" w:cs="\xef\xbc\xad\xef\xbc\xb3 \xe6\x98\x8e\xe6\x9c\x9d"/>'
    b'<w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/>'
    b'<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>'
    b'</w:rPr></w:rPrDefault></w:docDefaults>\n'
    b'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>\n'
    b'</w:styles>\n'
)

# document.xml shared by all variants — L12 echo
DOCUMENT = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    b' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">\n'
    b'<w:body>\n'
    b'<w:tbl><w:tblPr><w:tblW w:w="1968" w:type="dxa"/>'
    b'<w:tblLayout w:type="fixed"/>'
    b'<w:tblCellMar><w:left w:w="12" w:type="dxa"/><w:right w:w="12" w:type="dxa"/></w:tblCellMar>'
    b'</w:tblPr>'
    b'<w:tblGrid><w:gridCol w:w="1968"/></w:tblGrid>'
    b'<w:tr>'
    b'<w:tc><w:tcPr><w:tcW w:w="1968" w:type="dxa"/></w:tcPr>'
    b'<w:p><w:pPr>'
    b'<w:wordWrap w:val="0"/><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>'
    b'<w:adjustRightInd w:val="0"/>'
    b'<w:spacing w:line="240" w:lineRule="exact"/>'
    b'<w:ind w:left="215" w:right="76" w:hanging="192"/>'
    b'</w:pPr>'
    b'<w:r><w:rPr>'
    b'<w:rFonts w:hAnsi="\xef\xbc\xad\xef\xbc\xb3 \xe6\x98\x8e\xe6\x9c\x9d" w:cs="\xef\xbc\xad\xef\xbc\xb3 \xe6\x98\x8e\xe6\x9c\x9d" w:hint="eastAsia"/>'
    b'<w:spacing w:val="-9"/><w:kern w:val="0"/><w:szCs w:val="21"/>'
    b'</w:rPr>'
    b'<w:t>\xef\xbc\x91\xef\xbc\x8e\xe6\x8f\x90\xe4\xbe\x9b\xe3\x82\x92\xe5\x8f\x97\xe3\x81\x91\xe3\x81\x9f\xe5\x8c\xbf\xe5\x90\x8d\xe3\x83\x87\xe3\x83\xbc\xe3\x82\xbf\xe3\x81\xae\xe5\x90\x8d\xe7\xa7\xb0</w:t></w:r>'
    b'</w:p></w:tc></w:tr></w:tbl>\n'
    b'<w:sectPr>'
    b'<w:pgSz w:w="11906" w:h="16838"/>'
    b'<w:pgMar w:top="851" w:right="1134" w:bottom="567" w:left="1134"'
    b' w:header="851" w:footer="567" w:gutter="0"/>'
    b'<w:docGrid w:type="lines" w:linePitch="336"/>'
    b'</w:sectPr>\n'
    b'</w:body></w:document>\n'
)


def build_content_types(parts):
    """parts: dict of extra Override (PartName -> ContentType)."""
    s = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    s += b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
    s += b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'
    s += b'<Default Extension="xml" ContentType="application/xml"/>\n'
    s += b'<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n'
    s += b'<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>\n'
    s += b'<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>\n'
    for pn, ct in parts.items():
        s += f'<Override PartName="{pn}" ContentType="{ct}"/>\n'.encode('utf-8')
    s += b'</Types>\n'
    return s


def build_root_rels():
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>\n'
        b'</Relationships>\n'
    )


def build_doc_rels(extras=()):
    """extras: list of (Id, Type, Target)."""
    s = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n'
        b'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>\n'
    )
    for rid, typ, tgt in extras:
        s += f'<Relationship Id="{rid}" Type="{typ}" Target="{tgt}"/>\n'.encode('utf-8')
    s += b'</Relationships>\n'
    return s


def make_variant(name, settings_xml, styles_xml, extra_files=None, ct_extras=None, rel_extras=None):
    """extra_files: dict of zip-path -> bytes; ct_extras: dict; rel_extras: list."""
    extra_files = extra_files or {}
    ct_extras = ct_extras or {}
    rel_extras = rel_extras or []

    out_path = os.path.join(OUT_DIR, f"{name}.docx")
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', build_content_types(ct_extras))
        z.writestr('_rels/.rels', build_root_rels())
        z.writestr('word/_rels/document.xml.rels', build_doc_rels(rel_extras))
        z.writestr('word/document.xml', DOCUMENT)
        z.writestr('word/styles.xml', styles_xml)
        z.writestr('word/settings.xml', settings_xml)
        for p, b in extra_files.items():
            z.writestr(p, b)
    print(f"wrote {out_path}")
    return out_path


def main():
    # v1: minimal base — known to render '．' = 9.75pt (from S111-3)
    make_variant("v1_base", MIN_SETTINGS, MIN_STYLES)

    # v2: 15076df settings.xml swapped in (richer compat flags + character/spacing context)
    settings_15076 = read("word/settings.xml")
    make_variant("v2_settings", settings_15076, MIN_STYLES)

    # v3: 15076df styles.xml swapped in (full Word style table)
    styles_15076 = read("word/styles.xml")
    make_variant("v3_styles", MIN_SETTINGS, styles_15076)

    # v4: + theme1.xml (font theme — possibly drives MS Mincho fallback chain)
    theme_15076 = read("word/theme/theme1.xml")
    make_variant(
        "v4_theme", MIN_SETTINGS, MIN_STYLES,
        extra_files={"word/theme/theme1.xml": theme_15076},
        ct_extras={"/word/theme/theme1.xml": "application/vnd.openxmlformats-officedocument.theme+xml"},
        rel_extras=[("rId3",
                     "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                     "theme/theme1.xml")],
    )

    # v5: + fontTable.xml (font substitution metadata — PANOSE etc.)
    ft_15076 = read("word/fontTable.xml")
    make_variant(
        "v5_fontTable", MIN_SETTINGS, MIN_STYLES,
        extra_files={"word/fontTable.xml": ft_15076},
        ct_extras={"/word/fontTable.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"},
        rel_extras=[("rId3",
                     "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
                     "fontTable.xml")],
    )

    # v6: + webSettings.xml
    ws_15076 = read("word/webSettings.xml")
    make_variant(
        "v6_webSettings", MIN_SETTINGS, MIN_STYLES,
        extra_files={"word/webSettings.xml": ws_15076},
        ct_extras={"/word/webSettings.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"},
        rel_extras=[("rId3",
                     "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings",
                     "webSettings.xml")],
    )

    # v7: ALL swapped/added together (the "kitchen sink" — if THIS doesn't trigger
    # the bug, then document.xml structure itself matters and a follow-up variant
    # using actual L12 surroundings is needed)
    make_variant(
        "v7_all", settings_15076, styles_15076,
        extra_files={
            "word/theme/theme1.xml": theme_15076,
            "word/fontTable.xml": ft_15076,
            "word/webSettings.xml": ws_15076,
        },
        ct_extras={
            "/word/theme/theme1.xml": "application/vnd.openxmlformats-officedocument.theme+xml",
            "/word/fontTable.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
            "/word/webSettings.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
        },
        rel_extras=[
            ("rId3",
             "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
             "theme/theme1.xml"),
            ("rId4",
             "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
             "fontTable.xml"),
            ("rId5",
             "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings",
             "webSettings.xml"),
        ],
    )

    print(f"\nAll variants written to {OUT_DIR}")


if __name__ == "__main__":
    main()
