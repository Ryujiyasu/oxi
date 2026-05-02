"""Test the compatMode hypothesis for vertAnchor=text floating tables.

Hypothesis: compatMode=14 (Word 2010) ignores tblpY (slope=0).
            compatMode=15 (Word 2013+) respects tblpY (slope=1).

Build 4 minimal repros (each with 1 body para + floating tbl + tail):
  CT_v14_Y50    : compatMode=14, tblpY=50tw
  CT_v14_Y600   : compatMode=14, tblpY=600tw
  CT_v15_Y50    : compatMode=15, tblpY=50tw
  CT_v15_Y600   : compatMode=15, tblpY=600tw
  CT_none_Y50   : NO settings.xml
  CT_none_Y600  : NO settings.xml

Plus 2 inverse: TP3 stripped down to compat=15, FT_1para_Y600 with compat=14
in a separate output dir for direct flip-test.
"""
import zipfile
from pathlib import Path

OUT_DIR = Path(r"C:\Users\ryuji\oxi-1\tools\metrics\compat_test_repro")
OUT_DIR.mkdir(parents=True, exist_ok=True)


HEADER = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<w:body>'
)
FOOTER = (
    '<w:sectPr>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"'
    ' w:header="720" w:footer="720" w:gutter="0"/>'
    '<w:cols w:space="720"/>'
    '<w:docGrid w:linePitch="360"/>'
    '</w:sectPr>'
    '</w:body></w:document>'
)


def body_para(text: str) -> str:
    return (
        '<w:p><w:pPr><w:spacing w:before="0" w:after="0"'
        ' w:line="240" w:lineRule="auto"/></w:pPr>'
        '<w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"'
        ' w:eastAsia="ＭＳ 明朝" w:hint="eastAsia"/><w:sz w:val="21"/></w:rPr>'
        f'<w:t>{text}</w:t></w:r></w:p>'
    )


def floating_tbl(tblpY_tw: int, label: str) -> str:
    return (
        '<w:tbl><w:tblPr>'
        f'<w:tblpPr w:leftFromText="142" w:rightFromText="142"'
        f' w:vertAnchor="text" w:horzAnchor="margin" w:tblpY="{tblpY_tw}"/>'
        '<w:tblW w:type="dxa" w:w="9638"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="0000FF"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="0000FF"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="0000FF"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="0000FF"/>'
        '</w:tblBorders>'
        '</w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="9638"/></w:tblGrid>'
        '<w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="9638"/></w:tcPr>'
        f'{body_para(label)}'
        '</w:tc></w:tr></w:tbl>'
    )


def settings_xml(compat: int | None) -> str | None:
    """Return settings.xml content with given compatMode, or None to skip."""
    if compat is None:
        return None
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:compat>'
        f'<w:compatSetting w:name="compatibilityMode"'
        f' w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/>'
        '</w:compat>'
        '</w:settings>'
    )


def make_docx(name: str, body_xml: str, compat: int | None):
    out = OUT_DIR / f"{name}.docx"
    doc_xml = HEADER + body_xml + FOOTER

    if compat is not None:
        # Need settings.xml + ContentTypes override + relationship
        content_types = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml"'
            ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '<Override PartName="/word/settings.xml"'
            ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
            '</Types>'
        )
        doc_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"'
            ' Target="settings.xml"/>'
            '</Relationships>'
        )
    else:
        content_types = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml"'
            ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '</Types>'
        )
        doc_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
        )

    pkg_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
        ' Target="word/document.xml"/>'
        '</Relationships>'
    )

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", pkg_rels)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/document.xml", doc_xml)
        if compat is not None:
            z.writestr("word/settings.xml", settings_xml(compat))
    print(f"Wrote {out.name} (compat={compat})")


# Common body: 1 anchor para + floating tbl with tblpY + tail
def body(tblpY_tw: int, name: str) -> str:
    return body_para("anchor para A") + floating_tbl(tblpY_tw, f"{name} tblpY={tblpY_tw}tw") + body_para("tail paragraph")


def main():
    cases = [
        ("CT_v11_Y50",   50,   11),  # earlier compat (Word 2003)
        ("CT_v11_Y600", 600,   11),
        ("CT_v12_Y50",   50,   12),  # Word 2007
        ("CT_v12_Y600", 600,   12),
        ("CT_v14_Y50",   50,   14),  # Word 2010
        ("CT_v14_Y600", 600,   14),
        ("CT_v15_Y50",   50,   15),  # Word 2013+
        ("CT_v15_Y600", 600,   15),
        ("CT_none_Y50",  50, None),
        ("CT_none_Y600",600, None),
    ]
    for name, ytw, compat in cases:
        make_docx(name, body(ytw, name), compat)
    print(f"\nWrote {len(cases)} variants to {OUT_DIR}")


if __name__ == "__main__":
    main()
