"""
Build numbered-list test fixtures via python-docx.

Tests:
  - Single-level numbered list with single-digit, multi-digit, and 3-digit
    numbers (1., 2., ..., 9., 10., 11., ..., 100., ..., 999.)
  - Multiple list-style number formats: decimal (1.), upperRoman (I.),
    lowerRoman (i.), upperLetter (A.), lowerLetter (a.), bullet (•)
  - Mixed: 3-level nesting with multi-digit at level 2

Goal:
  Measure the X position of:
   - The MARKER (number/bullet) glyph
   - The TEXT after the marker (for various number widths)

Output: output/numbered_list_fixtures/
"""
import os
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


OUT_DIR = os.path.join(os.path.dirname(__file__), "output", "numbered_list_fixtures")
os.makedirs(OUT_DIR, exist_ok=True)


NUMBERING_XML_TMPL = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="{fmt}"/>
      <w:lvlText w:val="{lvl_text}"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>"""


def build_fixture(path, *, fmt="decimal", lvl_text="%1.", n_items=15,
                  start_num=1):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Pt(72)
    sec.bottom_margin = Pt(72)
    sec.left_margin = Inches(1.0)
    sec.right_margin = Inches(1.0)

    # Add list paragraphs with numId reference
    for i in range(n_items):
        p = doc.add_paragraph(f"Item text content {start_num + i}")
        # Apply numbering via XML manipulation
        pPr = p._p.get_or_add_pPr()
        numPr = OxmlElement("w:numPr")
        ilvl = OxmlElement("w:ilvl")
        ilvl.set(qn("w:val"), "0")
        numPr.append(ilvl)
        numId = OxmlElement("w:numId")
        numId.set(qn("w:val"), "1")
        numPr.append(numId)
        pPr.append(numPr)

        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(11)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)

    doc.save(path)

    # Inject numbering.xml
    import zipfile, shutil
    tmp_path = path + ".tmp"
    num_xml = NUMBERING_XML_TMPL.format(fmt=fmt, lvl_text=lvl_text)
    with zipfile.ZipFile(path, "r") as zin:
        names = zin.namelist()
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                if name == "word/numbering.xml":
                    continue  # we'll write our own
                data = zin.read(name)
                if name == "[Content_Types].xml":
                    ct = data.decode("utf-8")
                    if "numbering.xml" not in ct:
                        ct = ct.replace(
                            "</Types>",
                            '<Override PartName="/word/numbering.xml" '
                            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'
                        )
                        data = ct.encode("utf-8")
                elif name == "word/_rels/document.xml.rels":
                    rels = data.decode("utf-8")
                    if "numbering.xml" not in rels:
                        rels = rels.replace(
                            "</Relationships>",
                            '<Relationship Id="rIdNum1" '
                            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" '
                            'Target="numbering.xml"/></Relationships>'
                        )
                        data = rels.encode("utf-8")
                zout.writestr(name, data)
            zout.writestr("word/numbering.xml", num_xml.encode("utf-8"))
    shutil.move(tmp_path, path)


def main():
    cases = [
        # (filename, kwargs)
        ("NL_decimal_1to15.docx",       dict(fmt="decimal",     lvl_text="%1.",  n_items=15, start_num=1)),
        ("NL_decimal_8to22.docx",       dict(fmt="decimal",     lvl_text="%1.",  n_items=15, start_num=8)),  # crosses 9→10
        ("NL_decimal_98to112.docx",     dict(fmt="decimal",     lvl_text="%1.",  n_items=15, start_num=98)), # crosses 99→100
        ("NL_upperRoman_1to10.docx",    dict(fmt="upperRoman",  lvl_text="%1.",  n_items=10, start_num=1)),
        ("NL_lowerRoman_1to10.docx",    dict(fmt="lowerRoman",  lvl_text="%1.",  n_items=10, start_num=1)),
        ("NL_upperLetter_1to10.docx",   dict(fmt="upperLetter", lvl_text="%1.",  n_items=10, start_num=1)),
        ("NL_lowerLetter_1to10.docx",   dict(fmt="lowerLetter", lvl_text="%1.",  n_items=10, start_num=1)),
        ("NL_bullet_5.docx",            dict(fmt="bullet",      lvl_text="•",    n_items=5,  start_num=1)),
        ("NL_paren_decimal.docx",       dict(fmt="decimal",     lvl_text="(%1)", n_items=12, start_num=1)),
        ("NL_decimal_dash.docx",        dict(fmt="decimal",     lvl_text="%1)",  n_items=12, start_num=1)),
    ]
    for name, kwargs in cases:
        path = os.path.join(OUT_DIR, name)
        try:
            build_fixture(path, **kwargs)
            print(f"  built {name}")
        except Exception as e:
            print(f"  ERR {name}: {e}")
    print(f"\nBuilt {len(cases)} fixtures in {OUT_DIR}")


if __name__ == "__main__":
    main()
