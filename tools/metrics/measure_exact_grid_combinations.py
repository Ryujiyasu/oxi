"""Verify text_y_offset=0 for exact rule under various grid/header conditions.

Tests whether the rule "exact spacing → text at top of line box" holds when:
  - docGrid linesAndChars active (grid snap)
  - First line vs second line
  - With section header
  - With spaceBefore on the paragraph
"""
import win32com.client
import time
import sys
import json
import os

sys.stdout.reconfigure(encoding='utf-8')

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False

WD_EXACTLY = 4
WD_VERT_POS_REL_PAGE = 6


def make_doc_with_xml(xml_body):
    """Create a doc by injecting raw OOXML — for grid testing."""
    # Just use COM since simpler
    pass


def test_basic(font, fs, lh, top_pt=72.0, space_before=0):
    doc = word.Documents.Add()
    time.sleep(0.1)
    doc.PageSetup.TopMargin = top_pt
    doc.PageSetup.LeftMargin = 72
    doc.PageSetup.RightMargin = 72
    doc.PageSetup.BottomMargin = 72
    is_cjk = any(ord(c) > 0x2000 for c in font)
    text = "漢字テスト\n二行目漢字" if is_cjk else "ABCDE\nLine2"
    rng = doc.Range()
    rng.InsertAfter(text)
    rng = doc.Range()
    rng.Font.Name = font
    if is_cjk:
        rng.Font.NameFarEast = font
    rng.Font.Size = fs
    rng.ParagraphFormat.LineSpacingRule = WD_EXACTLY
    rng.ParagraphFormat.LineSpacing = lh
    rng.ParagraphFormat.SpaceBefore = space_before
    rng.ParagraphFormat.SpaceAfter = 0
    time.sleep(0.1)

    chars = doc.Range().Characters
    out = []
    for ci in range(1, min(15, chars.Count + 1)):
        try:
            c = chars(ci)
            ch = c.Text
            if ch in ('\r', '\x07', '\n'):
                continue
            cy = c.Information(WD_VERT_POS_REL_PAGE)
            out.append((ch, round(cy, 3)))
        except Exception:
            continue
    doc.Close(SaveChanges=False)
    return out


print("=== Test 1: Basic — verify offset=0 holds ===")
for font, fs, lh in [
    ("ＭＳ 明朝", 10.5, 17.0),
    ("ＭＳ 明朝", 10.5, 14.0),
    ("Calibri", 11.0, 17.0),
    ("Calibri", 11.0, 14.0),
]:
    s = test_basic(font, fs, lh)
    if s:
        l1y = s[0][1]
        l2y = next((y for c, y in s[1:] if y - l1y > 5), None)
        print(f"  {font:12s} {fs}pt exact {lh:5.1f}: l1y={l1y} l2y={l2y} actual_lh={l2y - l1y if l2y else None}")

print("\n=== Test 2: Different top margins (verify l1y == top_margin) ===")
for top in [36.0, 72.0, 100.0, 144.0]:
    s = test_basic("ＭＳ 明朝", 10.5, 17.0, top_pt=top)
    if s:
        l1y = s[0][1]
        delta = l1y - top
        print(f"  top={top:6.1f}pt → l1y={l1y:6.1f} delta={delta:+.2f}")

print("\n=== Test 3: spaceBefore (does it stack with line top?) ===")
for sb in [0, 6, 12, 24]:
    s = test_basic("ＭＳ 明朝", 10.5, 17.0, space_before=sb)
    if s:
        l1y = s[0][1]
        print(f"  spaceBefore={sb:3d}pt → l1y={l1y:6.1f} (expected 72+sb={72+sb})")

print("\n=== Test 4: Direct OOXML with charGrid linesAndChars ===")
# Create a minimal docx with charGrid
tmpdir = os.path.abspath("tools/metrics/output/_tmpdocx")
os.makedirs(tmpdir, exist_ok=True)
docx_path = os.path.join(tmpdir, "exact_grid.docx")

# Generate via simple OOXML
import zipfile
ooxml_doc = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:pPr><w:spacing w:line="340" w:lineRule="exact"/><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr></w:pPr>
<w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr><w:t>漢字行一</w:t></w:r>
</w:p>
<w:p><w:pPr><w:spacing w:line="340" w:lineRule="exact"/><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr></w:pPr>
<w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝"/><w:sz w:val="21"/></w:rPr><w:t>漢字行二</w:t></w:r>
</w:p>
<w:sectPr>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
<w:docGrid w:type="linesAndChars" w:linePitch="340" w:charSpace="0"/>
</w:sectPr>
</w:body>
</w:document>"""

content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", rels)
    z.writestr("word/document.xml", ooxml_doc)

doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
time.sleep(0.3)
chars = doc.Range().Characters
for ci in range(1, min(12, chars.Count + 1)):
    try:
        c = chars(ci)
        ch = c.Text
        if ch in ("\r", "\x07", "\n"):
            continue
        cy = c.Information(WD_VERT_POS_REL_PAGE)
        cx = c.Information(1)
        print(f"  ch={ch!r} x={cx:.2f} y={cy:.2f}")
    except Exception:
        pass
doc.Close(SaveChanges=False)

word.Quit()
