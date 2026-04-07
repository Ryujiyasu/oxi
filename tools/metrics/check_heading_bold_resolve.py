"""Check if Word headings have bold via styles.xml + COM measurement.

For pipeline_data/docx/heading_numbering_01.docx:
1. Extract styles.xml, find Heading 1/2/3 styles
2. Check whether each has explicit <w:b/> in rPr
3. Use COM to verify each heading paragraph's actual rendered Bold attribute
"""
import zipfile
import xml.etree.ElementTree as ET
import win32com.client
from pathlib import Path

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

doc_path = Path("pipeline_data/docx/heading_numbering_01.docx").absolute()

# 1) Inspect styles.xml
with zipfile.ZipFile(doc_path) as z:
    styles_xml = z.read("word/styles.xml").decode("utf-8")

root = ET.fromstring(styles_xml)
print("=== Style definitions (heading-related) ===")
for style in root.findall("w:style", NS):
    style_id = style.get(f"{{{NS['w']}}}styleId", "")
    name_el = style.find("w:name", NS)
    name = name_el.get(f"{{{NS['w']}}}val", "") if name_el is not None else ""
    if "eading" not in style_id and "eading" not in name:
        continue
    rpr = style.find("w:rPr", NS)
    has_b = False
    has_b_val_false = False
    if rpr is not None:
        b = rpr.find("w:b", NS)
        if b is not None:
            has_b = True
            v = b.get(f"{{{NS['w']}}}val")
            if v == "0" or v == "false":
                has_b = False
                has_b_val_false = True
    based_on = style.find("w:basedOn", NS)
    base_id = based_on.get(f"{{{NS['w']}}}val") if based_on is not None else ""
    print(f"  styleId={style_id:<30} name={name:<25} basedOn={base_id:<20} bold={'YES' if has_b else 'no' + (' (val=0)' if has_b_val_false else '')}")

# 2) COM verify on actual paragraphs
print("\n=== COM measurement ===")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Open(str(doc_path), ReadOnly=True)
import time; time.sleep(1)

print(f"{'idx':<4} {'style':<25} {'bold':<6} text")
for i in range(1, min(doc.Paragraphs.Count + 1, 30)):
    p = doc.Paragraphs(i)
    style_name = p.Style.NameLocal
    bold = p.Range.Bold  # 0/-1/9999999
    text = p.Range.Text.strip()[:40]
    if "見出し" in style_name or "Heading" in style_name or "見出" in style_name:
        marker = "*"
    else:
        marker = " "
    print(f"{marker}{i:<3} {style_name:<25} {bold!s:<6} {text}")

doc.Close(SaveChanges=False)
word.Quit()
