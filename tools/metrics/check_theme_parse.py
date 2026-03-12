"""Check what parse_theme returns for 1ec1's theme."""
import zipfile

docx_path = r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx\1ec1091177b1_006.docx"
with zipfile.ZipFile(docx_path) as z:
    with z.open('word/theme/theme1.xml') as f:
        content = f.read().decode('utf-8')

# Check the raw XML for font elements
import re
# Find majorFont section
major_match = re.search(r'<a:majorFont.*?</a:majorFont>', content, re.DOTALL)
if major_match:
    major_xml = major_match.group()
    print("=== majorFont XML ===")
    # Find ea element
    ea_match = re.search(r'<a:ea[^/]*/>', major_xml)
    print(f"ea element: {ea_match.group() if ea_match else 'NOT FOUND'}")
    # Find Jpan font
    jpan_match = re.search(r'<a:font[^>]*script="Jpan"[^/]*/>', major_xml)
    print(f"Jpan font: {jpan_match.group() if jpan_match else 'NOT FOUND'}")

minor_match = re.search(r'<a:minorFont.*?</a:minorFont>', content, re.DOTALL)
if minor_match:
    minor_xml = minor_match.group()
    print("\n=== minorFont XML ===")
    ea_match = re.search(r'<a:ea[^/]*/>', minor_xml)
    print(f"ea element: {ea_match.group() if ea_match else 'NOT FOUND'}")
    jpan_match = re.search(r'<a:font[^>]*script="Jpan"[^/]*/>', minor_xml)
    print(f"Jpan font: {jpan_match.group() if jpan_match else 'NOT FOUND'}")
