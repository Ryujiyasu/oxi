"""Check rFonts in 1ec1 para 2 runs via XML."""
import zipfile
import xml.etree.ElementTree as ET

docx_path = r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx\1ec1091177b1_006.docx"
with zipfile.ZipFile(docx_path) as z:
    with z.open('word/document.xml') as f:
        tree = ET.parse(f)

ns = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}

# Get paragraph 2 (0-indexed = 1)
body = tree.getroot().find('.//w:body', ns)
paras = body.findall('w:p', ns)
print(f"Total paragraphs in XML: {len(paras)}")

# Para index 1 = second paragraph
para2 = paras[1]
runs = para2.findall('w:r', ns)
print(f"\nPara 2 has {len(runs)} runs:")

for i, run in enumerate(runs):
    rpr = run.find('w:rPr', ns)
    text_el = run.find('w:t', ns)
    text = text_el.text if text_el is not None else "(no text)"
    
    if rpr is not None:
        rfonts = rpr.find('w:rFonts', ns)
        if rfonts is not None:
            attrs = {k.split('}')[1] if '}' in k else k: v for k, v in rfonts.attrib.items()}
            print(f"  Run {i}: rFonts={attrs}")
        else:
            print(f"  Run {i}: no rFonts")
        
        sz = rpr.find('w:sz', ns)
        if sz is not None:
            print(f"    sz={sz.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')}")
    else:
        print(f"  Run {i}: no rPr")
    
    print(f"    text: \"{text[:50]}\"")

# Check theme fonts
with zipfile.ZipFile(docx_path) as z:
    with z.open('word/theme/theme1.xml') as f:
        theme = ET.parse(f)

ns2 = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
major = theme.getroot().find('.//a:majorFont', ns2)
minor = theme.getroot().find('.//a:minorFont', ns2)

if major is not None:
    latin = major.find('a:latin', ns2)
    ea = major.find('a:ea', ns2)
    print(f"\nmajorFont: latin={latin.get('typeface') if latin is not None else 'N/A'}, ea={ea.get('typeface') if ea is not None else 'N/A'}")
    # Check for specific language overrides
    for font in major.findall('a:font', ns2):
        script = font.get('script')
        typeface = font.get('typeface')
        if script in ['Jpan', 'Hans', 'Hant']:
            print(f"  major override: script={script} typeface={typeface}")

if minor is not None:
    latin = minor.find('a:latin', ns2)
    ea = minor.find('a:ea', ns2)
    print(f"minorFont: latin={latin.get('typeface') if latin is not None else 'N/A'}, ea={ea.get('typeface') if ea is not None else 'N/A'}")
    for font in minor.findall('a:font', ns2):
        script = font.get('script')
        typeface = font.get('typeface')
        if script in ['Jpan', 'Hans', 'Hant']:
            print(f"  minor override: script={script} typeface={typeface}")
