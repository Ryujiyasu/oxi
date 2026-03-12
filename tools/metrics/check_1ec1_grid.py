"""Extract docGrid linePitch and paragraph spacing from 1ec1 OOXML."""
import zipfile
import xml.etree.ElementTree as ET
import os

docx_path = r"C:\Users\ryuji\oxi-1\tools\golden-test\documents\docx\1ec1091177b1_006.docx"

ns = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
}

with zipfile.ZipFile(docx_path) as z:
    # document.xml — docGrid and paragraph properties
    with z.open('word/document.xml') as f:
        tree = ET.parse(f)
        root = tree.getroot()

    # Find docGrid
    for grid in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}docGrid'):
        print(f"docGrid: {grid.attrib}")
        line_pitch = grid.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}linePitch')
        if line_pitch:
            print(f"  linePitch = {line_pitch} twips = {int(line_pitch)/20:.2f}pt")
        grid_type = grid.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type')
        print(f"  type = {grid_type}")

    # Find all paragraph properties with line spacing rules
    print("\nParagraphs with explicit lineSpacing:")
    body = root.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body')
    para_idx = 0
    for elem in body:
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
        if tag == 'p':
            para_idx += 1
            ppr = elem.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
            if ppr is not None:
                spacing = ppr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing')
                if spacing is not None:
                    attrs = {k.split('}')[-1] if '}' in k else k: v for k, v in spacing.attrib.items()}
                    if 'line' in attrs or 'before' in attrs or 'after' in attrs:
                        print(f"  Para {para_idx}: {attrs}")
                snap = ppr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}snapToGrid')
                if snap is not None:
                    val = snap.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'true')
                    print(f"  Para {para_idx}: snapToGrid={val}")

    # Check default paragraph style spacing
    print("\nStyles with lineSpacing:")
    with z.open('word/styles.xml') as f:
        styles_tree = ET.parse(f)
        styles_root = styles_tree.getroot()

    for style in styles_root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}style'):
        style_id = style.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId', '')
        name_el = style.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name')
        name = name_el.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '') if name_el is not None else ''
        ppr = style.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
        if ppr is not None:
            spacing = ppr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing')
            if spacing is not None:
                attrs = {k.split('}')[-1] if '}' in k else k: v for k, v in spacing.attrib.items()}
                if 'line' in attrs:
                    print(f"  {style_id} ({name}): {attrs}")

    # docDefaults
    print("\ndocDefaults rPr:")
    doc_defaults = styles_root.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}docDefaults')
    if doc_defaults is not None:
        rpr_default = doc_defaults.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
        if rpr_default is not None:
            for child in rpr_default:
                tag = child.tag.split('}')[-1]
                print(f"  {tag}: {child.attrib}")
        ppr_default = doc_defaults.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPrDefault')
        if ppr_default is not None:
            ppr = ppr_default.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
            if ppr is not None:
                for child in ppr:
                    tag = child.tag.split('}')[-1]
                    attrs = {k.split('}')[-1] if '}' in k else k: v for k, v in child.attrib.items()}
                    print(f"  pPrDefault/{tag}: {attrs}")
