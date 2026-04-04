"""Inspect document XML for charSpace, charGrid, section settings."""
import zipfile
import xml.etree.ElementTree as ET
import sys

ns = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}

def inspect(path):
    with zipfile.ZipFile(path) as zf:
        # document.xml
        with zf.open('word/document.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()

        # Section properties
        for sp in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr'):
            print("=== sectPr ===")
            for child in sp:
                tag = child.tag.split('}')[-1]
                print(f"  {tag}: {child.attrib}")

        # docGrid
        for dg in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}docGrid'):
            print(f"\ndocGrid: {dg.attrib}")

        # Check for charSpace in document defaults
        try:
            with zf.open('word/styles.xml') as f:
                stree = ET.parse(f)
                sroot = stree.getroot()

            # docDefaults
            for dd in sroot.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}docDefaults'):
                print("\n=== docDefaults ===")
                for rpr in dd.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPrDefault'):
                    for rp in rpr.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr'):
                        for child in rp:
                            tag = child.tag.split('}')[-1]
                            print(f"  rPrDefault.{tag}: {child.attrib}")
                for ppr in dd.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPrDefault'):
                    for pp in ppr.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr'):
                        for child in pp:
                            tag = child.tag.split('}')[-1]
                            print(f"  pPrDefault.{tag}: {child.attrib}")

            # Normal style
            for style in sroot.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}style'):
                sid = style.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId', '')
                if sid == 'Normal' or sid == 'a':
                    print(f"\n=== Style: {sid} ===")
                    for rp in style.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr'):
                        for child in rp:
                            tag = child.tag.split('}')[-1]
                            print(f"  rPr.{tag}: {child.attrib}")
                    for pp in style.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr'):
                        for child in pp:
                            tag = child.tag.split('}')[-1]
                            print(f"  pPr.{tag}: {child.attrib}")
        except:
            pass

        # First table properties
        for tbl in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl'):
            print("\n=== First table ===")
            for tp in tbl.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblPr'):
                for child in tp:
                    tag = child.tag.split('}')[-1]
                    print(f"  tblPr.{tag}: {child.attrib}")

            # First row, first cell props
            for tr in tbl.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr'):
                for tc in tr.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc'):
                    tcp = tc.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcPr')
                    if tcp is not None:
                        print(f"  tcPr:")
                        for child in tcp:
                            tag = child.tag.split('}')[-1]
                            print(f"    {tag}: {child.attrib}")
                    break
                break
            # Grid columns
            tblGrid = tbl.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblGrid')
            if tblGrid is not None:
                cols = []
                for gc in tblGrid:
                    w = gc.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w', '?')
                    cols.append(w)
                print(f"  tblGrid cols (twips): {cols}")
                print(f"  tblGrid cols (pt): {[round(int(c)/20, 2) for c in cols if c != '?']}")
            break  # only first table


if __name__ == '__main__':
    import os
    path = os.path.abspath("tools/golden-test/documents/docx/459f05f1e877_kyodokenkyuyoushiki01.docx")
    inspect(path)
