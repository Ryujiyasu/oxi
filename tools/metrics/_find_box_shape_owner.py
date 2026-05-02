# -*- coding: utf-8 -*-
import sys, zipfile, re

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
BOX = '□'

with zipfile.ZipFile('tools/golden-test/documents/docx/1ec1091177b1_006.docx') as z:
    doc = z.read('word/document.xml').decode('utf-8')

# Find all wsp ranges with id and name
shape_ranges = []
for m in re.finditer(r'<wps:wsp', doc):
    s = m.start()
    e = doc.find('</wps:wsp>', s) + len('</wps:wsp>')
    block = doc[s:e]
    docpr = re.search(r'<wp:docPr[^>]*?id="(\d+)"[^>]*?name="([^"]*)"', block)
    if docpr:
        shape_ranges.append((s, e, docpr.group(1), docpr.group(2)))

# Find each □ position and identify its enclosing shape
positions = [m.start() for m in re.finditer(BOX, doc)]
for idx, pos in enumerate(positions):
    enclosing = None
    for s, e, sid, name in shape_ranges:
        if s <= pos < e:
            enclosing = (sid, name)
            break
    # Get context: the wsp's anchor positionH + extent + prst
    if enclosing:
        for s, e, sid, name in shape_ranges:
            if sid == enclosing[0]:
                blk = doc[s:e]
                ph = re.search(r'<wp:positionH[^>]*relativeFrom="([^"]+)"[^>]*>(.*?)</wp:positionH>', blk, re.DOTALL)
                ext = re.search(r'<wp:extent\s+cx="(\d+)"\s+cy="(\d+)"', blk)
                prst = re.search(r'<a:prstGeom\s+prst="([^"]+)"', blk)
                adj = re.search(r'<a:gd\s+name="adj"\s+fmla="val\s+(-?\d+)"', blk)
                lins = re.search(r'<wps:bodyPr[^>]*?lIns="(\d+)"', blk)
                # Find # of paragraphs in this shape's txbxContent
                txbx = re.search(r'<w:txbxContent>(.*?)</w:txbxContent>', blk, re.DOTALL)
                paras_in_box = len(re.findall(r'<w:p[ >]', txbx.group(1))) if txbx else 0
                print(f"\nBOX[{idx+1}] pos={pos} → Shape id={sid} name={name!r}")
                print(f"  positionH relativeFrom={ph.group(1) if ph else None} child={ph.group(2)[:80] if ph else None}")
                print(f"  extent cx={ext.group(1) if ext else None} cy={ext.group(2) if ext else None}")
                print(f"  prst={prst.group(1) if prst else None} adj={adj.group(1) if adj else None}")
                print(f"  lIns={lins.group(1) if lins else None}")
                print(f"  paragraphs in txbx: {paras_in_box}")
                # show first text in txbx
                if txbx:
                    first_t = re.search(r'<w:t[^>]*>([^<]*)</w:t>', txbx.group(1))
                    print(f"  first text: {first_t.group(1)[:30] if first_t else None!r}")
                break
    else:
        print(f"\nBOX[{idx+1}] pos={pos} → NO enclosing wsp (= body paragraph)")
