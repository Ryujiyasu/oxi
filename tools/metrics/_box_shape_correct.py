# -*- coding: utf-8 -*-
import sys, zipfile, re
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

with zipfile.ZipFile('tools/golden-test/documents/docx/1ec1091177b1_006.docx') as z:
    doc = z.read('word/document.xml').decode('utf-8')
BOX = '□'
positions = [m.start() for m in re.finditer(BOX, doc)]

# Find all wp:anchor + ranges. Walk the AlternateContent tree.
# The anchor and wsp are siblings under graphic; the txbxContent is under wps:txbx under wps:wsp.
# Key insight: each AlternateContent block contains exactly one wp:anchor (positioning) + one wps:wsp (shape).

def find_enclosing_anchor(doc, pos):
    # Find the enclosing mc:AlternateContent
    ac_starts = [m.start() for m in re.finditer(r'<mc:AlternateContent>', doc)]
    ac_ends = [m.end() for m in re.finditer(r'</mc:AlternateContent>', doc)]
    for s in reversed(ac_starts):
        if s < pos:
            # Find matching close
            for e in ac_ends:
                if e > pos and e > s:
                    block = doc[s:e]
                    # Inspect anchor inside
                    docpr = re.search(r'<wp:docPr[^>]*?id="(\d+)"[^>]*?name="([^"]*)"', block)
                    extent = re.search(r'<wp:extent\s+cx="(\d+)"\s+cy="(\d+)"', block)
                    posH = re.search(r'<wp:positionH[^>]*relativeFrom="([^"]+)"[^>]*>(.*?)</wp:positionH>', block, re.DOTALL)
                    distL = re.search(r'<wp:anchor[^>]*?distL="(\d+)"', block)
                    distR = re.search(r'<wp:anchor[^>]*?distR="(\d+)"', block)
                    prst = re.search(r'<a:prstGeom\s+prst="([^"]+)"', block)
                    adj = re.search(r'<a:gd\s+name="adj"\s+fmla="val\s+(-?\d+)"', block)
                    lins = re.search(r'<wps:bodyPr[^>]*?lIns="(\d+)"', block)
                    bodypr = re.search(r'<wps:bodyPr[^>]*?>', block)
                    return {
                        'ac_start': s, 'ac_end': e,
                        'docpr_id': docpr.group(1) if docpr else None,
                        'docpr_name': docpr.group(2) if docpr else None,
                        'extent_cx_emu': int(extent.group(1)) if extent else None,
                        'extent_cy_emu': int(extent.group(2)) if extent else None,
                        'positionH_relativeFrom': posH.group(1) if posH else None,
                        'positionH_child': posH.group(2)[:60] if posH else None,
                        'distL_emu': int(distL.group(1)) if distL else None,
                        'distR_emu': int(distR.group(1)) if distR else None,
                        'prst': prst.group(1) if prst else None,
                        'adj': int(adj.group(1)) if adj else None,
                        'lIns_emu': int(lins.group(1)) if lins else None,
                        'bodypr_open': bodypr.group(0) if bodypr else None,
                    }
            break
    return None

for idx, pos in enumerate(positions):
    info = find_enclosing_anchor(doc, pos)
    if info:
        print(f"\nBOX[{idx+1}] pos={pos} → Shape id={info['docpr_id']} name={info['docpr_name']!r}")
        print(f"  positionH: relativeFrom={info['positionH_relativeFrom']} child={info['positionH_child']!r}")
        cx_pt = info['extent_cx_emu'] / 12700 if info['extent_cx_emu'] else None
        cy_pt = info['extent_cy_emu'] / 12700 if info['extent_cy_emu'] else None
        print(f"  extent: cx={info['extent_cx_emu']} EMU = {cx_pt:.2f}pt | cy={info['extent_cy_emu']} = {cy_pt:.2f}pt")
        distL_pt = info['distL_emu'] / 12700 if info['distL_emu'] else None
        distR_pt = info['distR_emu'] / 12700 if info['distR_emu'] else None
        print(f"  distL={info['distL_emu']}={distL_pt:.2f}pt distR={info['distR_emu']}={distR_pt:.2f}pt")
        lins_pt = info['lIns_emu'] / 12700 if info['lIns_emu'] is not None else None
        print(f"  prst={info['prst']} adj={info['adj']} lIns_emu={info['lIns_emu']} ({lins_pt}pt)")
        print(f"  bodypr_open: {info['bodypr_open']!r}")
    else:
        print(f"\nBOX[{idx+1}] pos={pos} → no enclosing anchor")
