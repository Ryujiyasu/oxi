# -*- coding: utf-8 -*-
import sys, zipfile, re
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
BOX = '□'

with zipfile.ZipFile('tools/golden-test/documents/docx/1ec1091177b1_006.docx') as z:
    doc = z.read('word/document.xml').decode('utf-8')

# Find each BOX position. Inspect everything BEFORE the BOX char in its paragraph.
positions = [m.start() for m in re.finditer(BOX, doc)]
for idx, pos in enumerate(positions[:6]):
    p_start = doc.rfind('<w:p ', 0, pos)
    if p_start < 0:
        p_start = doc.rfind('<w:p>', 0, pos)
    p_end = doc.find('</w:p>', pos) + len('</w:p>')
    para = doc[p_start:p_end]
    # Check pPr full
    ppr_m = re.search(r'<w:pPr>(.*?)</w:pPr>', para, re.DOTALL)
    pPr = ppr_m.group(0) if ppr_m else ''
    print(f"\n========== BOX[{idx+1}] pos={pos} ==========")
    print(f"pPr: {pPr}")
    # All runs before BOX
    pre_box = para[:pos - p_start]
    # Strip XML head to start at first <w:r
    first_r = pre_box.find('<w:r')
    if first_r > 0:
        pre_runs = pre_box[first_r:]
        # tokenize each <w:r .. </w:r> or <w:r ../>
        runs = re.findall(r'<w:r(?:\s[^>]*)?>.*?</w:r>|<w:r(?:\s[^>]*)?/>', pre_runs, re.DOTALL)
        print(f"Runs before BOX: {len(runs)}")
        for ri, r in enumerate(runs):
            t_m = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', r)
            tab_m = re.search(r'<w:tab/?>', r)
            br_m = re.search(r'<w:br/?>', r)
            sym_m = re.search(r'<w:sym[^/>]*/>', r)
            ftn_m = re.search(r'<w:footnoteReference[^/>]*/>', r)
            if t_m or tab_m or br_m or sym_m or ftn_m:
                print(f"  R{ri+1}: text={t_m} tab={'Y' if tab_m else ''} br={'Y' if br_m else ''} sym={'Y' if sym_m else ''} footnote={'Y' if ftn_m else ''}")
            else:
                print(f"  R{ri+1}: <empty> {r[:100]!r}")
    else:
        print("No runs before BOX")
