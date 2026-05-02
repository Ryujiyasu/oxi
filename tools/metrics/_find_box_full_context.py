# -*- coding: utf-8 -*-
import sys, zipfile, re

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

BOX = '□'

with zipfile.ZipFile('tools/golden-test/documents/docx/1ec1091177b1_006.docx') as z:
    doc = z.read('word/document.xml').decode('utf-8')
    styles = z.read('word/styles.xml').decode('utf-8')

# Find each □ paragraph's full context: pStyle, pPr, parent table cell ind, etc.
# Find first 8 BOX paragraphs and their tc context
positions = [m.start() for m in re.finditer(BOX, doc)]
print(f"BOX positions: {len(positions)}")

for idx, pos in enumerate(positions[:8]):
    # Walk back to enclosing <w:p
    p_start = doc.rfind('<w:p ', 0, pos)
    if p_start < 0:
        p_start = doc.rfind('<w:p>', 0, pos)
    p_end = doc.find('</w:p>', pos) + len('</w:p>')
    para = doc[p_start:p_end]
    # Walk back to enclosing <w:tc (table cell)
    tc_start = doc.rfind('<w:tc>', 0, p_start)
    in_tc = (tc_start >= 0 and doc.find('</w:tc>', tc_start) > p_end)
    # Walk back to enclosing <wps:wsp
    wsp_start = doc.rfind('<wps:wsp', 0, p_start)
    in_wsp = (wsp_start >= 0 and doc.find('</wps:wsp>', wsp_start) > p_end)
    # Walk back to enclosing <w:txbxContent
    in_txbx = '<w:txbxContent>' in doc[max(0, p_start-2000):p_start] and '</w:txbxContent>' in doc[p_end:p_end+2000]

    # pPr
    ppr_m = re.search(r'<w:pPr>(.*?)</w:pPr>', para, re.DOTALL)
    pStyle_v = re.search(r'<w:pStyle[^>]*?w:val="([^"]+)"', ppr_m.group(0)) if ppr_m else None
    ind_left = re.search(r'<w:ind[^>]*?w:left="(\d+)"', para)
    ind_fl = re.search(r'<w:ind[^>]*?w:firstLine="(-?\d+)"', para)
    ind_lc = re.search(r'<w:ind[^>]*?w:leftChars="(\d+)"', para)
    numpr = re.search(r'<w:numPr>', para)

    text = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', para))[:30]
    print(f"\n[{idx+1}] pos={pos} text={text!r}")
    print(f"  in_tc={in_tc} in_wsp={in_wsp} in_txbx={in_txbx}")
    print(f"  pStyle={pStyle_v.group(1) if pStyle_v else None}")
    print(f"  ind_left={ind_left.group(1) if ind_left else None} firstLine={ind_fl.group(1) if ind_fl else None} leftChars={ind_lc.group(1) if ind_lc else None}")
    print(f"  numPr={'YES' if numpr else 'no'}")
    if in_tc:
        # Get the tc's start, including its tcPr
        tc_text = doc[tc_start:p_start]
        tc_pr_m = re.search(r'<w:tcPr>(.*?)</w:tcPr>', tc_text, re.DOTALL)
        if tc_pr_m:
            tcPr = tc_pr_m.group(1)
            tcW = re.search(r'<w:tcW[^/>]*?w:w="(\d+)"', tcPr)
            print(f"  tcW={tcW.group(1) if tcW else None}")

    # Find pStyle in styles.xml for inherited indent
    if pStyle_v:
        sname = pStyle_v.group(1)
        # find style def
        sm = re.search(r'<w:style[^>]*w:styleId="' + re.escape(sname) + r'"[^>]*>.*?</w:style>', styles, re.DOTALL)
        if sm:
            sind = re.search(r'<w:ind[^>]*?w:left="(\d+)"', sm.group(0))
            sind_fl = re.search(r'<w:ind[^>]*?w:firstLine="(-?\d+)"', sm.group(0))
            sind_hg = re.search(r'<w:ind[^>]*?w:hanging="(-?\d+)"', sm.group(0))
            print(f"  STYLE {sname}: ind_left={sind.group(1) if sind else None} fl={sind_fl.group(1) if sind_fl else None} hg={sind_hg.group(1) if sind_hg else None}")
