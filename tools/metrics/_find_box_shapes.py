# -*- coding: utf-8 -*-
import sys, zipfile, re

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

BOX = '□'  # □

with zipfile.ZipFile('tools/golden-test/documents/docx/1ec1091177b1_006.docx') as z:
    doc = z.read('word/document.xml').decode('utf-8')

print(f"Doc length: {len(doc)}")
print(f"BOX count: {doc.count(BOX)}")

shapes = []
for m in re.finditer(r'<wps:wsp.*?</wps:wsp>', doc, re.DOTALL):
    wsp = m.group(0)
    docpr = re.search(r'<wp:docPr[^>]*?id="(\d+)"[^>]*?name="([^"]*)"', wsp)
    if not docpr:
        continue
    sid, name = docpr.group(1), docpr.group(2)
    txbx = re.search(r'<w:txbxContent>(.*?)</w:txbxContent>', wsp, re.DOTALL)
    if not txbx:
        continue
    paras = re.findall(r'<w:p[ >].*?</w:p>', txbx.group(1), re.DOTALL)
    for i, p in enumerate(paras):
        ts = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', p))
        if BOX in ts:
            ind = re.search(r'<w:ind[^>]*?w:left="(\d+)"', p)
            ind_left = ind.group(1) if ind else 'None'
            ind_fl = re.search(r'<w:ind[^>]*?w:firstLine="(-?\d+)"', p)
            ind_hg = re.search(r'<w:ind[^>]*?w:hanging="(-?\d+)"', p)
            shapes.append({
                'sid': sid, 'name': name, 'pidx': i+1,
                'text': ts[:40],
                'ind_left': ind_left,
                'ind_fl': ind_fl.group(1) if ind_fl else None,
                'ind_hg': ind_hg.group(1) if ind_hg else None,
            })

print(f"\n{len(shapes)} BOX paragraphs in shapes:")
for s in shapes:
    print(f"  Shape id={s['sid']} name={s['name']!r} P{s['pidx']}: ind_left={s['ind_left']} fl={s['ind_fl']} hg={s['ind_hg']} | {s['text']!r}")

# Also check non-shape (body) paragraphs
print(f"\nBody paragraphs containing BOX:")
body_paras = re.findall(r'<w:p[ >](?:(?!<w:p[ >]).)*?</w:p>', doc, re.DOTALL)
b_count = 0
for p in body_paras:
    # skip if inside txbxContent
    if '<wps:wsp' in p or '<v:textbox' in p:
        continue
    ts = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', p))
    if BOX in ts:
        ind = re.search(r'<w:ind[^>]*?w:left="(\d+)"', p)
        ind_left = ind.group(1) if ind else 'None'
        b_count += 1
        if b_count <= 8:
            print(f"  body P: ind_left={ind_left} | {ts[:40]!r}")
print(f"Body BOX count: {b_count}")
