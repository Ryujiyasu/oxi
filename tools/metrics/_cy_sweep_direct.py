# -*- coding: utf-8 -*-
import time, pythoncom, win32com.client as wc
import sys, fitz, json
import os, zipfile, shutil, tempfile, re
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

ORIG = r'C:\Users\ryuji\oxi-4\tools\golden-test\documents\docx\1ec1091177b1_006.docx'
OUTDIR = r'C:\Users\ryuji\oxi-4\pipeline_data\1ec1_cy_sweep'
os.makedirs(OUTDIR, exist_ok=True)

CY_VALUES_EMU = [500000, 1000000, 1500000, 2000000, 2500000, 3028950, 3500000, 4000000, 5000000, 6057900, 8000000]


def make_modified(out_path, cy_emu):
    tmp = tempfile.mkdtemp(prefix='cy_')
    try:
        with zipfile.ZipFile(ORIG) as z:
            z.extractall(tmp)
        doc_path = os.path.join(tmp, 'word', 'document.xml')
        with open(doc_path, encoding='utf-8') as f:
            doc = f.read()
        BOX5 = doc.find('□', 80000)
        ac_starts = [m.start() for m in re.finditer(r'<mc:AlternateContent>', doc)]
        ac_ends = [m.end() for m in re.finditer(r'</mc:AlternateContent>', doc)]
        for s in reversed(ac_starts):
            if s < BOX5:
                for e in ac_ends:
                    if e > BOX5 and e > s:
                        ac_block = doc[s:e]
                        new_ac = ac_block.replace('cy="3028950"', f'cy="{cy_emu}"')
                        new_doc = doc[:s] + new_ac + doc[e:]
                        break
                break
        with open(doc_path, 'w', encoding='utf-8') as f:
            f.write(new_doc)
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
            for root, _, names in os.walk(tmp):
                for fn in names:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, tmp).replace(os.sep, '/')
                    z.write(full, arc)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


pythoncom.CoInitialize()
w = wc.Dispatch('Word.Application')
time.sleep(3)
try:
    w.Visible = False
    w.DisplayAlerts = False
except: pass

print(f'{"cy_pt":>10} {"min_dim":>10} {"corner_r":>10} {"formula":>10} {"meas_x":>10} {"meas_inset":>11} {"excess":>8}')
results = []

for cy_emu in CY_VALUES_EMU:
    cy_pt = cy_emu / 12700
    min_dim = min(522.75, cy_pt)
    corner_r = 0.04015 * min_dim
    formula = corner_r * 0.293
    docx = os.path.join(OUTDIR, f'cy_{cy_emu}.docx')
    pdf = os.path.join(OUTDIR, f'cy_{cy_emu}.pdf')
    for f in (docx, pdf):
        try: os.remove(f)
        except: pass
    make_modified(docx, cy_emu)
    try:
        d = w.Documents.Open(docx, ReadOnly=True)
        time.sleep(0.5)
        d.SaveAs2(pdf, FileFormat=17)
        d.Close(SaveChanges=False)
    except Exception as e:
        print(f'cy_{cy_emu}: {e}')
        continue

    pdoc = fitz.open(pdf)
    boxes = []
    for pi in range(pdoc.page_count):
        for inst in pdoc[pi].search_for('□'):
            boxes.append({'x': inst.x0, 'y': inst.y0, 'page': pi + 1})
    pdoc.close()

    # Shape 9 □: those NOT in Shape 35's known positions
    # Shape 35 □ at x≈46.08 or 46.56 with y in 100-400 (approx). Shape 9 □ comes after.
    # Actually use page logic: Shape 9 should follow Shape 35's last □
    cands = [b for b in boxes if b['y'] > 400 and b['page'] == 1]
    if not cands:
        cands = [b for b in boxes if b['page'] >= 2]
    if not cands and len(boxes) >= 5:
        cands = [boxes[4]]  # 5th □ in original = Shape 9 P1
    if cands:
        meas_x = cands[0]['x']
        meas_inset = meas_x - 44.36 + 1.08 - 0.5
        excess = meas_inset - formula
        print(f'{cy_pt:10.2f} {min_dim:10.2f} {corner_r:10.2f} {formula:10.2f} {meas_x:10.2f} {meas_inset:11.2f} {excess:8.2f}')
        results.append({'cy_pt': cy_pt, 'min_dim': min_dim, 'corner_r': corner_r,
                       'formula': formula, 'meas_x': meas_x, 'meas_inset': meas_inset,
                       'excess': excess, 'page': cands[0]['page'], 'all_boxes': boxes})
    else:
        print(f'cy_{cy_emu}: no Shape 9 box found ({len(boxes)} total)')

try: w.Quit()
except: pass

with open(os.path.join(OUTDIR, 'results.json'), 'w', encoding='utf-8') as f:
    json.dump(results, f, indent=2, ensure_ascii=False)
print('Saved')
