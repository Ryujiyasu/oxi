# -*- coding: utf-8 -*-
"""V_CC: Sweep Shape 9 cy to map the inset function.

Test cy values: 100000, 500000, 1000000, 2000000, 3028950 (orig), 4000000, 6057900 (V_BB4), 8000000, 12000000.

For each: predicted formula inset = adj_frac × min(cx,cy) × 0.293
        measured inset = advance_x - 44.36 (base_paragraph_x_at_no_inset_no_border) + 1.08 (effectExtent) - 0.5 (ln/2)
"""
import os, sys, time, json, zipfile, shutil, tempfile, re
import pythoncom, win32com.client as wc
import fitz

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ORIG_DOCX = os.path.abspath("tools/golden-test/documents/docx/1ec1091177b1_006.docx")
OUT_DIR = os.path.abspath("pipeline_data/1ec1_cy_sweep")
os.makedirs(OUT_DIR, exist_ok=True)


def make_modified(out_path, *, cy_emu):
    tmp = tempfile.mkdtemp(prefix='cy_')
    try:
        with zipfile.ZipFile(ORIG_DOCX) as z:
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
                        # Replace BOTH cy=3028950 inside this block (xfrm and extent)
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
                    arc = os.path.relpath(full, tmp).replace('\\', '/')
                    z.write(full, arc)
        return True
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def render_pdf(word, docx, pdf):
    last = None
    for attempt in range(5):
        try:
            d = word.Documents.Open(docx, ReadOnly=True)
            time.sleep(0.4)
            d.SaveAs2(pdf, FileFormat=17)
            d.Close(SaveChanges=False)
            return True
        except Exception as e:
            last = e
            time.sleep(1.0 + attempt * 0.5)
    print(f"  PDF ERR: {last}")
    return False


def measure_shape9_p1(pdf):
    """Find the first □ in Shape 9 (excluding Shape 35 □s at 46.08/46.56)."""
    d = fitz.open(pdf)
    boxes = []
    for pi in range(d.page_count):
        for inst in d[pi].search_for("□"):
            boxes.append({"page": pi+1, "x": inst.x0, "y": inst.y0})
    d.close()
    # Shape 35 □s are at fixed positions; Shape 9 □s come after
    # Shape 9 P1 is the first BOX with x > 47 OR after y > 400 (heuristic)
    candidates = [b for b in boxes if b['y'] > 400]
    return candidates[0] if candidates else (boxes[-1] if boxes else None), boxes


CY_VALUES_EMU = [
    500000,    # 39.37pt
    1000000,   # 78.74
    1500000,   # 118.11
    2000000,   # 157.48
    2500000,   # 196.85
    3028950,   # 238.50 (orig)
    3500000,   # 275.59
    4000000,   # 314.96
    5000000,   # 393.70
    6057900,   # 476.97 (V_BB4)
    8000000,   # 629.92
    12000000,  # 944.88
]


def main():
    pythoncom.CoInitialize()
    word = None
    for attempt in range(5):
        try:
            word = wc.Dispatch("Word.Application")
            time.sleep(2.0)
            word.Visible = False
            word.DisplayAlerts = False
            break
        except Exception as e:
            print(f"Word startup {attempt+1}: {e}")
            time.sleep(8.0)
    if word is None:
        print("Failed Word"); return
    print("Sweep Shape 9 cy to map inset function. cx=522.75pt fixed.")
    print(f"{'cy_pt':>10} {'min(cx,cy)':>12} {'corner_r':>10} {'formula_inset':>15} {'measured_x':>12} {'measured_inset':>15} {'excess':>8}")
    results = []
    try:
        for cy_emu in CY_VALUES_EMU:
            cy_pt = cy_emu / 12700
            min_dim = min(522.75, cy_pt)
            corner_r = 0.04015 * min_dim
            formula_inset = corner_r * 0.293
            vid = f"V_CC_cy_{cy_emu}"
            docx = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.docx"))
            pdf = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.pdf"))
            for f in (docx, pdf):
                try: os.remove(f)
                except: pass
            ok = make_modified(docx, cy_emu=cy_emu)
            if not ok:
                continue
            ok = render_pdf(word, docx, pdf)
            if not ok:
                results.append({"cy_emu": cy_emu, "cy_pt": cy_pt, "error": "render"})
                continue
            shape9_box, all_boxes = measure_shape9_p1(pdf)
            if shape9_box is None:
                continue
            measured_x = shape9_box['x']
            # measured_inset = measured_x - 44.36 (base) + 1.08 (effectExtent) - 0.5 (ln/2)
            measured_inset = measured_x - 44.36 + 1.08 - 0.5
            excess = measured_inset - formula_inset
            print(f"{cy_pt:10.2f} {min_dim:12.2f} {corner_r:10.2f} {formula_inset:15.2f} {measured_x:12.2f} {measured_inset:15.2f} {excess:8.2f}")
            results.append({"cy_emu": cy_emu, "cy_pt": cy_pt, "min_dim": min_dim,
                          "corner_r": corner_r, "formula_inset": formula_inset,
                          "measured_x": measured_x, "measured_inset": measured_inset,
                          "excess": excess})
    finally:
        try: word.Quit()
        except: pass
    out = os.path.join(OUT_DIR, "results.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved -> {out}")


if __name__ == "__main__":
    main()
