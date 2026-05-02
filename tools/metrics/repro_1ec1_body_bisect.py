# -*- coding: utf-8 -*-
"""V_U: Bisect 1ec1 body to find element that pushes Shape 9 □ from 46.56pt to 55.32pt.

Top-level body elements (from _dump_1ec1_body_structure.py):
  [1] Title paragraph
  [2] Intro paragraph
  [3] Shape 34 (◎ section header)
  [4] Shape 35 instance 1 (BOX[1]/[2])
  [5]-[11] Empty paragraphs
  [12] Shape 8 (◎ second section header)
  [13] Empty paragraph
  [14] Shape 9 ← TARGET (BOX[3]+, including BOX[5] = □３ at 55.32pt actual)
  [15+] After Shape 9

Bisection variants:
  V_U_full = no removal (control = 55.32pt)
  V_U_keep_14 = remove [1]..[13]
  V_U_keep_12_14 = remove [1]..[11], [13] (keep Shape 8 + Shape 9)
  V_U_keep_4_14 = remove [1..3], [5..13] (keep Shape 35 + Shape 9)
  V_U_keep_3_4_14 = keep Shape 34 + Shape 35 + Shape 9
  V_U_keep_1_14 = keep Title + Shape 9
  V_U_keep_4_only = keep Shape 35 only? would test if Shape 35 affects Shape 9 in same doc
"""
import os, sys, time, json, zipfile, shutil, tempfile, re
import pythoncom, win32com.client
import fitz

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ORIG_DOCX = os.path.abspath("tools/golden-test/documents/docx/1ec1091177b1_006.docx")
OUT_DIR = os.path.abspath("pipeline_data/1ec1_body_bisect")
os.makedirs(OUT_DIR, exist_ok=True)


def get_top_level_body_elements():
    """Re-extract top-level body element positions."""
    with zipfile.ZipFile(ORIG_DOCX) as z:
        doc = z.read('word/document.xml').decode('utf-8')
    body_m = re.search(r'<w:body>(.*?)</w:body>', doc, re.DOTALL)
    body = body_m.group(1)

    # Walk top-level (similar to _dump script)
    items = []
    pos = 0
    while pos < len(body):
        if body[pos] in '\n\r\t ':
            pos += 1
            continue
        m = re.match(r'<(w:p|w:tbl|w:sectPr|mc:AlternateContent)([> ])', body[pos:])
        if m:
            tag = m.group(1)
            close_tag = f'</{tag}>'
            depth = 1
            search_from = pos + len(m.group(0))
            while depth > 0 and search_from < len(body):
                next_open = re.search(rf'<{re.escape(tag)}[ >]', body[search_from:])
                next_close = body.find(close_tag, search_from)
                if next_close < 0: break
                if next_open and next_open.start() + search_from < next_close:
                    depth += 1
                    search_from = next_open.start() + search_from + len(next_open.group(0))
                else:
                    depth -= 1
                    search_from = next_close + len(close_tag)
            items.append((tag, pos, search_from))
            pos = search_from
        else:
            pos += 1
    return doc, body, items


def make_modified(out_path, *, keep_indices):
    """keep_indices: 1-based list of elements to KEEP. Shape 9 (index 14) and sectPr (30) always kept."""
    doc, body, items = get_top_level_body_elements()
    # Always keep [14] (Shape 9) and [30] (sectPr)
    keep = set(keep_indices) | {14, 30}

    # Build new body
    new_body_parts = []
    for i, (tag, s, e) in enumerate(items, 1):
        if i in keep:
            new_body_parts.append(body[s:e])
    new_body = '\n'.join(new_body_parts)

    new_doc = doc.replace(body, '\n' + new_body + '\n', 1)

    tmp = tempfile.mkdtemp(prefix='vu_')
    try:
        with zipfile.ZipFile(ORIG_DOCX) as z:
            z.extractall(tmp)
        with open(os.path.join(tmp, 'word', 'document.xml'), 'w', encoding='utf-8') as f:
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


def render_pdf(word, docx_path, pdf_path):
    last = None
    for attempt in range(5):
        try:
            doc = word.Documents.Open(docx_path, ReadOnly=True)
            time.sleep(0.4)
            doc.SaveAs2(pdf_path, FileFormat=17)
            doc.Close(SaveChanges=False)
            return True
        except Exception as e:
            last = e
            time.sleep(1.0 + attempt * 0.5)
    print(f"  PDF ERR: {last}")
    return False


def measure_box(pdf_path):
    d = fitz.open(pdf_path)
    results = []
    for pi in range(d.page_count):
        page = d[pi]
        for inst in page.search_for("□"):
            results.append({"page": pi+1, "x": inst.x0, "y": inst.y0})
    d.close()
    return results


# Bisection variants
VARIANTS = [
    ("V_U_full", list(range(1, 31))),  # all elements (control = 55.32)
    ("V_U_keep_14_only", []),  # only Shape 9 (= V_T0 essentially) → expect 46.56
    ("V_U_keep_1", [1]),  # title only + Shape 9
    ("V_U_keep_2", [2]),  # intro only
    ("V_U_keep_3", [3]),  # Shape 34
    ("V_U_keep_4", [4]),  # Shape 35
    ("V_U_keep_12", [12]),  # Shape 8
    ("V_U_keep_13", [13]),  # empty para before Shape 9
    ("V_U_keep_5to11", [5,6,7,8,9,10,11]),  # 7 empty paragraphs
    ("V_U_keep_1_2_3_4", [1,2,3,4]),  # all pre-Shape-35 + Shape 35
    ("V_U_keep_1_to_12", list(range(1, 13))),  # everything before empty[13]
    ("V_U_keep_1_to_13", list(range(1, 14))),  # everything before Shape 9
]


def main():
    pythoncom.CoInitialize()
    word = None
    for attempt in range(5):
        try:
            word = win32com.client.Dispatch("Word.Application")
            time.sleep(2.0)
            word.Visible = False
            word.DisplayAlerts = False
            break
        except Exception as e:
            print(f"Word startup {attempt+1}: {e}")
            time.sleep(8.0)
    if word is None:
        print("Failed Word"); return
    print("1ec1 body bisection — find element causing Shape 9 □#5 to render at 55.32pt")
    print("Target: BOX[5] = first □ in Shape 9 = source's □３ (left=105)")
    print("V_U_full = 55.32pt (control). V_U_keep_14_only = 46.56pt expected (V_T0 result)\n")
    results = []
    try:
        for vid, keep_idx in VARIANTS:
            print(f"=== {vid} keep={keep_idx} ===")
            docx = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.docx"))
            pdf = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.pdf"))
            for f in (docx, pdf):
                try: os.remove(f)
                except: pass
            ok = make_modified(docx, keep_indices=keep_idx)
            if not ok:
                results.append({"id": vid, "error": "modify failed"})
                continue
            ok = render_pdf(word, docx, pdf)
            if not ok:
                results.append({"id": vid, "error": "render failed"})
                continue
            boxes = measure_box(pdf)
            # Find Shape 9 P1 box (= BOX[5] in original = first □ paragraph in Shape 9)
            # In modified docs, Shape 9 □ might be at different y but we look for first one matching x≈55 or x≈46
            shape9_boxes = [b for b in boxes if 40 < b['x'] < 60]
            print(f"  {len(boxes)} □ instances; first 4 in Shape 9-position-range:")
            for b in shape9_boxes[:4]:
                print(f"    □ x={b['x']:.2f}pt y={b['y']:.2f}pt")
            # Identify which is BOX[5] (P1 of Shape 9)
            results.append({"id": vid, "boxes": boxes})
    finally:
        try: word.Quit()
        except: pass
    out = os.path.join(OUT_DIR, "results.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved -> {out}")


if __name__ == "__main__":
    main()
