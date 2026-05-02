# -*- coding: utf-8 -*-
"""V_V: Bisect 1ec1 body elements AFTER Shape 9 (15-30) to find +8.76pt trigger.

Top-level body elements after Shape 9:
  [15]-[26] = 12 empty paragraphs (193 bytes each)
  [27] = w:tbl (12485b, form table)
  [28] = empty paragraph (193b)
  [29] = w:p with multiple shapes (Rectangle 17, 19, 図3, Rectangle 23 — country tax agency stamp)
  [30] = w:sectPr
"""
import os, sys, time, json, zipfile, shutil, tempfile, re
import pythoncom, win32com.client
import fitz

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ORIG_DOCX = os.path.abspath("tools/golden-test/documents/docx/1ec1091177b1_006.docx")
OUT_DIR = os.path.abspath("pipeline_data/1ec1_body_bisect_after")
os.makedirs(OUT_DIR, exist_ok=True)


def get_top_level_body_elements():
    with zipfile.ZipFile(ORIG_DOCX) as z:
        doc = z.read('word/document.xml').decode('utf-8')
    body_m = re.search(r'<w:body>(.*?)</w:body>', doc, re.DOTALL)
    body = body_m.group(1)
    items = []
    pos = 0
    while pos < len(body):
        if body[pos] in '\n\r\t ':
            pos += 1; continue
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
    doc, body, items = get_top_level_body_elements()
    keep = set(keep_indices) | {14, 30}  # Shape 9 + sectPr always kept
    new_body_parts = []
    for i, (tag, s, e) in enumerate(items, 1):
        if i in keep:
            new_body_parts.append(body[s:e])
    new_body = '\n'.join(new_body_parts)
    new_doc = doc.replace(body, '\n' + new_body + '\n', 1)
    tmp = tempfile.mkdtemp(prefix='vv_')
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


# Bisect: progressively add elements 15..30 to see which triggers shift
VARIANTS = [
    # Test: which shape(s) trigger the +8.76pt
    ("V_AA0_only_shape35_4", [4]),  # Just Shape 35 + Shape 9
    ("V_AA1_only_shape34_3", [3]),  # Just Shape 34 + Shape 9
    ("V_AA2_only_shape8_12", [12]),  # Just Shape 8 + Shape 9
    ("V_AA3_only_shapes_3_4_12", [3, 4, 12]),  # All preceding shapes + Shape 9
    ("V_AA4_only_table_27", [27]),  # Just table + Shape 9
    ("V_AA5_only_final_shape_29", [29]),  # Just final shape + Shape 9
    ("V_AA6_only_shape35_4_AND_table_27", [4, 27]),
    ("V_AA7_only_shape35_4_AND_final29", [4, 29]),
    ("V_AA8_only_shape35_4_AND_27_29", [4, 27, 29]),
    ("V_AA9_only_shape34_3_AND_27_29", [3, 27, 29]),
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
    print("Bisecting POST-Shape-9 elements (15-30) to find +8.76pt trigger\n")
    print("Target: V_V8 (all post-elements) → 55.32pt; isolate which subset triggers\n")
    results = []
    try:
        for vid, keep_idx in VARIANTS:
            print(f"=== {vid} keep_after_idx={keep_idx} ===")
            docx = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.docx"))
            pdf = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.pdf"))
            for f in (docx, pdf):
                try: os.remove(f)
                except: pass
            ok = make_modified(docx, keep_indices=keep_idx)
            if not ok:
                results.append({"id": vid, "error": "modify"})
                continue
            ok = render_pdf(word, docx, pdf)
            if not ok:
                results.append({"id": vid, "error": "render"})
                continue
            boxes = measure_box(pdf)
            shape9_boxes = [b for b in boxes if 40 < b['x'] < 70]
            for b in shape9_boxes[:4]:
                print(f"    □ x={b['x']:.2f}pt y={b['y']:.2f}pt")
            results.append({"id": vid, "keep": keep_idx, "boxes": boxes})
    finally:
        try: word.Quit()
        except: pass
    out = os.path.join(OUT_DIR, "results.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved -> {out}")


if __name__ == "__main__":
    main()
