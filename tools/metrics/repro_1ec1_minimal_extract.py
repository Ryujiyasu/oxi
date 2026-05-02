# -*- coding: utf-8 -*-
"""V_T: Take 1ec1 actual docx, strip body to ONLY Shape 9 + a body paragraph.
Keep all OTHER package parts (theme, styles, fontTable, settings) intact.
Test if Shape 9 still renders at 55.32pt with minimal body content.
"""
import os, sys, time, json, zipfile, shutil, tempfile, re
import pythoncom, win32com.client
import fitz

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ORIG_DOCX = os.path.abspath("tools/golden-test/documents/docx/1ec1091177b1_006.docx")
OUT_DIR = os.path.abspath("pipeline_data/1ec1_minimal_extract")
os.makedirs(OUT_DIR, exist_ok=True)


def make_minimal(out_path, *, mode):
    tmp = tempfile.mkdtemp(prefix="extr_")
    try:
        with zipfile.ZipFile(ORIG_DOCX) as z:
            z.extractall(tmp)
        doc_path = os.path.join(tmp, 'word', 'document.xml')
        with open(doc_path, encoding='utf-8') as f:
            doc = f.read()

        # Find Shape 9 AC range
        BOX5 = doc.find('□', 80000)
        ac_starts = [m.start() for m in re.finditer(r'<mc:AlternateContent>', doc)]
        ac_ends = [m.end() for m in re.finditer(r'</mc:AlternateContent>', doc)]
        for s in reversed(ac_starts):
            if s < BOX5:
                for e in ac_ends:
                    if e > BOX5 and e > s:
                        ac_s, ac_e = s, e
                        break
                break
        shape9_block = doc[ac_s:ac_e]

        # Find the parent <w:p> wrapping Shape 9
        p_start = doc.rfind('<w:p ', 0, ac_s)
        if p_start < 0: p_start = doc.rfind('<w:p>', 0, ac_s)
        p_end = doc.find('</w:p>', ac_e) + len('</w:p>')
        shape9_para = doc[p_start:p_end]

        # Find sectPr
        sectpr_m = re.search(r'<w:sectPr[^>]*>.*?</w:sectPr>', doc, re.DOTALL)
        sectpr = sectpr_m.group(0) if sectpr_m else ''

        # Extract document root attributes (namespaces)
        root_m = re.search(r'<w:document\s+([^>]+)>', doc)
        root_attrs = root_m.group(1) if root_m else ''

        if mode == "shape9_only":
            new_body = f'<w:p><w:r><w:t>Body</w:t></w:r></w:p>{shape9_para}'
        elif mode == "shape9_no_anchor_para":
            new_body = shape9_para
        elif mode == "shape9_with_3_body":
            new_body = '<w:p><w:r><w:t>B1</w:t></w:r></w:p><w:p><w:r><w:t>B2</w:t></w:r></w:p><w:p><w:r><w:t>B3</w:t></w:r></w:p>' + shape9_para

        new_doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document {root_attrs}>\n<w:body>\n{new_body}\n{sectpr}\n</w:body>\n</w:document>'

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


VARIANTS = [
    ("V_T0_shape9_only", "shape9_only"),
    ("V_T1_shape9_no_anchor", "shape9_no_anchor_para"),
    ("V_T2_shape9_3_body_paras", "shape9_with_3_body"),
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
    print("Target: 1ec1 unmodified Shape 9 □ at x=55.32pt")
    print("If V_T variants give 55.32pt → residual 8.76pt is from non-document.xml package parts (theme/styles/fontTable)")
    print("If V_T gives ~46.56pt → residual is from removed paragraphs/shapes in document\n")
    results = []
    try:
        for vid, mode in VARIANTS:
            print(f"=== {vid} (mode={mode}) ===")
            docx = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.docx"))
            pdf = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.pdf"))
            for f in (docx, pdf):
                try: os.remove(f)
                except: pass
            ok = make_minimal(docx, mode=mode)
            if not ok:
                results.append({"id": vid, "error": "modify failed"})
                continue
            print(f"  Built ({os.path.getsize(docx)} bytes)")
            ok = render_pdf(word, docx, pdf)
            if not ok:
                results.append({"id": vid, "error": "render failed"})
                continue
            boxes = measure_box(pdf)
            for i, b in enumerate(boxes[:6]):
                print(f"    □#{i+1}: P{b['page']} x={b['x']:.2f}pt y={b['y']:.2f}pt")
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
