# -*- coding: utf-8 -*-
"""V_Y: Test if my reconstruction process collapses the 55.32→46.56pt offset.

V_Y0: pure pass-through (read original, re-extract, re-zip with NO modifications)
V_Y1: extract document.xml, re-write same content, re-zip
V_Y2: my make_modified pass with all elements
V_Y3: byte-equivalent re-zip (no XML round-trip)
"""
import os, sys, time, json, zipfile, shutil, tempfile, re
import pythoncom, win32com.client as wc
import fitz

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ORIG_DOCX = os.path.abspath("tools/golden-test/documents/docx/1ec1091177b1_006.docx")
OUT_DIR = os.path.abspath("pipeline_data/1ec1_passthrough")
os.makedirs(OUT_DIR, exist_ok=True)


def make_y0_passthrough(out_path):
    """Just copy the original docx."""
    shutil.copy(ORIG_DOCX, out_path)


def make_y1_extract_rezip(out_path):
    """Extract, re-zip with no XML changes."""
    tmp = tempfile.mkdtemp(prefix='y1_')
    try:
        with zipfile.ZipFile(ORIG_DOCX) as z:
            z.extractall(tmp)
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
            for root, _, names in os.walk(tmp):
                for fn in names:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, tmp).replace('\\', '/')
                    z.write(full, arc)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def make_y2_xml_roundtrip(out_path):
    """Read XML, write back unchanged, re-zip."""
    tmp = tempfile.mkdtemp(prefix='y2_')
    try:
        with zipfile.ZipFile(ORIG_DOCX) as z:
            z.extractall(tmp)
        doc_path = os.path.join(tmp, 'word', 'document.xml')
        with open(doc_path, encoding='utf-8') as f:
            content = f.read()
        with open(doc_path, 'w', encoding='utf-8') as f:
            f.write(content)
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
            for root, _, names in os.walk(tmp):
                for fn in names:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, tmp).replace('\\', '/')
                    z.write(full, arc)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def make_y3_my_make_modified(out_path):
    """Use my make_modified with all elements kept (= reconstruction)."""
    tmp = tempfile.mkdtemp(prefix='y3_')
    try:
        with zipfile.ZipFile(ORIG_DOCX) as z:
            z.extractall(tmp)
        doc_path = os.path.join(tmp, 'word', 'document.xml')
        with open(doc_path, encoding='utf-8') as f:
            doc = f.read()
        body_m = re.search(r'<w:body>(.*?)</w:body>', doc, re.DOTALL)
        body = body_m.group(1)

        # Walk top-level
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
        # Keep ALL items
        new_body_parts = [body[s:e] for tag, s, e in items]
        new_body = '\n'.join(new_body_parts)
        new_doc = doc.replace(body, '\n' + new_body + '\n', 1)
        with open(doc_path, 'w', encoding='utf-8') as f:
            f.write(new_doc)
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
            for root, _, names in os.walk(tmp):
                for fn in names:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, tmp).replace('\\', '/')
                    z.write(full, arc)
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


def measure(pdf):
    d = fitz.open(pdf)
    res = []
    for pi in range(d.page_count):
        for inst in d[pi].search_for("□"):
            res.append({"page": pi+1, "x": inst.x0, "y": inst.y0})
    d.close()
    return res


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

    variants = [
        ("V_Y0_passthrough_copy", make_y0_passthrough),
        ("V_Y1_extract_rezip", make_y1_extract_rezip),
        ("V_Y2_xml_roundtrip", make_y2_xml_roundtrip),
        ("V_Y3_my_make_modified_full", make_y3_my_make_modified),
    ]
    results = []
    try:
        for vid, builder in variants:
            print(f"\n=== {vid} ===")
            docx = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.docx"))
            pdf = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.pdf"))
            for f in (docx, pdf):
                try: os.remove(f)
                except: pass
            builder(docx)
            ok = render_pdf(word, docx, pdf)
            if not ok:
                results.append({"id": vid, "error": "render"})
                continue
            boxes = measure(pdf)
            shape9 = [b for b in boxes if 40 < b['x'] < 70]
            print(f"  □ in Shape 9 region:")
            for b in shape9[:6]:
                print(f"    x={b['x']:.2f}pt y={b['y']:.2f}pt")
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
