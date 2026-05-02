# -*- coding: utf-8 -*-
"""V_EE: What about Shape 35 triggers Shape 9 P1's +8.76pt offset?

Test variants:
- V_EE0: Shape 9 + a COPY of Shape 35 (control)
- V_EE1: Shape 9 + Shape 35 with adj=4015 (match Shape 9's adj)
- V_EE2: Shape 9 + Shape 35 with cy=Shape 9's cy
- V_EE3: Shape 9 + Shape 35 without BOX content (replace with normal text)
- V_EE4: Shape 9 + Shape 35 with prst=rect
- V_EE5: Shape 9 + minimal generic shape (different prst)
"""
import os, sys, time, json, zipfile, shutil, tempfile, re
import pythoncom, win32com.client as wc
import fitz

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ORIG_DOCX = os.path.abspath("tools/golden-test/documents/docx/1ec1091177b1_006.docx")
OUT_DIR = os.path.abspath("pipeline_data/1ec1_shape35_trigger")
os.makedirs(OUT_DIR, exist_ok=True)


def make_variant(out_path, *, mode):
    """Take 1ec1, keep only Shape 35 + Shape 9, then modify Shape 35 by mode."""
    tmp = tempfile.mkdtemp(prefix='ee_')
    try:
        with zipfile.ZipFile(ORIG_DOCX) as z:
            z.extractall(tmp)
        doc_path = os.path.join(tmp, 'word', 'document.xml')
        with open(doc_path, encoding='utf-8') as f:
            doc = f.read()

        # Walk top-level body
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

        # Keep elements 4 (Shape 35) + 14 (Shape 9) + 30 (sectPr)
        keep = {4, 14, 30}
        new_body_parts = []
        for i, (tag, s, e) in enumerate(items, 1):
            if i in keep:
                chunk = body[s:e]
                # Modify Shape 35 if it's element 4
                if i == 4:
                    if mode == "control":
                        pass
                    elif mode == "adj_4015":
                        chunk = chunk.replace('val 8396', 'val 4015')
                    elif mode == "cy_match_shape9":
                        chunk = chunk.replace('cy="1657350"', 'cy="3028950"')
                    elif mode == "remove_box_content":
                        # Replace □ with X in Shape 35's content
                        # Find Shape 35 content within element 4
                        pass  # complex, skip for now
                    elif mode == "prst_rect":
                        chunk = chunk.replace('prst="roundRect"', 'prst="rect"')
                    elif mode == "lIns_match_shape9":
                        chunk = chunk.replace('lIns="91440"', 'lIns="36000"').replace('rIns="91440"', 'rIns="36000"')
                    elif mode == "remove_solidFill":
                        chunk = re.sub(r'<a:solidFill><a:sysClr val="window" lastClr="FFFFFF"/></a:solidFill>', '', chunk, count=1)
                    elif mode == "remove_line":
                        chunk = re.sub(r'<a:ln w="12700"[^>]*>.*?</a:ln>', '', chunk, count=1, flags=re.DOTALL)
                    elif mode == "no_box_paras":
                        # Replace all □ with X within Shape 35's chunk
                        # Specifically inside w:txbxContent
                        chunk = chunk.replace('□', 'X')
                new_body_parts.append(chunk)
        new_body = '\n'.join(new_body_parts)
        new_doc = doc.replace(body, '\n' + new_body + '\n', 1)
        with open(doc_path, 'w', encoding='utf-8') as f:
            f.write(new_doc)
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
            for root, _, names in os.walk(tmp):
                for fn in names:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, tmp).replace(os.sep, '/')
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
    print(f"  ERR: {last}")
    return False


def measure(pdf):
    d = fitz.open(pdf)
    res = []
    for pi in range(d.page_count):
        for inst in d[pi].search_for("□"):
            res.append({"page": pi+1, "x": inst.x0, "y": inst.y0})
    d.close()
    return res


VARIANTS = [
    ("V_EE0_control_shape35_plus_9", "control"),
    ("V_EE1_shape35_adj_4015", "adj_4015"),
    ("V_EE2_shape35_cy_match_9", "cy_match_shape9"),
    ("V_EE3_shape35_prst_rect", "prst_rect"),
    ("V_EE4_shape35_lIns_match_9", "lIns_match_shape9"),
    ("V_EE5_shape35_no_solidFill", "remove_solidFill"),
    ("V_EE6_shape35_no_line", "remove_line"),
    ("V_EE7_shape35_no_box_replace_with_X", "no_box_paras"),
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
    print("V_AA0 control: Shape 9 P1 at x=55.32pt (trigger active)")
    print("V_U_keep_14_only: Shape 9 P1 at x=46.56pt (no trigger)\n")
    print("Testing what Shape 35 modification removes the trigger:\n")
    results = []
    try:
        for vid, mode in VARIANTS:
            print(f"=== {vid} (mode={mode}) ===")
            docx = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.docx"))
            pdf = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.pdf"))
            for f in (docx, pdf):
                try: os.remove(f)
                except: pass
            ok = make_variant(docx, mode=mode)
            if not ok:
                continue
            ok = render_pdf(word, docx, pdf)
            if not ok:
                results.append({"id": vid, "error": "render"})
                continue
            boxes = measure(pdf)
            # Shape 9 P1 = first BOX with x near 46.56 OR 55.32 (NOT 46.08 = Shape 35 left=0)
            # Sort by y, find first BOX whose x is NOT 46.08 (Shape 35 boxes typically at 46.08)
            # Shape 9 P1 will be visually distinct
            # Show all
            for b in sorted(boxes, key=lambda x: x['y']):
                print(f"   x={b['x']:.2f} y={b['y']:.2f}")
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
