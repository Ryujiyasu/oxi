# -*- coding: utf-8 -*-
"""V_S: Modify 1ec1 to strip Shape 9 content down. Tests if content interacts with positioning."""
import os, sys, time, json, zipfile, shutil, tempfile, re
import pythoncom, win32com.client
import fitz

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ORIG_DOCX = os.path.abspath("tools/golden-test/documents/docx/1ec1091177b1_006.docx")
OUT_DIR = os.path.abspath("pipeline_data/1ec1_strip_shape9")
os.makedirs(OUT_DIR, exist_ok=True)


def make_modified(out_path, *, modification):
    tmp = tempfile.mkdtemp(prefix="strip_")
    try:
        with zipfile.ZipFile(ORIG_DOCX) as z:
            z.extractall(tmp)
        doc_path = os.path.join(tmp, 'word', 'document.xml')
        with open(doc_path, encoding='utf-8') as f:
            doc = f.read()

        BOX5 = doc.find('□', 80000)
        ac_starts = [m.start() for m in re.finditer(r'<mc:AlternateContent>', doc)]
        ac_ends = [m.end() for m in re.finditer(r'</mc:AlternateContent>', doc)]
        ac_s = ac_e = None
        for s in reversed(ac_starts):
            if s < BOX5:
                for e in ac_ends:
                    if e > BOX5 and e > s:
                        ac_s, ac_e = s, e
                        break
                break
        ac = doc[ac_s:ac_e]

        if modification == "strip_content":
            # Replace ALL paragraphs in txbxContent with single □３
            new_txbx = '''<w:txbxContent><w:p><w:pPr><w:snapToGrid w:val="0"/><w:spacing w:line="440" w:lineRule="exact"/><w:ind w:leftChars="50" w:left="105"/><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:rFonts w:asciiTheme="majorEastAsia" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorEastAsia" w:cs="FrankRuehl" w:hint="eastAsia"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t>□３</w:t></w:r></w:p></w:txbxContent>'''
            new_ac = re.sub(r'<w:txbxContent>.*?</w:txbxContent>', new_txbx, ac, count=1, flags=re.DOTALL)
        elif modification == "strip_solidFill":
            new_ac = re.sub(r'<a:solidFill><a:sysClr val="window"[^/]*/></a:solidFill>', '', ac, count=1)
        elif modification == "strip_line":
            new_ac = re.sub(r'<a:ln w="12700".*?</a:ln>', '', ac, count=1, flags=re.DOTALL)
        elif modification == "strip_effectLst":
            new_ac = ac.replace('<a:effectLst/>', '')
        elif modification == "strip_effectExtent":
            new_ac = re.sub(r'<wp:effectExtent[^/>]*?/>', '<wp:effectExtent l="0" t="0" r="0" b="0"/>', ac, count=1)
        elif modification == "remove_compatLnSpc":
            new_ac = ac.replace('compatLnSpc="1"', '')
        elif modification == "set_adj_to_zero":
            new_ac = re.sub(r'<a:gd name="adj" fmla="val 4015"/>', '<a:gd name="adj" fmla="val 0"/>', ac)
        elif modification == "set_adj_to_35000":
            new_ac = re.sub(r'<a:gd name="adj" fmla="val 4015"/>', '<a:gd name="adj" fmla="val 35000"/>', ac)
        elif modification == "set_prst_rect":
            new_ac = re.sub(r'<a:prstGeom prst="roundRect"><a:avLst>.*?</a:avLst></a:prstGeom>', '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>', ac, count=1)
        else:
            new_ac = ac

        new_doc = doc[:ac_s] + new_ac + doc[ac_e:]
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


def measure_box5(pdf_path):
    """Measure x position of the 5th □ instance (Shape 9 first BOX)."""
    d = fitz.open(pdf_path)
    page = d[0]
    insts = page.search_for("□")
    if len(insts) < 5:
        d.close()
        return [{"x": inst.x0} for inst in insts]
    d.close()
    return [{"x": insts[i].x0} for i in range(min(6, len(insts)))]


VARIANTS = [
    ("V_S0_orig", "none"),
    ("V_S1_strip_content", "strip_content"),
    ("V_S2_strip_solidFill", "strip_solidFill"),
    ("V_S3_strip_line", "strip_line"),
    ("V_S4_strip_effectLst", "strip_effectLst"),
    ("V_S5_strip_effectExtent", "strip_effectExtent"),
    ("V_S6_remove_compatLnSpc", "remove_compatLnSpc"),
    ("V_S7_set_adj_zero", "set_adj_to_zero"),
    ("V_S8_set_adj_35000", "set_adj_to_35000"),
    ("V_S9_set_prst_rect", "set_prst_rect"),
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
    print("1ec1 unmodified □#5 = 55.32pt; V_O3 synth = 46.56pt; gap = 8.76pt\n")
    print("If a modification drops □#5 to ~46.56pt, that property is the cause.\n")
    results = []
    try:
        for vid, mod in VARIANTS:
            print(f"=== {vid} (mod={mod}) ===")
            docx = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.docx"))
            pdf = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.pdf"))
            for f in (docx, pdf):
                try: os.remove(f)
                except: pass
            ok = make_modified(docx, modification=mod)
            if not ok:
                results.append({"id": vid, "error": "mod failed"})
                continue
            ok = render_pdf(word, docx, pdf)
            if not ok:
                results.append({"id": vid, "error": "render failed"})
                continue
            boxes = measure_box5(pdf)
            for i, b in enumerate(boxes):
                print(f"    □#{i+1}: x={b['x']:.2f}pt")
            box5_x = boxes[4]["x"] if len(boxes) >= 5 else None
            results.append({"id": vid, "box5_x": box5_x, "all_boxes": boxes})
            if box5_x is not None and abs(box5_x - 46.56) < 1.0:
                print(f"  >>> COLLAPSED to formula! ({box5_x:.2f}pt)")
    finally:
        try: word.Quit()
        except: pass
    out = os.path.join(OUT_DIR, "results.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved -> {out}")


if __name__ == "__main__":
    main()
