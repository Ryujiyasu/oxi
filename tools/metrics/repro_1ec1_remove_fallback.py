# -*- coding: utf-8 -*-
"""V_R alternate: take 1ec1 actual docx, remove Shape 9's mc:Fallback,
re-zip, render. If Shape 9 □ position changes from 55.32pt to ~46.56pt,
then mc:Fallback IS the cause of the residual.
"""
import os, sys, time, json, zipfile, shutil, tempfile, re
import pythoncom, win32com.client
import fitz

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ORIG_DOCX = os.path.abspath("tools/golden-test/documents/docx/1ec1091177b1_006.docx")
OUT_DIR = os.path.abspath("pipeline_data/1ec1_modified_no_fallback")
os.makedirs(OUT_DIR, exist_ok=True)


def make_modified_docx(out_path, *, remove_shape9_fallback=True, remove_all_fallbacks=False):
    """Open 1ec1, modify document.xml to remove Shape 9 mc:Fallback, save as new docx."""
    tmp = tempfile.mkdtemp(prefix="modr_")
    try:
        # Extract original
        with zipfile.ZipFile(ORIG_DOCX) as z:
            z.extractall(tmp)
        doc_path = os.path.join(tmp, 'word', 'document.xml')
        with open(doc_path, encoding='utf-8') as f:
            doc = f.read()

        if remove_all_fallbacks:
            new_doc = re.sub(r'<mc:Fallback>.*?</mc:Fallback>', '', doc, flags=re.DOTALL)
            print(f"  Removed all mc:Fallback ({len(doc) - len(new_doc)} chars)")
        elif remove_shape9_fallback:
            # Find AC containing BOX[5] (pos 84340)
            BOX5 = doc.find('□', 80000)
            if BOX5 < 0:
                print("  BOX5 not found"); return False
            ac_starts = [m.start() for m in re.finditer(r'<mc:AlternateContent>', doc)]
            ac_ends = [m.end() for m in re.finditer(r'</mc:AlternateContent>', doc)]
            for s in reversed(ac_starts):
                if s < BOX5:
                    for e in ac_ends:
                        if e > BOX5 and e > s:
                            ac_block = doc[s:e]
                            new_ac = re.sub(r'<mc:Fallback>.*?</mc:Fallback>', '', ac_block, flags=re.DOTALL)
                            new_doc = doc[:s] + new_ac + doc[e:]
                            print(f"  Removed Shape 9 mc:Fallback ({len(ac_block) - len(new_ac)} chars)")
                            break
                    break
        else:
            new_doc = doc
        with open(doc_path, 'w', encoding='utf-8') as f:
            f.write(new_doc)

        # Re-zip
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


def measure_all_box(pdf_path):
    d = fitz.open(pdf_path)
    results = []
    for pi in range(d.page_count):
        page = d[pi]
        for inst in page.search_for("□"):
            results.append({"page": pi+1, "x": inst.x0, "y": inst.y0})
    d.close()
    return results


VARIANTS = [
    ("V_R_orig_unmodified", {"remove_shape9_fallback": False}),
    ("V_R_remove_shape9_fallback", {"remove_shape9_fallback": True}),
    ("V_R_remove_ALL_fallbacks", {"remove_all_fallbacks": True}),
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
    print("1ec1 actual unmodified target:")
    print("  □#5 (Shape 9 BOX[5]) at x=55.32pt")
    print("  □#1 (Shape 35 BOX[1]) at x=46.08pt\n")
    results = []
    try:
        for vid, kwargs in VARIANTS:
            print(f"=== {vid} ===")
            docx = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.docx"))
            pdf = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.pdf"))
            for f in (docx, pdf):
                try: os.remove(f)
                except: pass
            ok = make_modified_docx(docx, **kwargs)
            if not ok:
                results.append({"id": vid, "error": "modification failed"})
                continue
            print(f"  Built ({os.path.getsize(docx)} bytes)")
            ok = render_pdf(word, docx, pdf)
            if not ok:
                results.append({"id": vid, "error": "render failed"})
                continue
            boxes = measure_all_box(pdf)
            print(f"  □ instances on PDF:")
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
