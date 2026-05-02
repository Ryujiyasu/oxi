# -*- coding: utf-8 -*-
"""V_BB: Test if Shape 9's X shift depends on its Y position on page.

Hypothesis: 8.76pt X shift correlates with Shape 9 vertical position.
- V_Y3 (all elements): Shape 9 at y=447 → x=55.32
- V_AA0 (only Shape 35): Shape 9 at y=101 → x=46.56

If hypothesis correct, manipulating preceding content size (without changing element count)
should change Shape 9's Y, AND its X."""
import os, sys, time, json, zipfile, shutil, tempfile, re
import pythoncom, win32com.client as wc
import fitz

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ORIG_DOCX = os.path.abspath("tools/golden-test/documents/docx/1ec1091177b1_006.docx")
OUT_DIR = os.path.abspath("pipeline_data/1ec1_y_position_test")
os.makedirs(OUT_DIR, exist_ok=True)


def make_modified(out_path, *, mode):
    tmp = tempfile.mkdtemp(prefix='ybb_')
    try:
        with zipfile.ZipFile(ORIG_DOCX) as z:
            z.extractall(tmp)
        doc_path = os.path.join(tmp, 'word', 'document.xml')
        with open(doc_path, encoding='utf-8') as f:
            doc = f.read()

        if mode == "shrink_shape35_height":
            # Find Shape 35 (id=35) extent and reduce cy by half
            # Shape 35 cy is 1657350 EMU (130.50pt) — try 100000 EMU (~7.87pt)
            # This makes Shape 35 super short, pushing Shape 9 closer to top
            # Actually simpler: shorten its bodyPr-displayed text via wrap
            # Let's try changing extent
            new_doc = doc.replace('cy="1657350"', 'cy="500000"', 2)  # both xfrm + extent
        elif mode == "remove_intro_text":
            # Find element [2] (intro paragraph with lots of text) and replace text with empty
            # Element 2 is the long intro. Replace its w:t content with just space
            # Find first <w:p containing "納税者の方が" text
            target_text = "納税者の方が期限内に納付されるよう"
            idx = doc.find(target_text)
            if idx > 0:
                p_start = doc.rfind('<w:p ', 0, idx)
                if p_start < 0: p_start = doc.rfind('<w:p>', 0, idx)
                p_end = doc.find('</w:p>', idx) + len('</w:p>')
                # Replace all w:t content in that paragraph with empty
                para = doc[p_start:p_end]
                para_new = re.sub(r'<w:t[^>]*>[^<]*</w:t>', '<w:t></w:t>', para)
                new_doc = doc[:p_start] + para_new + doc[p_end:]
            else:
                new_doc = doc
        elif mode == "shape9_posOffset_zero":
            # Set Shape 9 positionV posOffset to 0 (was 231140) — places shape at anchor-para top
            BOX5 = doc.find('□', 80000)
            ac_starts = [m.start() for m in re.finditer(r'<mc:AlternateContent>', doc)]
            ac_ends = [m.end() for m in re.finditer(r'</mc:AlternateContent>', doc)]
            for s in reversed(ac_starts):
                if s < BOX5:
                    for e in ac_ends:
                        if e > BOX5 and e > s:
                            ac_block = doc[s:e]
                            new_ac = re.sub(r'<wp:posOffset>231140</wp:posOffset>', '<wp:posOffset>0</wp:posOffset>', ac_block, count=1)
                            new_doc = doc[:s] + new_ac + doc[e:]
                            break
                    break
        elif mode == "shape9_extent_double":
            # Make Shape 9 extent double height (cy 3028950 → 6057900)
            BOX5 = doc.find('□', 80000)
            ac_starts = [m.start() for m in re.finditer(r'<mc:AlternateContent>', doc)]
            ac_ends = [m.end() for m in re.finditer(r'</mc:AlternateContent>', doc)]
            for s in reversed(ac_starts):
                if s < BOX5:
                    for e in ac_ends:
                        if e > BOX5 and e > s:
                            ac_block = doc[s:e]
                            new_ac = ac_block.replace('cy="3028950"', 'cy="6057900"')
                            new_doc = doc[:s] + new_ac + doc[e:]
                            break
                    break
        else:
            new_doc = doc

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


def measure(pdf):
    d = fitz.open(pdf)
    res = []
    for pi in range(d.page_count):
        for inst in d[pi].search_for("□"):
            res.append({"page": pi+1, "x": inst.x0, "y": inst.y0})
    d.close()
    return res


VARIANTS = [
    ("V_BB0_orig", "none"),
    ("V_BB1_shrink_shape35", "shrink_shape35_height"),
    ("V_BB2_empty_intro_text", "remove_intro_text"),
    ("V_BB3_shape9_posOffset_zero", "shape9_posOffset_zero"),
    ("V_BB4_shape9_extent_double", "shape9_extent_double"),
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
    print("Hypothesis: Shape 9 X shift correlates with Y position on page.")
    print("Original: x=55.32 at y=447. Need to test by manipulating Y.\n")
    results = []
    try:
        for vid, mode in VARIANTS:
            print(f"=== {vid} ({mode}) ===")
            docx = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.docx"))
            pdf = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.pdf"))
            for f in (docx, pdf):
                try: os.remove(f)
                except: pass
            ok = make_modified(docx, mode=mode)
            if not ok:
                results.append({"id": vid, "error": "modify"})
                continue
            ok = render_pdf(word, docx, pdf)
            if not ok:
                results.append({"id": vid, "error": "render"})
                continue
            boxes = measure(pdf)
            shape9 = [b for b in boxes if 40 < b['x'] < 70]
            print(f"  Total □: {len(boxes)}; in Shape range:")
            for b in shape9[:8]:
                print(f"    □ x={b['x']:.2f}pt y={b['y']:.2f}pt P{b['page']}")
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
