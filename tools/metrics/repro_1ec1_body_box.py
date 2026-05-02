# -*- coding: utf-8 -*-
"""V_J body □ position investigation — based on master's proven OOXML structure.

Reuses CTYPES, RELS_ROOT, WORD_RELS, SETTINGS from repro_1ec1_textbox_ind.
Uses STYLES with explicit MS Mincho minorEastAsia + MS Gothic majorEastAsia
to mirror 1ec1 theme effect WITHOUT a theme1.xml file.
"""
import os, sys, time, json, zipfile, shutil, tempfile
import pythoncom, win32com.client
import fitz

sys.path.insert(0, "tools/metrics")
from repro_1ec1_textbox_ind import CTYPES, RELS_ROOT, WORD_RELS

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath("pipeline_data/1ec1_body_box")
os.makedirs(OUT_DIR, exist_ok=True)

# Mirror 1ec1's docDefaults: kern=2, sz=21, eastAsia=ja-JP
STYLES_V_J = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century"/>
<w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/>
<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>
</w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>
</w:styles>'''

# Mirror 1ec1's settings
SETTINGS_V_J = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:defaultTabStop w:val="840"/>
<w:characterSpacingControl w:val="compressPunctuation"/>
<w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>
</w:compat>
</w:settings>'''


def doc_xml(*, ind_xml="", run_font_ascii="ＭＳ ゴシック", run_font_ea="ＭＳ ゴシック",
            run_kern_xml="", run_sz="28", text="□１"):
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
<w:pPr>
<w:snapToGrid w:val="0"/>
<w:spacing w:line="480" w:lineRule="exact"/>
<w:jc w:val="left"/>
{ind_xml}
</w:pPr>
<w:r>
<w:rPr>
<w:rFonts w:ascii="{run_font_ascii}" w:eastAsia="{run_font_ea}" w:hAnsi="{run_font_ascii}" w:hint="eastAsia"/>
<w:sz w:val="{run_sz}"/><w:szCs w:val="{run_sz}"/>
{run_kern_xml}
</w:rPr>
<w:t>{text}</w:t>
</w:r>
</w:p>
<w:sectPr>
<w:pgSz w:w="11906" w:h="16838" w:code="9"/>
<w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992" w:gutter="0"/>
<w:cols w:space="425"/>
<w:docGrid w:type="linesAndChars" w:linePitch="357"/>
</w:sectPr>
</w:body>
</w:document>'''


def write_docx(path, **kwargs):
    tmp = tempfile.mkdtemp(prefix="vj_")
    try:
        files = [
            ("[Content_Types].xml", CTYPES),
            ("_rels/.rels", RELS_ROOT),
            ("word/_rels/document.xml.rels", WORD_RELS),
            ("word/styles.xml", STYLES_V_J),
            ("word/settings.xml", SETTINGS_V_J),
            ("word/document.xml", doc_xml(**kwargs)),
        ]
        for relpath, content in files:
            full = os.path.join(tmp, relpath.replace("/", os.sep))
            os.makedirs(os.path.dirname(full), exist_ok=True)
            with open(full, "w", encoding="utf-8") as f:
                f.write(content)
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, names in os.walk(tmp):
                for fn in names:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, tmp).replace("\\", "/")
                    z.write(full, arc)
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
    page = d[0]
    instances = page.search_for("□")
    if not instances:
        d.close()
        return None
    inst = instances[0]
    zoom = 4.0
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    w, h, n = pix.width, pix.height, pix.n
    s = pix.samples
    top_px = int(inst.y0 * zoom)
    bottom_px = int(inst.y1 * zoom)
    left_search = max(0, int((inst.x0 - 2) * zoom))
    right_search = min(w, int((inst.x1 + 1) * zoom))
    leftmost = None
    for py in range(max(0, top_px), min(h, bottom_px)):
        for px in range(left_search, right_search):
            off = (py * w + px) * n
            r, g, bb = s[off], s[off+1], s[off+2]
            if r < 200 and g < 200 and bb < 200:
                if leftmost is None or px < leftmost:
                    leftmost = px
                break
    d.close()
    return {
        "search_x0": inst.x0,
        "search_x1": inst.x1,
        "leftmost_pt": leftmost / zoom if leftmost else None,
    }


VARIANTS = [
    ("V_J0_baseline_no_ind", {"ind_xml": ""}),
    ("V_J1_minorEA_MSMincho", {"ind_xml": "", "run_font_ascii": "Century", "run_font_ea": "ＭＳ 明朝"}),
    ("V_J2_left_105", {"ind_xml": '<w:ind w:left="105"/>'}),
    ("V_J3_leftChars_50_only", {"ind_xml": '<w:ind w:leftChars="50"/>'}),
    ("V_J4_left_105_AND_leftChars_50", {"ind_xml": '<w:ind w:left="105" w:leftChars="50"/>'}),
    ("V_J5_kern_off", {"ind_xml": "", "run_kern_xml": '<w:kern w:val="0"/>'}),
    ("V_J6_sz_21_default_run", {"ind_xml": "", "run_sz": "21"}),
    ("V_J7_no_run_font_uses_default", {"ind_xml": "", "run_font_ascii": "Century", "run_font_ea": "ＭＳ 明朝"}),
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
            print(f"Word startup attempt {attempt+1}: {e}")
            try:
                if word: word.Quit()
            except: pass
            word = None
            time.sleep(6.0)
    if word is None:
        print("Failed to start Word")
        return
    LEFT_MARGIN_PT = 851 / 20
    print(f"Page left margin: {LEFT_MARGIN_PT}pt")
    results = []
    try:
        for vid, kwargs in VARIANTS:
            print(f"\n=== {vid} ===")
            docx = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.docx"))
            pdf = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.pdf"))
            for f in (docx, pdf):
                try: os.remove(f)
                except: pass
            write_docx(docx, **kwargs)
            ok = render_pdf(word, docx, pdf)
            if not ok:
                results.append({"id": vid, **kwargs, "error": "render failed"})
                continue
            m = measure_box(pdf)
            if m:
                ex_adv = m["search_x0"] - LEFT_MARGIN_PT
                ex_vis = (m["leftmost_pt"] - LEFT_MARGIN_PT) if m["leftmost_pt"] else None
                vis_str = f"{ex_vis:.2f}" if ex_vis is not None else "NA"
                print(f"  search={m['search_x0']:.2f}pt visible={m['leftmost_pt']:.2f}pt | excess_adv={ex_adv:.2f}pt visible_excess={vis_str}pt")
                results.append({"id": vid, **kwargs, "measurement": m,
                                "excess_advance_pt": ex_adv,
                                "excess_visible_pt": ex_vis})
            else:
                results.append({"id": vid, **kwargs, "error": "no glyph"})
    finally:
        try: word.Quit()
        except: pass
    out = os.path.join(OUT_DIR, "results.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved -> {out}")


if __name__ == "__main__":
    main()
