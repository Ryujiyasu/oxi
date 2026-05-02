# -*- coding: utf-8 -*-
"""V_K: Refine V_J — try to reproduce 1ec1's +5pt body □ offset.

Theories:
- T1: Run split (□ alone vs □１ together)
- T2: Default tab stop interaction
- T3: Page header presence shifts content
- T4: Earlier <w:sectPr> via continuous section break
- T5: Specific font with explicit name (no theme)
"""
import os, sys, time, json, zipfile, shutil, tempfile
import pythoncom, win32com.client
import fitz

sys.path.insert(0, "tools/metrics")
from repro_1ec1_textbox_ind import CTYPES, RELS_ROOT, WORD_RELS

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath("pipeline_data/1ec1_body_box_v2")
os.makedirs(OUT_DIR, exist_ok=True)

STYLES_V_K = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

SETTINGS_V_K = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:defaultTabStop w:val="840"/>
<w:characterSpacingControl w:val="compressPunctuation"/>
<w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>
</w:compat>
</w:settings>'''


def runs_v_k0_oneRun():
    """V_K0: □１ as one run (like V_J0)"""
    return '<w:r><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック" w:cs="FrankRuehl" w:hint="eastAsia"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t>□１</w:t></w:r>'

def runs_v_k1_splitRuns():
    """V_K1: □ then １ as separate runs (mirror 1ec1's actual structure)"""
    return ('<w:r w:rsidRPr="00C213C0"><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック" w:cs="FrankRuehl" w:hint="eastAsia"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t>□</w:t></w:r>'
            '<w:r w:rsidR="00304A2F"><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック" w:cs="FrankRuehl" w:hint="eastAsia"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t>１</w:t></w:r>')

def runs_v_k2_withRpr_in_pPr():
    """V_K2: pPr contains a w:rPr child (like 1ec1's pPr does)"""
    return runs_v_k0_oneRun()

def runs_v_k3_no_kern_in_run():
    """V_K3: leave out kern from run (rely on default)"""
    return '<w:r><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック" w:cs="FrankRuehl" w:hint="eastAsia"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t>□１</w:t></w:r>'


def doc_xml(*, ind_xml="", runs_xml, ppr_extra=""):
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
<w:pPr>
<w:snapToGrid w:val="0"/>
<w:spacing w:line="480" w:lineRule="exact"/>
<w:jc w:val="left"/>
{ind_xml}
{ppr_extra}
</w:pPr>
{runs_xml}
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
    tmp = tempfile.mkdtemp(prefix="vk_")
    try:
        files = [
            ("[Content_Types].xml", CTYPES),
            ("_rels/.rels", RELS_ROOT),
            ("word/_rels/document.xml.rels", WORD_RELS),
            ("word/styles.xml", STYLES_V_K),
            ("word/settings.xml", SETTINGS_V_K),
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
    return {"search_x0": inst.x0, "search_x1": inst.x1,
            "leftmost_pt": leftmost / zoom if leftmost else None}


# Each run kit + ppr_extra to test
PPR_RPR_1EC1 = '<w:rPr><w:rFonts w:asciiTheme="majorEastAsia" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorEastAsia" w:cs="FrankRuehl"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'

VARIANTS = [
    ("V_K0_oneRun_baseline", {"ind_xml": "", "runs_xml": runs_v_k0_oneRun()}),
    ("V_K1_splitRuns_baseline", {"ind_xml": "", "runs_xml": runs_v_k1_splitRuns()}),
    ("V_K2_pPr_with_rPr_child", {"ind_xml": "", "runs_xml": runs_v_k0_oneRun(), "ppr_extra": PPR_RPR_1EC1}),
    ("V_K3_pPr_rPr_AND_split", {"ind_xml": "", "runs_xml": runs_v_k1_splitRuns(), "ppr_extra": PPR_RPR_1EC1}),
    ("V_K4_no_run_kern", {"ind_xml": "", "runs_xml": runs_v_k3_no_kern_in_run()}),
    ("V_K5_full_1ec1_clone", {"ind_xml": "", "runs_xml": runs_v_k1_splitRuns(), "ppr_extra": PPR_RPR_1EC1}),
    # Variants with ind to study the +9pt mystery
    ("V_K6_ind_left_105_oneRun", {"ind_xml": '<w:ind w:left="105"/>', "runs_xml": runs_v_k0_oneRun()}),
    ("V_K7_ind_left_AND_leftChars_split", {"ind_xml": '<w:ind w:leftChars="50" w:left="105"/>', "runs_xml": runs_v_k1_splitRuns(), "ppr_extra": PPR_RPR_1EC1}),
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
            time.sleep(6.0)
    if word is None:
        print("Failed Word startup"); return
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
                results.append({"id": vid, "error": "render failed"})
                continue
            m = measure_box(pdf)
            if m:
                ex_a = m["search_x0"] - LEFT_MARGIN_PT
                ex_v = (m["leftmost_pt"] - LEFT_MARGIN_PT) if m["leftmost_pt"] else None
                vs = f"{ex_v:.2f}" if ex_v is not None else "NA"
                print(f"  search={m['search_x0']:.2f}pt visible={m['leftmost_pt']:.2f}pt | excess_adv={ex_a:.2f}pt visible_excess={vs}pt")
                results.append({"id": vid, "measurement": m, "excess_advance_pt": ex_a, "excess_visible_pt": ex_v})
            else:
                results.append({"id": vid, "error": "no glyph"})
    finally:
        try: word.Quit()
        except: pass
    out = os.path.join(OUT_DIR, "results.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved -> {out}")


if __name__ == "__main__":
    main()
