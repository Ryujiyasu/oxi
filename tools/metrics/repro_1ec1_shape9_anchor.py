# -*- coding: utf-8 -*-
"""V_N: Shape 9 anchor attribute test — vary positionV posOffset, effectExtent,
relativeHeight, wp14 attrs."""
import os, sys, time, json, zipfile, shutil, tempfile
import pythoncom, win32com.client
import fitz

sys.path.insert(0, "tools/metrics")
from repro_1ec1_textbox_ind import CTYPES, RELS_ROOT, WORD_RELS

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath("pipeline_data/1ec1_shape9_anchor")
os.makedirs(OUT_DIR, exist_ok=True)

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:defaultTabStop w:val="840"/>
<w:characterSpacingControl w:val="compressPunctuation"/>
<w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>
</w:compat>
</w:settings>'''


def doc_xml(*, posOffset=0, effExt='l="0" t="0" r="0" b="0"', relHeight=1,
            wp14_attrs="", positionH_relativeFrom="margin", positionH_child='<wp:align>center</wp:align>',
            distL=114300, distR=114300):
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
<w:body>
<w:p><w:r><w:t>BodyPara1</w:t></w:r></w:p>
<w:p>
  <w:r>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="{distL}" distR="{distR}" simplePos="0" relativeHeight="{relHeight}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" {wp14_attrs}>
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="{positionH_relativeFrom}">{positionH_child}</wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>{posOffset}</wp:posOffset></wp:positionV>
            <wp:extent cx="6638925" cy="3028950"/>
            <wp:effectExtent {effExt}/>
            <wp:wrapNone/>
            <wp:docPr id="9" name="Shape9"/>
            <wp:cNvGraphicFramePr/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr/>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="6638925" cy="3028950"/></a:xfrm>
                    <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 4015"/></a:avLst></a:prstGeom>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p>
                        <w:pPr>
                          <w:snapToGrid w:val="0"/>
                          <w:spacing w:line="440" w:lineRule="exact"/>
                          <w:jc w:val="left"/>
                          <w:ind w:leftChars="50" w:left="105"/>
                          <w:rPr><w:rFonts w:asciiTheme="majorEastAsia" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorEastAsia" w:cs="FrankRuehl"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
                        </w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック" w:cs="FrankRuehl" w:hint="eastAsia"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t>□３</w:t></w:r>
                      </w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" spcFirstLastPara="0" vertOverflow="overflow" horzOverflow="overflow" vert="horz" wrap="square" lIns="36000" tIns="0" rIns="36000" bIns="0" numCol="1" spcCol="0" rtlCol="0" fromWordArt="0" anchor="t" anchorCtr="0" forceAA="0" compatLnSpc="1"/>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
    </mc:AlternateContent>
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
    tmp = tempfile.mkdtemp(prefix="vn_")
    try:
        files = [
            ("[Content_Types].xml", CTYPES),
            ("_rels/.rels", RELS_ROOT),
            ("word/_rels/document.xml.rels", WORD_RELS),
            ("word/styles.xml", STYLES),
            ("word/settings.xml", SETTINGS),
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
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
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
    return {"search_x0": inst.x0, "leftmost_pt": leftmost / zoom if leftmost else None}


VARIANTS = [
    ("V_N0_baseline_my_default", {}),
    ("V_N1_posOffset_231140", {"posOffset": 231140}),
    ("V_N2_effectExtent_real", {"effExt": 'l="0" t="0" r="28575" b="19050"'}),
    ("V_N3_relHeight_real", {"relHeight": 251670528}),
    ("V_N4_wp14_anchorId", {"wp14_attrs": 'wp14:anchorId="3140AB3F" wp14:editId="18F65ABE"'}),
    ("V_N5_all_real", {"posOffset": 231140, "effExt": 'l="0" t="0" r="28575" b="19050"', "relHeight": 251670528, "wp14_attrs": 'wp14:anchorId="3140AB3F" wp14:editId="18F65ABE"'}),
    ("V_N6_distL_distR_zero", {"distL": 0, "distR": 0}),
    ("V_N7_positionH_page_align_center", {"positionH_relativeFrom": "page"}),
    ("V_N8_positionH_column", {"positionH_relativeFrom": "column"}),
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
    print(f"Target: 1ec1 Shape 9 BOX[5] advance 55.32pt visible 57.00pt")
    print(f"Current V_M0 baseline: advance 47.16pt → 8.16pt gap")
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
                results.append({"id": vid, "kwargs": kwargs, "measurement": m,
                                "excess_advance_pt": ex_a, "excess_visible_pt": ex_v})
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
