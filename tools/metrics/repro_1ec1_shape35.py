# -*- coding: utf-8 -*-
"""V_L: Replicate 1ec1 Shape 35 exactly to verify □ position formula.

Shape 35 properties (BOX[1-4]):
  prst=roundRect, adj=8396
  extent cx=6648450 EMU = 523.50pt, cy=1657350 = 130.50pt
  positionH: relativeFrom=margin, align=center
  distL=distR=114300 (9.00pt)
  lIns=rIns=91440 (7.20pt), tIns=bIns=0
  bodyPr: spcFirstLastPara=0, vertOverflow=overflow, horzOverflow=overflow,
          numCol=1, anchor=t, compatLnSpc=1, anchorCtr=0, forceAA=0
  Paragraph: snapToGrid=0, spacing line=480 exact, jc=left, NO ind
  Run font: majorEastAsia/MS Gothic at sz=28, kern=2, hint=eastAsia

Predict:
  shape_left = 42.55 + (510.20 - 523.50) / 2 = 35.90pt
  content_origin = shape_left + lIns = 35.90 + 7.20 = 43.10pt
  + roundRect inset for adj=8396 (TBD)
"""
import os, sys, time, json, zipfile, shutil, tempfile
import pythoncom, win32com.client
import fitz

sys.path.insert(0, "tools/metrics")
from repro_1ec1_textbox_ind import CTYPES, RELS_ROOT, WORD_RELS

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath("pipeline_data/1ec1_shape35")
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


def doc_xml(*, prst="roundRect", adj=8396, extent_cx=6648450, extent_cy=1657350,
            lins_emu=91440, ind_xml="", spacing_xml='<w:spacing w:line="480" w:lineRule="exact"/>',
            text="□１"):
    adj_xml = f'<a:avLst><a:gd name="adj" fmla="val {adj}"/></a:avLst>' if adj else '<a:avLst/>'
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
<w:body>
<w:p><w:r><w:t>BodyPara1</w:t></w:r></w:p>
<w:p>
  <w:r>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"
                     relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
            <wp:extent cx="{extent_cx}" cy="{extent_cy}"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapNone/>
            <wp:docPr id="35" name="Shape35"/>
            <wp:cNvGraphicFramePr/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr/>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="{extent_cx}" cy="{extent_cy}"/></a:xfrm>
                    <a:prstGeom prst="{prst}">{adj_xml}</a:prstGeom>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p>
                        <w:pPr>
                          <w:snapToGrid w:val="0"/>
                          {spacing_xml}
                          <w:jc w:val="left"/>
                          {ind_xml}
                          <w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック" w:cs="FrankRuehl"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
                        </w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック" w:cs="FrankRuehl" w:hint="eastAsia"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t>{text}</w:t></w:r>
                      </w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" spcFirstLastPara="0" vertOverflow="overflow" horzOverflow="overflow" vert="horz" wrap="square" lIns="{lins_emu}" tIns="0" rIns="{lins_emu}" bIns="0" numCol="1" spcCol="0" rtlCol="0" fromWordArt="0" anchor="t" anchorCtr="0" forceAA="0" compatLnSpc="1"/>
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
    tmp = tempfile.mkdtemp(prefix="vl_")
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
    return {"search_x0": inst.x0, "leftmost_pt": leftmost / zoom if leftmost else None}


VARIANTS = [
    ("V_L0_shape35_clone_no_ind", {}),
    ("V_L1_shape35_with_ind_105", {"ind_xml": '<w:ind w:left="105"/>'}),
    ("V_L2_shape35_ind_left_AND_chars", {"ind_xml": '<w:ind w:leftChars="50" w:left="105"/>'}),
    ("V_L3_shape35_ind_chars_only", {"ind_xml": '<w:ind w:leftChars="50"/>'}),
    ("V_L4_shape35_rect_no_round", {"prst": "rect", "adj": None}),
    ("V_L5_shape35_lIns_2.835", {"lins_emu": 36000}),
    ("V_L6_shape35_adj_zero", {"adj": 0}),
    ("V_L7_shape9_clone_no_ind", {"adj": 4015, "extent_cx": 6638925, "extent_cy": 3028950, "lins_emu": 36000}),
    ("V_L8_shape9_with_ind_105", {"adj": 4015, "extent_cx": 6638925, "extent_cy": 3028950, "lins_emu": 36000, "ind_xml": '<w:ind w:leftChars="50" w:left="105"/>'}),
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
    print(f"1ec1 actual:")
    print(f"  Shape35 BOX[1] no ind: visible 47.75pt advance 46.08pt")
    print(f"  Shape35 BOX[3] no ind: visible 47.75pt advance 46.08pt")
    print(f"  Shape9 BOX[5] ind=105+chars50: visible 57.00pt advance 55.32pt")
    print()
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
