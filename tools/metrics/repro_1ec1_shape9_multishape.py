# -*- coding: utf-8 -*-
"""V_P: Test if multi-shape document context (Shape 35 before Shape 9) shifts BOX[5].

Hypothesis: Word may apply different layout when multiple floating shapes exist
in the same document. Test by adding Shape 35 above Shape 9.
"""
import os, sys, time, json, zipfile, shutil, tempfile
import pythoncom, win32com.client
import fitz

sys.path.insert(0, "tools/metrics")
from repro_1ec1_textbox_ind import CTYPES, RELS_ROOT, WORD_RELS

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath("pipeline_data/1ec1_shape9_multishape")
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


SHAPE_35_BLOCK = '''<w:p>
  <w:r>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251660288" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
            <wp:extent cx="6648450" cy="1657350"/>
            <wp:effectExtent l="0" t="0" r="19050" b="19050"/>
            <wp:wrapNone/>
            <wp:docPr id="35" name="Shape35"/>
            <wp:cNvGraphicFramePr/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr/>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="6648450" cy="1657350"/></a:xfrm>
                    <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 8396"/></a:avLst></a:prstGeom>
                    <a:solidFill><a:sysClr val="window" lastClr="FFFFFF"/></a:solidFill>
                    <a:ln w="12700"><a:solidFill><a:sysClr val="windowText" lastClr="000000"/></a:solidFill></a:ln>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p>
                        <w:pPr>
                          <w:snapToGrid w:val="0"/>
                          <w:spacing w:line="480" w:lineRule="exact"/>
                          <w:jc w:val="left"/>
                        </w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック" w:hint="eastAsia"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t>□１　Shape 35 paragraph</w:t></w:r>
                      </w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" spcFirstLastPara="0" vertOverflow="overflow" horzOverflow="overflow" wrap="square" lIns="91440" tIns="0" rIns="91440" bIns="0" anchor="t" compatLnSpc="1"/>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
    </mc:AlternateContent>
  </w:r>
</w:p>'''

SHAPE_9_BLOCK = '''<w:p>
  <w:pPr><w:rPr><w:rFonts w:asciiTheme="minorEastAsia" w:hAnsiTheme="minorEastAsia"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:pPr>
  <w:r>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251670528" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>231140</wp:posOffset></wp:positionV>
            <wp:extent cx="6638925" cy="3028950"/>
            <wp:effectExtent l="0" t="0" r="28575" b="19050"/>
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
                    <a:solidFill><a:sysClr val="window" lastClr="FFFFFF"/></a:solidFill>
                    <a:ln w="12700"><a:solidFill><a:sysClr val="windowText" lastClr="000000"/></a:solidFill></a:ln>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p>
                        <w:pPr>
                          <w:snapToGrid w:val="0"/>
                          <w:spacing w:line="440" w:lineRule="exact"/>
                          <w:ind w:leftChars="50" w:left="105"/>
                          <w:jc w:val="left"/>
                          <w:rPr><w:rFonts w:asciiTheme="majorEastAsia" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorEastAsia" w:cs="FrankRuehl"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
                        </w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック" w:cs="FrankRuehl" w:hint="eastAsia"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t>□３</w:t></w:r>
                      </w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" spcFirstLastPara="0" vertOverflow="overflow" horzOverflow="overflow" wrap="square" lIns="36000" tIns="0" rIns="36000" bIns="0" anchor="t" compatLnSpc="1"/>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
    </mc:AlternateContent>
  </w:r>
</w:p>'''

# Filler text paragraphs
FILLER_BODY = ''.join([
    f'<w:p><w:r><w:t>Filler line {i}</w:t></w:r></w:p>'
    for i in range(8)
])


def doc_xml(*, scenario):
    if scenario == "shape9_only":
        body_content = '<w:p><w:r><w:t>BodyPara1</w:t></w:r></w:p>' + SHAPE_9_BLOCK
    elif scenario == "shape35_then_shape9":
        body_content = '<w:p><w:r><w:t>BodyPara1</w:t></w:r></w:p>' + SHAPE_35_BLOCK + SHAPE_9_BLOCK
    elif scenario == "shape35_filler_shape9":
        body_content = '<w:p><w:r><w:t>BodyPara1</w:t></w:r></w:p>' + SHAPE_35_BLOCK + FILLER_BODY + SHAPE_9_BLOCK
    elif scenario == "filler_then_shape9":
        body_content = '<w:p><w:r><w:t>BodyPara1</w:t></w:r></w:p>' + FILLER_BODY + SHAPE_9_BLOCK
    elif scenario == "shape35_alone":
        body_content = '<w:p><w:r><w:t>BodyPara1</w:t></w:r></w:p>' + SHAPE_35_BLOCK
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
<w:body>
{body_content}
<w:sectPr>
<w:pgSz w:w="11906" w:h="16838" w:code="9"/>
<w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992" w:gutter="0"/>
<w:cols w:space="425"/>
<w:docGrid w:type="linesAndChars" w:linePitch="357"/>
</w:sectPr>
</w:body>
</w:document>'''


def write_docx(path, **kwargs):
    tmp = tempfile.mkdtemp(prefix="vp_")
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


def measure_all_box(pdf_path):
    """Return positions of ALL □ instances on all pages."""
    d = fitz.open(pdf_path)
    results = []
    for pi in range(d.page_count):
        page = d[pi]
        instances = page.search_for("□")
        for inst in instances:
            results.append({
                "page": pi+1,
                "search_x0": inst.x0,
                "search_y0": inst.y0,
            })
    d.close()
    return results


VARIANTS = [
    ("V_P0_shape9_only", {"scenario": "shape9_only"}),
    ("V_P1_shape35_then_shape9", {"scenario": "shape35_then_shape9"}),
    ("V_P2_shape35_filler_shape9", {"scenario": "shape35_filler_shape9"}),
    ("V_P3_filler_then_shape9", {"scenario": "filler_then_shape9"}),
    ("V_P4_shape35_alone", {"scenario": "shape35_alone"}),
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
    print(f"Target: 1ec1 Shape 9 BOX[5] advance 55.32pt\n")
    results = []
    try:
        for vid, kwargs in VARIANTS:
            print(f"=== {vid} ===")
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
            boxes = measure_all_box(pdf)
            for i, b in enumerate(boxes):
                ex_a = b['search_x0'] - LEFT_MARGIN_PT
                marker = ' (Shape35)' if i == 0 and 'shape35' in vid else ' (Shape9)'
                print(f"  □#{i+1}: P{b['page']} x={b['search_x0']:.2f}pt y={b['search_y0']:.2f}pt | excess={ex_a:.2f}pt{marker}")
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
