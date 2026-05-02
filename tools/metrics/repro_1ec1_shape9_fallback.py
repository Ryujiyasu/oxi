# -*- coding: utf-8 -*-
"""V_R: Test if mc:Fallback presence causes 8.76pt extra inset for Shape 9.

Build V_O3 equivalent clone + embed 1ec1's actual mc:Fallback VML block.
If rendering shifts to 55.32pt (matching 1ec1), Fallback is the cause.
If stays at 46.56pt (= V_O3), Fallback is irrelevant.
"""
import os, sys, time, json, zipfile, shutil, tempfile
import pythoncom, win32com.client
import fitz

sys.path.insert(0, "tools/metrics")
from repro_1ec1_textbox_ind import CTYPES, RELS_ROOT, WORD_RELS

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath("pipeline_data/1ec1_shape9_fallback_test")
os.makedirs(OUT_DIR, exist_ok=True)

# Load 1ec1 Shape 9 mc:Fallback block
with open('pipeline_data/1ec1_shape9_fallback.xml', encoding='utf-8') as f:
    REAL_FALLBACK = f.read()

# Override the v:roundrect's id to avoid collision (we'll use id="9" matching wp:docPr)
# But actually keep original to maximize fidelity

# Need wp14 namespace + extra ones used by VML
EXTRA_NS = '''xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'''

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


def doc_xml(*, include_fallback=False):
    fallback_xml = REAL_FALLBACK if include_fallback else ''
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
 {EXTRA_NS}>
<w:body>
<w:p><w:r><w:t>BodyPara1</w:t></w:r></w:p>
<w:p>
  <w:pPr><w:rPr><w:rFonts w:asciiTheme="minorEastAsia" w:hAnsiTheme="minorEastAsia"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:pPr>
  <w:r>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251670528" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" wp14:anchorId="3140AB3F" wp14:editId="18F65ABE">
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
                    <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:sysClr val="windowText" lastClr="000000"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>
                    <a:effectLst/>
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
                  <wps:bodyPr rot="0" spcFirstLastPara="0" vertOverflow="overflow" horzOverflow="overflow" vert="horz" wrap="square" lIns="36000" tIns="0" rIns="36000" bIns="0" numCol="1" spcCol="0" rtlCol="0" fromWordArt="0" anchor="t" anchorCtr="0" forceAA="0" compatLnSpc="1"/>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
      {fallback_xml}
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
    tmp = tempfile.mkdtemp(prefix="vr_")
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
    d.close()
    return {"search_x0": inst.x0}


MINIMAL_VML_FALLBACK = '''<mc:Fallback>
<w:pict>
<v:roundrect id="Shape9V" o:spid="_x0000_s1029" style="position:absolute;margin-left:0;margin-top:18.2pt;width:522.75pt;height:238.5pt;z-index:251670528;visibility:visible;mso-wrap-style:square;mso-position-horizontal:center;mso-position-horizontal-relative:margin;v-text-anchor:top" arcsize="2631f" filled="t" fillcolor="white" stroked="t" strokecolor="black" strokeweight="1pt">
<v:textbox inset="2.835pt,0,2.835pt,0">
<w:txbxContent>
<w:p><w:pPr><w:ind w:leftChars="50" w:left="105"/></w:pPr><w:r><w:t>Fallback content</w:t></w:r></w:p>
</w:txbxContent>
</v:textbox>
</v:roundrect>
</w:pict>
</mc:Fallback>'''


def doc_xml_minimal_fb(*, include_fallback=False):
    fallback_xml = MINIMAL_VML_FALLBACK if include_fallback else ''
    return doc_xml(include_fallback=False).replace('</mc:Choice>\n      ', f'</mc:Choice>\n      {fallback_xml}')


VARIANTS = [
    ("V_R0_no_fallback", {"include_fallback": False}),
    ("V_R1_with_real_fallback", {"include_fallback": True}),
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
    print(f"Target: 1ec1 Shape 9 BOX[5] advance 55.32pt")
    print(f"V_O3 (no fallback) baseline: 46.56pt → 8.76pt gap to investigate\n")
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
            print(f"  built ({os.path.getsize(docx)} bytes)")
            ok = render_pdf(word, docx, pdf)
            if not ok:
                results.append({"id": vid, "error": "render failed"})
                continue
            m = measure_box(pdf)
            if m:
                ex_a = m["search_x0"] - LEFT_MARGIN_PT
                print(f"  search={m['search_x0']:.2f}pt | excess_adv={ex_a:.2f}pt")
                results.append({"id": vid, "kwargs": kwargs, "measurement": m, "excess_advance_pt": ex_a})
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
