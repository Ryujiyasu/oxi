# -*- coding: utf-8 -*-
"""V_FF: Self-authored minimal repro of the V_EE rule.

Build a fresh minimal OOXML with TWO floating shapes:
  Shape A (Shape 35 analogue): contains paragraph starting with □１
  Shape B (Shape 9 analogue): contains paragraph starting with □３ with w:ind w:left=105

If V_EE rule is real:
  - With Shape A present: Shape B's □ at ~55.32pt (= +14pt override of 5.25pt explicit)
  - Without Shape A (V_FF1): Shape B's □ at ~46.56pt (formula match)

Variants:
  V_FF0: Shape A (□１) + Shape B (□３ left=105) — expect trigger
  V_FF1: Shape B alone — expect formula match (46.56)
  V_FF2: Shape A (X１) + Shape B (□３ left=105) — non-□ in A, expect formula match (control)
  V_FF3: Shape A (□１) + Shape B (□３ no ind) — no explicit ind, expect base position
  V_FF4: Shape A (○１) + Shape B (○３ left=105) — different bullet char, test universality
  V_FF5: Shape A (■１) + Shape B (■３ left=105) — yet another bullet char
"""
import os, sys, time, json, zipfile, shutil, tempfile
import pythoncom, win32com.client as wc
import fitz

sys.path.insert(0, "tools/metrics")
from repro_1ec1_textbox_ind import CTYPES, RELS_ROOT, WORD_RELS

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath("pipeline_data/v_ff_minimal")
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


def shape_block(*, shape_id, name, cy_emu, lins_emu, adj, paragraphs_xml):
    return f'''<w:p>
  <w:r>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="{shape_id}51670528" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
            <wp:extent cx="6638925" cy="{cy_emu}"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapNone/>
            <wp:docPr id="{shape_id}" name="{name}"/>
            <wp:cNvGraphicFramePr/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr/>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="6638925" cy="{cy_emu}"/></a:xfrm>
                    <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val {adj}"/></a:avLst></a:prstGeom>
                    <a:solidFill><a:sysClr val="window" lastClr="FFFFFF"/></a:solidFill>
                    <a:ln w="12700"><a:solidFill><a:sysClr val="windowText" lastClr="000000"/></a:solidFill></a:ln>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      {paragraphs_xml}
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" spcFirstLastPara="0" vertOverflow="overflow" horzOverflow="overflow" wrap="square" lIns="{lins_emu}" tIns="0" rIns="{lins_emu}" bIns="0" anchor="t" compatLnSpc="1"/>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
    </mc:AlternateContent>
  </w:r>
</w:p>'''


def para(*, text, ind_xml=""):
    return f'''<w:p>
  <w:pPr>
    <w:snapToGrid w:val="0"/>
    <w:spacing w:line="440" w:lineRule="exact"/>
    <w:jc w:val="left"/>
    {ind_xml}
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック" w:hint="eastAsia"/>
      <w:sz w:val="28"/><w:szCs w:val="28"/>
    </w:rPr>
    <w:t>{text}</w:t>
  </w:r>
</w:p>'''


def doc_xml(*, body_blocks):
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
<w:body>
<w:p><w:r><w:t>BodyPara1</w:t></w:r></w:p>
{body_blocks}
<w:sectPr>
<w:pgSz w:w="11906" w:h="16838" w:code="9"/>
<w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992" w:gutter="0"/>
<w:cols w:space="425"/>
<w:docGrid w:type="linesAndChars" w:linePitch="357"/>
</w:sectPr>
</w:body>
</w:document>'''


def write_docx(path, **kwargs):
    tmp = tempfile.mkdtemp(prefix='ff_')
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
                    arc = os.path.relpath(full, tmp).replace(os.sep, '/')
                    z.write(full, arc)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# Shape A = analogue of 1ec1 Shape 35 (cy=130.50pt, lIns=7.20pt, adj=8396)
# Shape B = analogue of 1ec1 Shape 9 (cy=238.50pt, lIns=2.835pt, adj=4015)


def build_v_ff(variant):
    if variant == "V_FF0_with_box1_in_A":
        body = shape_block(shape_id=35, name="ShapeA", cy_emu=1657350, lins_emu=91440, adj=8396,
                           paragraphs_xml=para(text="□１"))
        body += shape_block(shape_id=9, name="ShapeB", cy_emu=3028950, lins_emu=36000, adj=4015,
                            paragraphs_xml=para(text="□３", ind_xml='<w:ind w:leftChars="50" w:left="105"/>'))
    elif variant == "V_FF1_no_shape_A":
        body = shape_block(shape_id=9, name="ShapeB", cy_emu=3028950, lins_emu=36000, adj=4015,
                           paragraphs_xml=para(text="□３", ind_xml='<w:ind w:leftChars="50" w:left="105"/>'))
    elif variant == "V_FF2_X1_in_A":
        body = shape_block(shape_id=35, name="ShapeA", cy_emu=1657350, lins_emu=91440, adj=8396,
                           paragraphs_xml=para(text="X１"))
        body += shape_block(shape_id=9, name="ShapeB", cy_emu=3028950, lins_emu=36000, adj=4015,
                            paragraphs_xml=para(text="□３", ind_xml='<w:ind w:leftChars="50" w:left="105"/>'))
    elif variant == "V_FF3_no_ind_in_B":
        body = shape_block(shape_id=35, name="ShapeA", cy_emu=1657350, lins_emu=91440, adj=8396,
                           paragraphs_xml=para(text="□１"))
        body += shape_block(shape_id=9, name="ShapeB", cy_emu=3028950, lins_emu=36000, adj=4015,
                            paragraphs_xml=para(text="□３"))
    elif variant == "V_FF4_circle_bullet":
        body = shape_block(shape_id=35, name="ShapeA", cy_emu=1657350, lins_emu=91440, adj=8396,
                           paragraphs_xml=para(text="○１"))
        body += shape_block(shape_id=9, name="ShapeB", cy_emu=3028950, lins_emu=36000, adj=4015,
                            paragraphs_xml=para(text="○３", ind_xml='<w:ind w:leftChars="50" w:left="105"/>'))
    elif variant == "V_FF5_filled_square":
        body = shape_block(shape_id=35, name="ShapeA", cy_emu=1657350, lins_emu=91440, adj=8396,
                           paragraphs_xml=para(text="■１"))
        body += shape_block(shape_id=9, name="ShapeB", cy_emu=3028950, lins_emu=36000, adj=4015,
                            paragraphs_xml=para(text="■３", ind_xml='<w:ind w:leftChars="50" w:left="105"/>'))
    return doc_xml(body_blocks=body)


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


def measure(pdf, target_chars):
    d = fitz.open(pdf)
    res = {}
    for ch in target_chars:
        positions = []
        for pi in range(d.page_count):
            for inst in d[pi].search_for(ch):
                positions.append({"page": pi+1, "x": inst.x0, "y": inst.y0})
        res[ch] = positions
    d.close()
    return res


VARIANTS = [
    ("V_FF0_with_box1_in_A", "□"),
    ("V_FF1_no_shape_A", "□"),
    ("V_FF2_X1_in_A", "□"),
    ("V_FF3_no_ind_in_B", "□"),
    ("V_FF4_circle_bullet", "○"),
    ("V_FF5_filled_square", "■"),
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
    print("V_FF minimal repro of cross-shape bullet trigger.")
    print("Per V_EE: Shape A with □ + Shape B with □+w:left=105 → Shape B □ at 55.32 (trigger)")
    print("If V_FF0 gives 55.32 and V_FF1 gives 46.56 → rule confirmed on minimal docs\n")
    results = []
    try:
        for vid, char in VARIANTS:
            print(f"=== {vid} ===")
            docx = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.docx"))
            pdf = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.pdf"))
            for f in (docx, pdf):
                try: os.remove(f)
                except: pass
            doc = build_v_ff(vid)
            tmp = tempfile.mkdtemp(prefix='ff_')
            try:
                files = [
                    ("[Content_Types].xml", CTYPES),
                    ("_rels/.rels", RELS_ROOT),
                    ("word/_rels/document.xml.rels", WORD_RELS),
                    ("word/styles.xml", STYLES),
                    ("word/settings.xml", SETTINGS),
                    ("word/document.xml", doc),
                ]
                for relpath, content in files:
                    full = os.path.join(tmp, relpath.replace("/", os.sep))
                    os.makedirs(os.path.dirname(full), exist_ok=True)
                    with open(full, "w", encoding="utf-8") as f:
                        f.write(content)
                with zipfile.ZipFile(docx, "w", zipfile.ZIP_DEFLATED) as z:
                    for root, _, names in os.walk(tmp):
                        for fn in names:
                            full = os.path.join(root, fn)
                            arc = os.path.relpath(full, tmp).replace(os.sep, '/')
                            z.write(full, arc)
            finally:
                shutil.rmtree(tmp, ignore_errors=True)
            ok = render_pdf(word, docx, pdf)
            if not ok:
                results.append({"id": vid, "error": "render"})
                continue
            positions = measure(pdf, [char])
            print(f"  {char} positions:")
            for p in positions[char]:
                print(f"    x={p['x']:.2f} y={p['y']:.2f} P{p['page']}")
            results.append({"id": vid, "positions": positions[char]})
    finally:
        try: word.Quit()
        except: pass
    out = os.path.join(OUT_DIR, "results.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved -> {out}")


if __name__ == "__main__":
    main()
