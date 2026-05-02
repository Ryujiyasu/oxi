"""Minimal repro: Word ind:left rule for textbox at positionH=margin/center
with extent overflow.

Build minimal OOXML with:
  - Page A4 (595.30 × 841.90pt), margins L=R=42.55pt → avail=510.20pt
  - DML floating textbox: positionH=margin/center, posOffset=0
  - extent W=522.75pt (overflows by +12.55pt → centers extending 6.275pt left)
  - bodyPr lIns=2.835pt rIns=2.835pt
  - Single paragraph with single CJK char `□` (14pt MS Gothic)
  - Vary <w:ind w:left="..."> across variants

Predicted positions for content per ECMA-376:
  shape_left = 42.55 - (522.75 - 510.20)/2 = 36.275pt
  content_x_oxi = shape_left + lIns + ind:left
                = 36.275 + 2.835 + ind:left

Variants:
  V0  ind=0           → predicted 39.11pt
  V1  ind=2.5pt (50tw) → predicted 41.61pt
  V2  ind=5.25pt (105tw) → predicted 44.36pt (matches 1ec1 Shape 4)
  V3  ind=10pt (200tw)  → predicted 49.11pt
  V4  ind=20pt (400tw)  → predicted 59.11pt

If Word ignores ind:left → all V0-V4 produce SAME content_x
If Word applies normally → content_x increases by ind:left amount
If Word compensates shape_left → content_x stays same as V0

Also test variants with leftChars only (no left twip):
  V5  leftChars=50 only → tests ECMA-376 leftChars-overrides-left
  V6  leftChars=100 only

And no-overflow control:
  V7  extent=400pt (no overflow), ind=10pt → standard centering w/ ind
"""
import os
import sys
import time
import json
import zipfile
import shutil
import tempfile
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath("pipeline_data/r35_1ec1_repro_docs")
RESULT = os.path.abspath("pipeline_data/r35_1ec1_repro_2026-05-02.json")

CTYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

RELS_ROOT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

WORD_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="ＭＳ ゴシック" w:eastAsia="ＭＳ ゴシック" w:hAnsi="ＭＳ ゴシック"/>
<w:kern w:val="2"/><w:sz w:val="28"/>
<w:lang w:val="en-US" w:eastAsia="ja-JP"/>
</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/></w:style>
</w:styles>"""

SETTINGS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>
</w:settings>"""


def doc_xml(extent_emu, lins_emu, ind_left_tw=None, ind_left_chars=None):
    """Build document.xml with textbox parameters."""
    ind_xml = ""
    if ind_left_tw is not None or ind_left_chars is not None:
        attrs = []
        if ind_left_chars is not None:
            attrs.append(f'w:leftChars="{ind_left_chars}"')
        if ind_left_tw is not None:
            attrs.append(f'w:left="{ind_left_tw}"')
        ind_xml = f'<w:ind {" ".join(attrs)}/>'

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
            <wp:extent cx="{extent_emu}" cy="600000"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapNone/>
            <wp:docPr id="9" name="Repro Shape"/>
            <wp:cNvGraphicFramePr/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr/>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="{extent_emu}" cy="600000"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p>
                        <w:pPr>
                          <w:snapToGrid w:val="0"/>
                          {ind_xml}
                          <w:jc w:val="left"/>
                        </w:pPr>
                        <w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>□3</w:t></w:r>
                      </w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" wrap="square" lIns="{lins_emu}" tIns="0" rIns="{lins_emu}" bIns="0" anchor="t"/>
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
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="720" w:footer="720" w:gutter="0"/>
<w:cols w:space="720"/>
<w:docGrid w:type="lines" w:linePitch="360"/>
</w:sectPr>
</w:body>
</w:document>'''


def write_docx(path, extent_emu, lins_emu, ind_left_tw=None, ind_left_chars=None):
    tmp = tempfile.mkdtemp(prefix="reproind_")
    try:
        os.makedirs(os.path.join(tmp, "_rels"), exist_ok=True)
        os.makedirs(os.path.join(tmp, "word", "_rels"), exist_ok=True)
        files = [
            ("[Content_Types].xml", CTYPES),
            ("_rels/.rels", RELS_ROOT),
            ("word/_rels/document.xml.rels", WORD_RELS),
            ("word/styles.xml", STYLES),
            ("word/settings.xml", SETTINGS),
            ("word/document.xml", doc_xml(extent_emu, lins_emu, ind_left_tw, ind_left_chars)),
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


# Variant: (label, extent_emu, lins_emu, ind_tw, ind_chars, predicted_oxi_x)
# 1pt = 12700 EMU. 522.75pt = 6638925 EMU (matches Shape 4)
# 2.835pt = 36000 EMU (matches Shape 4 lIns)
# Page width 11906tw=595.3pt, margin 851tw=42.55pt, avail=510.2pt
EXT_OVERFLOW_EMU = 6638925  # 522.75pt → overflow +12.55pt
EXT_NO_OVERFLOW_EMU = 5080000  # 400pt → no overflow
LINS_2_835 = 36000
LINS_7_2 = 91440

VARIANTS = [
    # Series A: overflow extent, vary ind:left in twips, lIns=2.835pt
    ("V0_overflow_ind0", EXT_OVERFLOW_EMU, LINS_2_835, 0, None, 39.11),
    ("V1_overflow_ind50tw", EXT_OVERFLOW_EMU, LINS_2_835, 50, None, 41.61),
    ("V2_overflow_ind105tw", EXT_OVERFLOW_EMU, LINS_2_835, 105, None, 44.36),
    ("V3_overflow_ind200tw", EXT_OVERFLOW_EMU, LINS_2_835, 200, None, 49.11),
    ("V4_overflow_ind400tw", EXT_OVERFLOW_EMU, LINS_2_835, 400, None, 59.11),
    # Series B: leftChars only (test leftChars-overrides-left rule)
    ("V5_overflow_chars50", EXT_OVERFLOW_EMU, LINS_2_835, None, 50, None),
    ("V6_overflow_chars100", EXT_OVERFLOW_EMU, LINS_2_835, None, 100, None),
    # Series C: no-overflow control with same ind
    ("V7_nooverflow_ind105tw", EXT_NO_OVERFLOW_EMU, LINS_2_835, 105, None, None),
    ("V8_nooverflow_ind0", EXT_NO_OVERFLOW_EMU, LINS_2_835, 0, None, None),
    # Series D: lIns variation matching Shape 2 (lIns=7.2pt) with ind:0
    ("V9_overflow_lIns7_2_ind0", EXT_OVERFLOW_EMU, LINS_7_2, 0, None, None),
]


def measure(word, path):
    d = word.Documents.Open(path, ReadOnly=True)
    time.sleep(0.5)
    out = {}
    out["page_w"] = float(d.Sections(1).PageSetup.PageWidth)
    out["page_margin_l"] = float(d.Sections(1).PageSetup.LeftMargin)
    # Enumerate shapes
    n_shapes = d.Shapes.Count
    out["n_shapes"] = n_shapes
    for i in range(1, n_shapes + 1):
        s = d.Shapes(i)
        try:
            tf = s.TextFrame
            if tf and tf.HasText:
                tr = tf.TextRange
                p1 = tr.Paragraphs(1)
                pr = p1.Range
                first_chars = pr.Characters(1)
                fx = float(first_chars.Information(5))
                fy = float(first_chars.Information(6))
                out["shape_first_x"] = fx
                out["shape_first_y"] = fy
                out["shape_left"] = float(s.Left)
                out["shape_top"] = float(s.Top)
                out["shape_width"] = float(s.Width)
                out["margin_left"] = float(tf.MarginLeft)
                out["margin_right"] = float(tf.MarginRight)
                out["paragraph_li"] = float(pr.ParagraphFormat.LeftIndent)
                out["paragraph_fli"] = float(pr.ParagraphFormat.FirstLineIndent)
        except Exception as e:
            out[f"shape_{i}_err"] = str(e)
    d.Close(SaveChanges=False)
    return out


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    results = {}
    for label, ext_emu, lins_emu, ind_tw, ind_chars, pred in VARIANTS:
        path = os.path.join(OUT_DIR, f"{label}.docx")
        write_docx(path, ext_emu, lins_emu, ind_tw, ind_chars)
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        try:
            try:
                m = measure(word, path)
            except Exception as e:
                results[label] = {"error": str(e)}
                print(f"[{label}] ERR: {e}", flush=True)
                continue
            ext_pt = ext_emu / 12700
            lins_pt = lins_emu / 12700
            avail = m["page_w"] - 2 * m["page_margin_l"]
            shape_left_pred = m["page_margin_l"] - max(0, ext_pt - avail) / 2
            ind_pt = (ind_tw or 0) / 20.0
            content_x_pred = shape_left_pred + lins_pt + ind_pt
            results[label] = {
                "ext_pt": ext_pt, "lins_pt": lins_pt,
                "ind_tw": ind_tw, "ind_chars": ind_chars,
                "ind_pt_oxi_calc": ind_pt,
                "shape_left_predicted": round(shape_left_pred, 4),
                "content_x_predicted_oxi": round(content_x_pred, 4),
                **m,
            }
            print(f"\n[{label}] ext={ext_pt:.2f} lIns={lins_pt:.3f} "
                  f"ind={ind_tw}tw/{ind_chars}chars",
                  flush=True)
            print(f"  Predicted shape_left={shape_left_pred:.3f} "
                  f"content_x_oxi={content_x_pred:.3f}", flush=True)
            print(f"  Word measured: shape_left={m.get('shape_left')} "
                  f"shape_first_x={m.get('shape_first_x')} "
                  f"li={m.get('paragraph_li'):.3f} margins L/R={m.get('margin_left'):.3f}/{m.get('margin_right'):.3f}",
                  flush=True)
        finally:
            try: word.Quit()
            except: pass
            time.sleep(0.8)

    with open(RESULT, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2, default=str)
    print(f"\nWrote {RESULT}", flush=True)


if __name__ == "__main__":
    main()
