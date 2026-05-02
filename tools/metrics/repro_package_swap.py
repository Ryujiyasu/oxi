# -*- coding: utf-8 -*-
"""V_DD: Test if 1ec1's package parts (theme, fontTable, styles) carry +8.76pt.

Build V_Q-style synthetic Shape 9 clone (gave 2.78pt match) but use 1ec1's actual
theme/styles/fontTable/settings. If gives 11.54pt → package matters."""
import os, sys, time, json, zipfile, shutil, tempfile, re
import pythoncom, win32com.client as wc
import fitz

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ORIG_DOCX = os.path.abspath("tools/golden-test/documents/docx/1ec1091177b1_006.docx")
OUT_DIR = os.path.abspath("pipeline_data/1ec1_package_swap")
os.makedirs(OUT_DIR, exist_ok=True)


# Synthetic minimal document body (just BodyPara1 + Shape 9 with □３)
SYNTH_DOC = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
 xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
<w:body>
<w:p><w:r><w:t>BodyPara1</w:t></w:r></w:p>
<w:p>
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
                        <w:r><w:rPr><w:rFonts w:asciiTheme="majorEastAsia" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorEastAsia" w:cs="FrankRuehl" w:hint="eastAsia"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t>□３</w:t></w:r>
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


def build_with_1ec1_package(out_path, *, replace_document=True, drop_files=None):
    """Take 1ec1's package, replace document.xml with synthetic minimal."""
    drop_files = drop_files or []
    tmp = tempfile.mkdtemp(prefix='pkg_')
    try:
        with zipfile.ZipFile(ORIG_DOCX) as z:
            z.extractall(tmp)
        if replace_document:
            with open(os.path.join(tmp, 'word', 'document.xml'), 'w', encoding='utf-8') as f:
                f.write(SYNTH_DOC)
        # Remove specific files
        for fname in drop_files:
            full = os.path.join(tmp, fname.replace('/', os.sep))
            if os.path.exists(full):
                os.remove(full)
        # Update Content_Types if files dropped
        # Skip for now — should still work for testing
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
            for root, _, names in os.walk(tmp):
                for fn in names:
                    full = os.path.join(root, fn)
                    arc = os.path.relpath(full, tmp).replace(os.sep, '/')
                    z.write(full, arc)
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
    print(f"  ERR: {last}")
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
    ("V_DD0_1ec1_package_synth_doc", {"replace_document": True}),
    ("V_DD1_drop_theme", {"replace_document": True, "drop_files": ["word/theme/theme1.xml"]}),
    ("V_DD2_drop_fontTable", {"replace_document": True, "drop_files": ["word/fontTable.xml"]}),
    ("V_DD3_drop_styles_use_minimal", {"replace_document": True, "drop_files": ["word/styles.xml"]}),
    ("V_DD4_drop_settings", {"replace_document": True, "drop_files": ["word/settings.xml"]}),
    ("V_DD5_drop_image", {"replace_document": True, "drop_files": ["word/media/image1.png"]}),
    ("V_DD6_drop_endnotes_footnotes", {"replace_document": True, "drop_files": ["word/endnotes.xml", "word/footnotes.xml"]}),
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
    print("If V_DD0 (1ec1's full package + synthetic Shape 9) gives 55.32pt → package carries cache")
    print("If V_DD0 gives 47.16pt (matches V_O3 pure synth) → package doesn't matter\n")
    results = []
    try:
        for vid, kwargs in VARIANTS:
            print(f"=== {vid} ===")
            docx = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.docx"))
            pdf = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.pdf"))
            for f in (docx, pdf):
                try: os.remove(f)
                except: pass
            try:
                build_with_1ec1_package(docx, **kwargs)
            except Exception as e:
                print(f"  build failed: {e}")
                continue
            ok = render_pdf(word, docx, pdf)
            if not ok:
                results.append({"id": vid, "error": "render"})
                continue
            boxes = measure(pdf)
            for b in boxes[:4]:
                print(f"    □ x={b['x']:.2f} y={b['y']:.2f} P{b['page']}")
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
