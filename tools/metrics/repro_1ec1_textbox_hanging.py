"""Verify whether Word's textbox ind:left ignore-rule extends to
paragraphs with w:hanging (list-like indents).

Hypothesis options:
  A: Word ignores ALL w:ind w:left (incl. with hanging)
  B: Word applies ind:left when paragraph has w:hanging
  C: Word applies ind:left only when paragraph also has list (numPr)

Tests (using same overflow textbox setup as previous repro):
  H0  ind=0                          → control (x=39.0)
  H1  left=840tw + hanging=210tw    → list-like body indent 42pt, first-line 31.5pt
  H2  left=840tw + firstLine=210tw  → first-line=42+10.5pt, body 42pt
  H3  left=840tw alone               → already proven ignored (x=39.0 expected)
  H4  hanging=210tw alone            → no left indent
  H5  left=840tw + hanging=210tw + numPr → full list paragraph

Expected interpretations:
  If B (hanging coexist applies): H1 first char x = 36.275 + 2.835 + (42 - 10.5) = 70.61pt
  If A (always ignore): H1 x ≈ 39.0pt
"""
import os, sys, time, json, zipfile, shutil, tempfile
import win32com.client
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath("pipeline_data/r35_1ec1_hanging_docs")
RESULT = os.path.abspath("pipeline_data/r35_1ec1_hanging_2026-05-02.json")

# Reuse boilerplate from main repro
sys.path.insert(0, "tools/metrics")
from repro_1ec1_textbox_ind import (CTYPES, RELS_ROOT, WORD_RELS, STYLES,
    SETTINGS, EXT_OVERFLOW_EMU, LINS_2_835, measure)


def doc_xml_hanging(extent_emu, lins_emu, ind_attrs, has_numpr=False):
    """ind_attrs: dict like {'left': 840, 'hanging': 210}"""
    attr_str = " ".join(f'w:{k}="{v}"' for k, v in ind_attrs.items())
    ind_xml = f'<w:ind {attr_str}/>' if attr_str else ""
    numpr = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' if has_numpr else ""

    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
<w:body>
<w:p><w:r><w:t>BodyPara1</w:t></w:r></w:p>
<w:p>
<w:r><mc:AlternateContent><mc:Choice Requires="wps">
<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"
relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
<wp:simplePos x="0" y="0"/>
<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>
<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
<wp:extent cx="{extent_emu}" cy="600000"/>
<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>
<wp:docPr id="9" name="HangingRepro"/>
<wp:cNvGraphicFramePr/>
<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
<wps:wsp><wps:cNvSpPr/><wps:spPr>
<a:xfrm><a:off x="0" y="0"/><a:ext cx="{extent_emu}" cy="600000"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
</wps:spPr>
<wps:txbx><w:txbxContent>
<w:p>
<w:pPr>
<w:snapToGrid w:val="0"/>
{numpr}
{ind_xml}
<w:jc w:val="left"/>
</w:pPr>
<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>□3</w:t></w:r>
</w:p>
</w:txbxContent></wps:txbx>
<wps:bodyPr rot="0" wrap="square" lIns="{lins_emu}" tIns="0" rIns="{lins_emu}" bIns="0" anchor="t"/>
</wps:wsp></a:graphicData></a:graphic>
</wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>
</w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="720" w:footer="720" w:gutter="0"/>
<w:cols w:space="720"/><w:docGrid w:type="lines" w:linePitch="360"/>
</w:sectPr></w:body></w:document>'''


def write_docx(path, ext, lins, ind_attrs, has_numpr=False):
    tmp = tempfile.mkdtemp(prefix="hangrepro_")
    try:
        os.makedirs(os.path.join(tmp, "_rels"), exist_ok=True)
        os.makedirs(os.path.join(tmp, "word", "_rels"), exist_ok=True)
        files = [
            ("[Content_Types].xml", CTYPES),
            ("_rels/.rels", RELS_ROOT),
            ("word/_rels/document.xml.rels", WORD_RELS),
            ("word/styles.xml", STYLES),
            ("word/settings.xml", SETTINGS),
            ("word/document.xml", doc_xml_hanging(ext, lins, ind_attrs, has_numpr)),
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


VARIANTS = [
    # H0: control (= V0 from previous repro)
    ("H0_ind0", {}),
    # H1: left + hanging (1ec1 Shape 4 list paragraphs use this)
    ("H1_left840_hanging210", {"left": 840, "hanging": 210}),
    # H2: left + firstLine (positive first-line)
    ("H2_left840_firstLine210", {"left": 840, "firstLine": 210}),
    # H3: left only (already proven ignored — control)
    ("H3_left840_only", {"left": 840}),
    # H4: hanging only (no left)
    ("H4_hanging210_only", {"hanging": 210}),
]


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    results = {}
    for label, ind_attrs in VARIANTS:
        path = os.path.join(OUT_DIR, f"{label}.docx")
        write_docx(path, EXT_OVERFLOW_EMU, LINS_2_835, ind_attrs)
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False; word.DisplayAlerts = False
        try:
            try:
                m = measure(word, path)
            except Exception as e:
                results[label] = {"error": str(e)}
                print(f"[{label}] ERR: {e}", flush=True)
                continue
            results[label] = m
            print(f"[{label}] ind={ind_attrs} → "
                  f"shape_first_x={m.get('shape_first_x')} "
                  f"li={m.get('paragraph_li')} "
                  f"fli={m.get('paragraph_fli')}",
                  flush=True)
        finally:
            try: word.Quit()
            except: pass
            time.sleep(1.0)

    with open(RESULT, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2, default=str)
    print(f"\nWrote {RESULT}")


if __name__ == "__main__":
    main()
