"""Multi-paragraph textbox: does Word apply ind to 2nd+ paragraph?

Single-para repro showed Word ignores ind. But 1ec1 fix regression
suggests multi-paragraph behavior differs.

Hypothesis: Word applies ind for non-first paragraphs.

Test: textbox with 2 paragraphs:
  M0: P1 plain, P2 plain → both at base x (control)
  M1: P1 plain, P2 ind:left=42 hanging=10.5 → P2 indented?
  M2: P1 ind:left=42, P2 ind:left=42 hanging=10.5 → both indented?
  M3: P1 plain (□3-style), P2 hanging=10.5 only → just hanging?
"""
import os, sys, time, json, zipfile, shutil, tempfile
import win32com.client
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, "tools/metrics")
from repro_1ec1_textbox_ind import (CTYPES, RELS_ROOT, WORD_RELS, STYLES,
    SETTINGS, EXT_OVERFLOW_EMU, LINS_2_835, measure)

OUT_DIR = os.path.abspath("pipeline_data/r35_1ec1_multipara_docs")
RESULT = os.path.abspath("pipeline_data/r35_1ec1_multipara_2026-05-02.json")


def doc_xml_multi(extent_emu, lins_emu, p1_ind, p2_ind):
    def ind_str(ind):
        if not ind: return ""
        return "<w:ind " + " ".join(f'w:{k}="{v}"' for k, v in ind.items()) + "/>"
    p1ind_xml = ind_str(p1_ind)
    p2ind_xml = ind_str(p2_ind)
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
<wp:extent cx="{extent_emu}" cy="1200000"/>
<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>
<wp:docPr id="9" name="MultiRepro"/>
<wp:cNvGraphicFramePr/>
<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
<wps:wsp><wps:cNvSpPr/><wps:spPr>
<a:xfrm><a:off x="0" y="0"/><a:ext cx="{extent_emu}" cy="1200000"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
</wps:spPr>
<wps:txbx><w:txbxContent>
<w:p><w:pPr><w:snapToGrid w:val="0"/>{p1ind_xml}<w:jc w:val="left"/></w:pPr>
<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>□3</w:t></w:r></w:p>
<w:p><w:pPr><w:snapToGrid w:val="0"/>{p2ind_xml}<w:jc w:val="left"/></w:pPr>
<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>・XX</w:t></w:r></w:p>
</w:txbxContent></wps:txbx>
<wps:bodyPr rot="0" wrap="square" lIns="{lins_emu}" tIns="0" rIns="{lins_emu}" bIns="0" anchor="t"/>
</wps:wsp></a:graphicData></a:graphic>
</wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>
</w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="720" w:footer="720" w:gutter="0"/>
<w:cols w:space="720"/><w:docGrid w:type="lines" w:linePitch="360"/>
</w:sectPr></w:body></w:document>'''


def write_docx(path, ext, lins, p1_ind, p2_ind):
    tmp = tempfile.mkdtemp(prefix="multipara_")
    try:
        os.makedirs(os.path.join(tmp, "_rels"), exist_ok=True)
        os.makedirs(os.path.join(tmp, "word", "_rels"), exist_ok=True)
        files = [
            ("[Content_Types].xml", CTYPES),
            ("_rels/.rels", RELS_ROOT),
            ("word/_rels/document.xml.rels", WORD_RELS),
            ("word/styles.xml", STYLES),
            ("word/settings.xml", SETTINGS),
            ("word/document.xml", doc_xml_multi(ext, lins, p1_ind, p2_ind)),
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


def measure_both_paras(word, path):
    d = word.Documents.Open(path, ReadOnly=True)
    time.sleep(0.5)
    out = {"page_w": float(d.Sections(1).PageSetup.PageWidth),
           "page_margin_l": float(d.Sections(1).PageSetup.LeftMargin)}
    s = d.Shapes(1)
    tf = s.TextFrame
    tr = tf.TextRange
    paras = []
    for pi in range(1, tr.Paragraphs.Count + 1):
        p = tr.Paragraphs(pi)
        pr = p.Range
        first = pr.Characters(1)
        try:
            x = float(first.Information(5))
            y = float(first.Information(6))
        except:
            x = y = None
        paras.append({"i": pi, "text": pr.Text[:30],
                       "x": x, "y": y,
                       "li": float(pr.ParagraphFormat.LeftIndent),
                       "fli": float(pr.ParagraphFormat.FirstLineIndent)})
    out["paragraphs"] = paras
    out["margin_left"] = float(tf.MarginLeft)
    d.Close(SaveChanges=False)
    return out


VARIANTS = [
    # (label, p1_ind, p2_ind)
    ("M0_both_plain", {}, {}),
    ("M1_p1plain_p2_left42_hang10.5", {}, {"left": 840, "hanging": 210}),
    ("M2_both_left42_hang10.5", {"left": 840, "hanging": 210}, {"left": 840, "hanging": 210}),
    ("M3_p1plain_p2_hang10.5_only", {}, {"hanging": 210}),
    ("M4_p1plain_p2_left105only", {}, {"left": 105}),
]


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    results = {}
    for label, p1, p2 in VARIANTS:
        path = os.path.join(OUT_DIR, f"{label}.docx")
        write_docx(path, EXT_OVERFLOW_EMU, LINS_2_835, p1, p2)
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False; word.DisplayAlerts = False
        try:
            try:
                m = measure_both_paras(word, path)
                results[label] = m
                print(f"\n[{label}] p1={p1} p2={p2}")
                for p in m["paragraphs"]:
                    print(f"  p{p['i']} {p['text']!r}: x={p['x']} y={p['y']} li={p['li']} fli={p['fli']}",
                          flush=True)
            except Exception as e:
                results[label] = {"error": str(e)}
                print(f"[{label}] ERR: {e}", flush=True)
        finally:
            try: word.Quit()
            except: pass
            time.sleep(1.5)

    with open(RESULT, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2, default=str)
    print(f"\nWrote {RESULT}")


if __name__ == "__main__":
    main()
