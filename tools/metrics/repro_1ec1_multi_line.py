"""Multi-line paragraph in textbox: does Word apply ind for line 2+?

ind ignored for first line. But hanging means line 2+ should be at left
indent. If wrapping triggers ind application, that's the missing rule.
"""
import os, sys, time, json, zipfile, shutil, tempfile
import win32com.client
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, "tools/metrics")
from repro_1ec1_textbox_ind import (CTYPES, RELS_ROOT, WORD_RELS, STYLES,
    SETTINGS, EXT_OVERFLOW_EMU, LINS_2_835)

OUT_DIR = os.path.abspath("pipeline_data/r35_1ec1_multiline_docs")


def doc_xml_multi_line(extent_emu, lins_emu, ind_attrs, long_text):
    attr_str = " ".join(f'w:{k}="{v}"' for k, v in ind_attrs.items())
    ind_xml = f'<w:ind {attr_str}/>' if attr_str else ""
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
<wp:extent cx="{extent_emu}" cy="2400000"/>
<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>
<wp:docPr id="9" name="MultiLineRepro"/>
<wp:cNvGraphicFramePr/>
<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
<wps:wsp><wps:cNvSpPr/><wps:spPr>
<a:xfrm><a:off x="0" y="0"/><a:ext cx="{extent_emu}" cy="2400000"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
</wps:spPr>
<wps:txbx><w:txbxContent>
<w:p><w:pPr><w:snapToGrid w:val="0"/>{ind_xml}<w:jc w:val="left"/></w:pPr>
<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>{long_text}</w:t></w:r></w:p>
</w:txbxContent></wps:txbx>
<wps:bodyPr rot="0" wrap="square" lIns="{lins_emu}" tIns="0" rIns="{lins_emu}" bIns="0" anchor="t"/>
</wps:wsp></a:graphicData></a:graphic>
</wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>
</w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="720" w:footer="720" w:gutter="0"/>
<w:cols w:space="720"/><w:docGrid w:type="lines" w:linePitch="360"/>
</w:sectPr></w:body></w:document>'''


def write_docx(path, ext, lins, ind_attrs, long_text):
    tmp = tempfile.mkdtemp(prefix="multiline_")
    try:
        os.makedirs(os.path.join(tmp, "_rels"), exist_ok=True)
        os.makedirs(os.path.join(tmp, "word", "_rels"), exist_ok=True)
        files = [
            ("[Content_Types].xml", CTYPES),
            ("_rels/.rels", RELS_ROOT),
            ("word/_rels/document.xml.rels", WORD_RELS),
            ("word/styles.xml", STYLES),
            ("word/settings.xml", SETTINGS),
            ("word/document.xml", doc_xml_multi_line(ext, lins, ind_attrs, long_text)),
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


def measure(word, path):
    d = word.Documents.Open(path, ReadOnly=True)
    time.sleep(0.5)
    s = d.Shapes(1)
    tf = s.TextFrame
    tr = tf.TextRange
    p = tr.Paragraphs(1)
    pr = p.Range
    chars = pr.Characters
    char_xs = []
    for ci in range(1, min(chars.Count, 40) + 1):
        c = chars(ci)
        try:
            char_xs.append({
                "i": ci, "ch": c.Text,
                "x": float(c.Information(5)),
                "y": float(c.Information(6)),
            })
        except Exception:
            pass
    out = {
        "li": float(pr.ParagraphFormat.LeftIndent),
        "fli": float(pr.ParagraphFormat.FirstLineIndent),
        "char_xs": char_xs,
    }
    d.Close(SaveChanges=False)
    return out


# Long enough to wrap multiple lines. Single ・ + many CJK chars.
LONG = "・" + "あ" * 60  # 61 chars total. At 14pt CJK in textbox content, this should wrap.

VARIANTS = [
    ("L0_long_plain", {}, LONG),
    ("L1_long_left42", {"left": 840}, LONG),
    ("L2_long_left42_hang10.5", {"left": 840, "hanging": 210}, LONG),
    ("L3_long_hang10.5_only", {"hanging": 210}, LONG),
]

results = {}
os.makedirs(OUT_DIR, exist_ok=True)
for label, ind, txt in VARIANTS:
    path = os.path.join(OUT_DIR, f"{label}.docx")
    write_docx(path, EXT_OVERFLOW_EMU, LINS_2_835, ind, txt)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False; word.DisplayAlerts = False
    try:
        try:
            m = measure(word, path)
            results[label] = m
            print(f"\n[{label}] ind={ind}")
            print(f"  li={m['li']} fli={m['fli']}")
            # Show position by line (group by y)
            ys = sorted(set(c['y'] for c in m['char_xs']))
            for yi, y in enumerate(ys[:5]):
                first = next((c for c in m['char_xs'] if c['y'] == y), None)
                if first:
                    print(f"  line{yi+1} y={y:.1f}: first char {first['ch']!r} x={first['x']:.2f}")
        except Exception as e:
            results[label] = {"error": str(e)}
            print(f"[{label}] ERR: {e}")
    finally:
        try: word.Quit()
        except: pass
        time.sleep(1.5)

with open("pipeline_data/r35_1ec1_multiline_2026-05-02.json", "w", encoding="utf-8") as f:
    json.dump(results, f, ensure_ascii=False, indent=2, default=str)
