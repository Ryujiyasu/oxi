"""V_I_compatLnSpc test: Add compatLnSpc=1 + horzOverflow=overflow to bodyPr.

1ec1's Shape 9 bodyPr has:
  vertOverflow="overflow" horzOverflow="overflow" compatLnSpc="1"

My earlier synthetic Baseline missed these. Test if adding them shifts □
toward 1ec1's 48pt position.
"""
import os
import sys
import time
import json
import zipfile
import shutil
import tempfile
import pythoncom
import win32com.client
import fitz
from collections import Counter

sys.path.insert(0, "tools/metrics")
from repro_1ec1_textbox_ind import CTYPES, RELS_ROOT, WORD_RELS, STYLES, SETTINGS, EXT_OVERFLOW_EMU, LINS_2_835

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath("pipeline_data/1ec1_compatLnSpc_test")
os.makedirs(OUT_DIR, exist_ok=True)


def doc_xml(*, prst, adj, body_attrs, line):
    spacing_xml = f'<w:spacing w:line="{line}" w:lineRule="exact"/>' if line else ''
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
            <wp:extent cx="{EXT_OVERFLOW_EMU}" cy="3028950"/>
            <wp:effectExtent l="0" t="0" r="28575" b="19050"/>
            <wp:wrapNone/>
            <wp:docPr id="9" name="ReproShape"/>
            <wp:cNvGraphicFramePr/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr/>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="{EXT_OVERFLOW_EMU}" cy="3028950"/></a:xfrm>
                    <a:prstGeom prst="{prst}">{adj_xml}</a:prstGeom>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p>
                        <w:pPr>
                          <w:snapToGrid w:val="0"/>
                          {spacing_xml}
                          <w:jc w:val="left"/>
                        </w:pPr>
                        <w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>□3</w:t></w:r>
                      </w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr {body_attrs}/>
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


def write_docx(path, **kwargs):
    tmp = tempfile.mkdtemp(prefix="cls_")
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
            time.sleep(0.5)
            doc.SaveAs2(pdf_path, FileFormat=17)
            doc.Close(SaveChanges=False)
            return True
        except Exception as e:
            last = e
            time.sleep(1.0 + attempt * 0.5)
    print(f"  PDF ERR: {last}")
    return False


def measure_box_x(pdf_path):
    try:
        d = fitz.open(pdf_path)
        page = d[0]
        # Use search_for □ to get exact position
        instances = page.search_for("□")
        zoom = 4.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        w, h, n = pix.width, pix.height, pix.n
        s = pix.samples
        results = []
        for inst in instances:
            top_px = int(inst.y0 * zoom)
            bottom_px = int(inst.y1 * zoom)
            left_search = max(0, int((inst.x0 - 5) * zoom))
            right_search = min(w, int((inst.x1 + 2) * zoom))
            leftmost = None
            for py in range(max(0, top_px), min(h, bottom_px)):
                for px in range(left_search, right_search):
                    off = (py * w + px) * n
                    r, g, b = s[off], s[off+1], s[off+2]
                    if r < 200 and g < 200 and b < 200:
                        if leftmost is None or px < leftmost:
                            leftmost = px
                        break
            if leftmost is not None:
                results.append({
                    "search_left_pt": inst.x0,
                    "leftmost_px": leftmost,
                    "leftmost_pt": leftmost / zoom,
                })
        d.close()
        return results
    except Exception as e:
        return [{"error": str(e)}]


# bodyPr attribute strings
BODYPR_MINIMAL = f'rot="0" wrap="square" lIns="{LINS_2_835}" tIns="0" rIns="{LINS_2_835}" bIns="0" anchor="t"'
BODYPR_1EC1 = f'rot="0" spcFirstLastPara="0" vertOverflow="overflow" horzOverflow="overflow" vert="horz" wrap="square" lIns="{LINS_2_835}" tIns="0" rIns="{LINS_2_835}" bIns="0" numCol="1" spcCol="0" rtlCol="0" fromWordArt="0" anchor="t" anchorCtr="0" forceAA="0" compatLnSpc="1"'

VARIANTS = [
    ("V_I0_baseline_minimal", "rect", None, BODYPR_MINIMAL, 440),
    ("V_I1_roundRect_adj4015_minimal", "roundRect", 4015, BODYPR_MINIMAL, 440),
    ("V_I2_roundRect_adj4015_full_bodyPr", "roundRect", 4015, BODYPR_1EC1, 440),
    ("V_I3_compatLnSpc_only", "rect", None, f'{BODYPR_MINIMAL} compatLnSpc="1"', 440),
    ("V_I4_overflow_only", "rect", None, f'rot="0" vertOverflow="overflow" horzOverflow="overflow" wrap="square" lIns="{LINS_2_835}" tIns="0" rIns="{LINS_2_835}" bIns="0" anchor="t"', 440),
]


def main():
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    time.sleep(2.0)
    word.Visible = False
    word.DisplayAlerts = False
    results = []
    try:
        for vid, prst, adj, bp, line in VARIANTS:
            print(f"\n=== {vid} ===")
            docx = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.docx"))
            pdf = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.pdf"))
            for f in (docx, pdf):
                try: os.remove(f)
                except: pass
            write_docx(docx, prst=prst, adj=adj, body_attrs=bp, line=line)
            ok = render_pdf(word, docx, pdf)
            if not ok:
                results.append({"id": vid, "error": "render failed"})
                continue
            measurements = measure_box_x(pdf)
            for i, m in enumerate(measurements):
                if "error" not in m:
                    print(f"  □#{i+1}: search_L={m['search_left_pt']:.2f}pt → visible at {m['leftmost_pt']:.2f}pt")
                    break
            results.append({"id": vid, "prst": prst, "adj": adj, "measurements": measurements})
    finally:
        try: word.Quit()
        except: pass
    out = os.path.join(OUT_DIR, "results.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved -> {out}")


if __name__ == "__main__":
    main()
