"""V_H_roundRect 仮説検証: prst="rect" → prst="roundRect" の単一変更で □ が
43pt → 48pt にシフトするか pin。

Direct comparison:
  V_RECT:      prst="rect"      (Baseline reproduces partial sweep result)
  V_ROUNDRECT: prst="roundRect" (1ec1's actual setup)

Other 8 settings identical (overflow extent + center + lIns=2.835 + spacing
line=440 exact + snapToGrid=0 + jc=left + sz=28 + dist=114300).

If V_RECT → V_ROUNDRECT shows +5.5pt shift, +9pt offset = roundRect content
padding (corner-radius effect).
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

sys.path.insert(0, "tools/metrics")
from repro_1ec1_textbox_ind import CTYPES, RELS_ROOT, WORD_RELS, STYLES, SETTINGS, EXT_OVERFLOW_EMU, LINS_2_835

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath("pipeline_data/1ec1_roundRect_test")
os.makedirs(OUT_DIR, exist_ok=True)
RESULT = os.path.abspath("pipeline_data/1ec1_roundRect_test_results.json")


def doc_xml_prst(prst_val):
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
            <wp:extent cx="{EXT_OVERFLOW_EMU}" cy="600000"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapNone/>
            <wp:docPr id="9" name="ReproShape"/>
            <wp:cNvGraphicFramePr/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr/>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="{EXT_OVERFLOW_EMU}" cy="600000"/></a:xfrm>
                    <a:prstGeom prst="{prst_val}"><a:avLst/></a:prstGeom>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p>
                        <w:pPr>
                          <w:snapToGrid w:val="0"/>
                          <w:spacing w:line="440" w:lineRule="exact"/>
                          <w:jc w:val="left"/>
                        </w:pPr>
                        <w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>□3</w:t></w:r>
                      </w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" wrap="square" lIns="{LINS_2_835}" tIns="0" rIns="{LINS_2_835}" bIns="0" anchor="t"/>
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


def write_docx(path, prst_val):
    tmp = tempfile.mkdtemp(prefix="rrtest_")
    try:
        files = [
            ("[Content_Types].xml", CTYPES),
            ("_rels/.rels", RELS_ROOT),
            ("word/_rels/document.xml.rels", WORD_RELS),
            ("word/styles.xml", STYLES),
            ("word/settings.xml", SETTINGS),
            ("word/document.xml", doc_xml_prst(prst_val)),
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
    """Find leftmost dark pixel of □ glyph (the second-leftmost cluster after textbox border)."""
    try:
        d = fitz.open(pdf_path)
        page = d[0]
        zoom = 4.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        w, h, n = pix.width, pix.height, pix.n
        s = pix.samples
        from collections import Counter
        x_counts = Counter()
        # Scan textbox y range (approx 60-250pt)
        for py in range(int(60 * zoom), min(int(250 * zoom), h)):
            for px in range(w):
                off = (py * w + px) * n
                r, g, b = s[off], s[off+1], s[off+2]
                if r < 200 and g < 200 and b < 200:
                    x_counts[px] += 1
                    break
        d.close()
        # Find clusters of leftmost x positions (textbox border + glyph)
        # Group by tolerance 4px
        sorted_x = sorted(x_counts.items())
        clusters = []
        for x, cnt in sorted_x:
            if clusters and x - clusters[-1]["max_x"] <= 4:
                clusters[-1]["max_x"] = x
                clusters[-1]["count"] += cnt
            else:
                clusters.append({"min_x": x, "max_x": x, "count": cnt})
        # Sort by count descending
        clusters.sort(key=lambda c: -c["count"])
        return {
            "img_w_h": (w, h),
            "zoom": zoom,
            "clusters": [(c["min_x"], c["max_x"], c["count"], c["min_x"] / zoom) for c in clusters[:5]],
        }
    except Exception as e:
        return {"error": str(e)}


def main():
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    time.sleep(2.0)
    word.Visible = False
    word.DisplayAlerts = False
    results = []
    try:
        for vid, prst_val in [("V_RECT", "rect"), ("V_ROUNDRECT", "roundRect")]:
            print(f"\n=== {vid} (prst={prst_val}) ===")
            docx = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.docx"))
            pdf = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.pdf"))
            for f in (docx, pdf):
                try: os.remove(f)
                except: pass
            write_docx(docx, prst_val)
            ok = render_pdf(word, docx, pdf)
            if not ok:
                results.append({"id": vid, "prst": prst_val, "error": "PDF render failed"})
                continue
            m = measure_box_x(pdf)
            print(f"  Img: {m.get('img_w_h')}")
            print(f"  Top clusters (min_x, max_x, count, min_x_pt):")
            for cluster in m.get("clusters", []):
                print(f"    {cluster}")
            results.append({"id": vid, "prst": prst_val, **m})
    finally:
        try: word.Quit()
        except: pass
    with open(RESULT, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved -> {RESULT}")


if __name__ == "__main__":
    main()
