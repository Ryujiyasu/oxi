"""Identify the +9pt offset property in 1ec1 textbox via 8 variants.

Build on master's proven pattern (repro_1ec1_textbox_ind.py) — reuse CTYPES,
RELS_ROOT, WORD_RELS, STYLES, SETTINGS, write_docx semantics.

Variants (per user's 依頼書):
  Baseline: 1ec1-equivalent (extent overflow + center + lIns=2.835 +
            spacing line=440 exact + snapToGrid=0 + jc=left + sz=28)
  V_A:     no spacing element
  V_B:     no snapToGrid=0 element
  V_C:     spacing line=240 (default)
  V_D:     jc=center
  V_E:     no overflow extent (= 510pt = avail margin)
  V_F:     lIns=0
  V_G:     distL=distR=0
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

OUT_DIR = os.path.abspath("pipeline_data/1ec1_9pt_v2")
os.makedirs(OUT_DIR, exist_ok=True)
RESULT = os.path.abspath("pipeline_data/1ec1_9pt_v2_results.json")

EXT_NO_OVERFLOW_EMU = 6477000  # 510.0pt — fits avail margin (510.2)
DIST_DEFAULT = 114300  # 9pt
DIST_ZERO = 0
LINS_ZERO = 0


def doc_xml_variant(*, extent_emu, lins_emu, dist_lr, jc, snap_xml, spacing_xml):
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
          <wp:anchor distT="0" distB="0" distL="{dist_lr}" distR="{dist_lr}" simplePos="0"
                     relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
            <wp:extent cx="{extent_emu}" cy="600000"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapNone/>
            <wp:docPr id="9" name="ReproShape"/>
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
                          {snap_xml}
                          {spacing_xml}
                          <w:jc w:val="{jc}"/>
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


def write_docx_v(path, **kwargs):
    tmp = tempfile.mkdtemp(prefix="reprov2_")
    try:
        files = [
            ("[Content_Types].xml", CTYPES),
            ("_rels/.rels", RELS_ROOT),
            ("word/_rels/document.xml.rels", WORD_RELS),
            ("word/styles.xml", STYLES),
            ("word/settings.xml", SETTINGS),
            ("word/document.xml", doc_xml_variant(**kwargs)),
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
    for attempt in range(3):
        try:
            doc = word.Documents.Open(docx_path, ReadOnly=True)
            time.sleep(0.4)
            doc.SaveAs2(pdf_path, FileFormat=17)
            doc.Close(SaveChanges=False)
            return True
        except Exception as e:
            last = e
            time.sleep(0.8 + attempt * 0.5)
    print(f"  PDF ERR: {last}")
    return False


def measure_box_x(pdf_path):
    """Find leftmost dark pixel of □ (= second-leftmost text after BodyPara1)."""
    try:
        d = fitz.open(pdf_path)
        page = d[0]
        zoom = 4.0  # higher precision
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        w, h, n = pix.width, pix.height, pix.n
        s = pix.samples
        # Body para "BodyPara1" is at top — page top margin = 56.7pt = 227 px @ 4x
        # Textbox content follows at ~80-150pt = 320-600 px
        # Scan rows in y range 300-700 px
        # Find leftmost dark pixel per row, group by x
        from collections import Counter
        x_counts = Counter()
        rows_data = []
        for py in range(int(60 * zoom), min(int(250 * zoom), h)):
            for px in range(w):
                off = (py * w + px) * n
                r, g, b = s[off], s[off+1], s[off+2]
                if r < 200 and g < 200 and b < 200:
                    x_counts[px] += 1
                    rows_data.append((py, px))
                    break

        # Top 5 most common leftmost x positions (with px tolerance)
        top5 = x_counts.most_common(20)
        # Convert to pt
        top5_pt = [(px / zoom, c) for px, c in top5]

        d.close()
        return {
            "img_w_h": (w, h),
            "n_dark_rows": len(rows_data),
            "leftmost_x_top": [(px, c, round(px / zoom, 3)) for px, c in top5],
        }
    except Exception as e:
        return {"error": str(e)}


VARIANTS = [
    # id, extent, lins, dist, jc, snap_xml, spacing_xml
    ("Baseline", EXT_OVERFLOW_EMU, LINS_2_835, DIST_DEFAULT, "left",
     '<w:snapToGrid w:val="0"/>', '<w:spacing w:line="440" w:lineRule="exact"/>'),
    ("V_A_no_spacing", EXT_OVERFLOW_EMU, LINS_2_835, DIST_DEFAULT, "left",
     '<w:snapToGrid w:val="0"/>', ''),
    ("V_B_no_snapToGrid0", EXT_OVERFLOW_EMU, LINS_2_835, DIST_DEFAULT, "left",
     '', '<w:spacing w:line="440" w:lineRule="exact"/>'),
    ("V_C_line_240_auto", EXT_OVERFLOW_EMU, LINS_2_835, DIST_DEFAULT, "left",
     '<w:snapToGrid w:val="0"/>', '<w:spacing w:line="240" w:lineRule="auto"/>'),
    ("V_D_jc_center", EXT_OVERFLOW_EMU, LINS_2_835, DIST_DEFAULT, "center",
     '<w:snapToGrid w:val="0"/>', '<w:spacing w:line="440" w:lineRule="exact"/>'),
    ("V_E_no_overflow", EXT_NO_OVERFLOW_EMU, LINS_2_835, DIST_DEFAULT, "left",
     '<w:snapToGrid w:val="0"/>', '<w:spacing w:line="440" w:lineRule="exact"/>'),
    ("V_F_lIns_0", EXT_OVERFLOW_EMU, LINS_ZERO, DIST_DEFAULT, "left",
     '<w:snapToGrid w:val="0"/>', '<w:spacing w:line="440" w:lineRule="exact"/>'),
    ("V_G_dist_0", EXT_OVERFLOW_EMU, LINS_2_835, DIST_ZERO, "left",
     '<w:snapToGrid w:val="0"/>', '<w:spacing w:line="440" w:lineRule="exact"/>'),
]


def main():
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    time.sleep(2.0)
    word.Visible = False
    word.DisplayAlerts = False

    results = []
    try:
        for vid, extent, lins, dist, jc, snap_xml, spacing_xml in VARIANTS:
            print(f"\n=== {vid} ===")
            docx = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.docx"))
            pdf = os.path.abspath(os.path.join(OUT_DIR, f"{vid}.pdf"))
            for f in (docx, pdf):
                try: os.remove(f)
                except: pass
            write_docx_v(docx,
                         extent_emu=extent, lins_emu=lins, dist_lr=dist,
                         jc=jc, snap_xml=snap_xml, spacing_xml=spacing_xml)
            print(f"  built {docx}")
            ok = render_pdf(word, docx, pdf)
            if not ok:
                results.append({"id": vid, "error": "PDF render failed"})
                continue
            m = measure_box_x(pdf)
            print(f"  PDF dim: {m.get('img_w_h')}, dark rows: {m.get('n_dark_rows')}")
            print(f"  Top leftmost x (px, count, pt):")
            for px, c, pt in m.get("leftmost_x_top", [])[:5]:
                print(f"    px={px} ({pt}pt) — {c} rows")
            results.append({
                "id": vid,
                "extent_emu": extent, "lins_emu": lins, "dist_emu": dist,
                "jc": jc, "snap_xml": snap_xml, "spacing_xml": spacing_xml,
                **m,
            })
    finally:
        try: word.Quit()
        except: pass

    with open(RESULT, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved -> {RESULT}")


if __name__ == "__main__":
    main()
