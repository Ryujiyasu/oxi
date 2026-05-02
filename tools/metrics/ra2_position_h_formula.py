"""
§17.5 positionH/posOffset formula — horizontal counterpart of §17.3.

Build fixtures with anchored shape having varying `<wp:positionH
relativeFrom="X"><wp:posOffset>N</wp:posOffset>`. Measure Shape.Left
via Word COM. Confirm:

  Shape.Left (pt) = posOffset_emu / 12700 + ref_origin_X(relativeFrom)

ECMA-376 relativeFrom values for positionH:
  page, margin, column, character, leftMargin, rightMargin,
  insideMargin, outsideMargin

COM int (wdRelativeHorizontalPosition):
  0 = Margin, 1 = Page, 2 = Column, 3 = Character,
  4 = LeftMargin, 5 = RightMargin, 6 = InsideMargin, 7 = OutsideMargin

Sweep:
  - 6 relativeFrom values × 4 posOffset values = 24 fixtures
  - posOffset_emu values: 0, 50tw≈25pt, 600tw≈300pt, 1000tw≈500pt
    (in EMU = pt × 12700)
"""
import os
import sys
import time
import json
import zipfile
import shutil

import win32com.client
import pythoncom

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.join(os.path.dirname(__file__), "output")
FIX_DIR = os.path.join(OUT_DIR, "position_h_fixtures")
os.makedirs(FIX_DIR, exist_ok=True)
OUT_JSON = os.path.join(OUT_DIR, "ra2_position_h_formula.json")


# Build via raw OOXML manipulation since python-docx doesn't expose
# wp:positionH relativeFrom directly.
DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="0" distR="0"
                     simplePos="0" relativeHeight="1" behindDoc="0"
                     locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="{rel_from}">
              <wp:posOffset>{offset_emu}</wp:posOffset>
            </wp:positionH>
            <wp:positionV relativeFrom="paragraph">
              <wp:posOffset>0</wp:posOffset>
            </wp:positionV>
            <wp:extent cx="914400" cy="457200"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapNone/>
            <wp:docPr id="1" name="ProbeShape"/>
            <wp:cNvGraphicFramePr/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr/>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    <a:solidFill><a:srgbClr val="DDDDDD"/></a:solidFill>
                  </wps:spPr>
                  <wps:bodyPr wrap="square"/>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="708" w:footer="708" w:gutter="0"/>
      <w:cols w:space="708"/>
    </w:sectPr>
  </w:body>
</w:document>"""


CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""


RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""


def build_fixture(path, *, rel_from, offset_emu):
    """Build a minimal docx with a single anchored rectangle shape."""
    doc_xml = DOCUMENT_XML.format(rel_from=rel_from, offset_emu=offset_emu)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
        z.writestr("_rels/.rels", RELS_XML)
        z.writestr("word/document.xml", doc_xml)


RPC_REJECTED_CODES = {-2147418111, -2147023174, -2147023170}

def retry(fn, *args, retries=15, delay=0.3, **kwargs):
    last_exc = None
    for i in range(retries):
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            last_exc = e
            code = e.args[0] if hasattr(e, "args") and len(e.args) >= 1 else None
            if code in RPC_REJECTED_CODES or "rejected" in str(e).lower():
                pythoncom.PumpWaitingMessages()
                time.sleep(delay * (1.3 ** i))
                continue
            raise
    raise last_exc


def measure(word, path):
    wdoc = retry(lambda: word.Documents.Open(path))
    try:
        wdoc.Repaginate()
        time.sleep(0.05)
        if wdoc.Shapes.Count == 0:
            return {"err": "no shape"}
        s = wdoc.Shapes(1)
        return {
            "shape_left_pt": round(s.Left, 4),
            "shape_top_pt": round(s.Top, 4),
            "rel_h_int": s.RelativeHorizontalPosition,
            "rel_v_int": s.RelativeVerticalPosition,
        }
    finally:
        wdoc.Close(False)


def main():
    rel_from_values = [
        "page", "margin", "column", "character",
        "leftMargin", "rightMargin",
    ]
    # posOffset in EMU; 12700 EMU = 1pt
    offsets_emu = [0, 50 * 1270, 100 * 12700, 200 * 12700, 300 * 12700]
    # = 0pt, 5pt, 100pt, 200pt, 300pt

    # Build fixtures
    fixtures = []
    for rf in rel_from_values:
        for off in offsets_emu:
            name = f"PH_{rf}_off{off}.docx"
            path = os.path.join(FIX_DIR, name)
            try:
                build_fixture(path, rel_from=rf, offset_emu=off)
                fixtures.append((rf, off, path))
            except Exception as e:
                print(f"  ERR build {name}: {e}")
    print(f"Built {len(fixtures)} fixtures.")

    # Measure
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    time.sleep(2.0)
    try:
        retry(lambda: word.Documents.Count)
    except Exception:
        pass

    results = []
    try:
        for rf, off, path in fixtures:
            try:
                m = measure(word, path)
                pred_pt = off / 12700.0
                rec = {
                    "rel_from_xml": rf,
                    "offset_emu": off,
                    "offset_pt": round(pred_pt, 4),
                    "shape_left_pt": m.get("shape_left_pt"),
                    "rel_h_com_int": m.get("rel_h_int"),
                    "ref_origin_x": (m.get("shape_left_pt") - pred_pt) if m.get("shape_left_pt") is not None else None,
                }
                results.append(rec)
                print(f"  rel_from={rf:14s} off={off:>9} ({pred_pt:>5.1f}pt)  shape.Left={rec['shape_left_pt']:>6}  COM_h_int={rec['rel_h_com_int']}  ref_origin_X={rec['ref_origin_x']}")
            except Exception as e:
                print(f"  ERR measure {rf} {off}: {e}")
    finally:
        with open(OUT_JSON, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        print(f"\nSaved {len(results)} records to {OUT_JSON}")
        try:
            word.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
