"""
Ra: LM0 multi-line table cell — v5 robust isolated Word.

v4 hit RPC rejection on first call + locked temp file cascade. v5 fixes:
  - DispatchEx for a fresh Word instance (won't share with master's Word)
  - Per-iteration unique temp filename (no shared lock)
  - Tear-down cleanup of leftover *_tmp_*.docx files

Coverage matches v4: Calibri/Yu Mincho/Meiryo/HGS Mincho E/TNR × 7 sizes × 4 n.
"""
import io
import json
import os
import sys
import time
import zipfile
import uuid
from pathlib import Path
import pythoncom
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(__file__).with_name("output") / "lm0_multiline_cell_v5.json"
OUT.parent.mkdir(parents=True, exist_ok=True)
TMP_DIR = Path("pipeline_data") / "_lm0_cell_v5_tmp"
TMP_DIR.mkdir(parents=True, exist_ok=True)

CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'

SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992" w:gutter="0"/><w:cols w:space="425"/></w:sectPr>'

FONTS = [
    ("Calibri", "Calibri"),
    ("Yu Mincho", "Yu Mincho"),
    ("Meiryo", "Meiryo"),
    ("HGS明朝E", "HGS Mincho E"),
    ("Times New Roman", "TNR"),
]
SIZES = [9.0, 10.0, 10.5, 11.0, 12.0, 14.0, 18.0]
NS = [1, 2, 3, 4]


def run_xml(font_xml, sz_half, text):
    return (
        f'<w:r><w:rPr>'
        f'<w:rFonts w:ascii="{font_xml}" w:eastAsia="{font_xml}" w:hAnsi="{font_xml}"/>'
        f'<w:sz w:val="{sz_half}"/><w:szCs w:val="{sz_half}"/>'
        f'</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'
    )


def ppr_xml(font_xml, sz_half):
    return (
        f'<w:pPr><w:rPr>'
        f'<w:rFonts w:ascii="{font_xml}" w:eastAsia="{font_xml}" w:hAnsi="{font_xml}"/>'
        f'<w:sz w:val="{sz_half}"/><w:szCs w:val="{sz_half}"/>'
        f'</w:rPr></w:pPr>'
    )


def cell_paragraph(font_xml, sz_half, n_lines):
    ppr = ppr_xml(font_xml, sz_half)
    pieces = []
    is_cjk = "明朝" in font_xml or "Mincho" in font_xml or "Meiryo" in font_xml
    text = "あ" if is_cjk else "Ag"
    for i in range(n_lines):
        pieces.append(run_xml(font_xml, sz_half, f"L{i+1}{text}"))
        if i < n_lines - 1:
            pieces.append('<w:r><w:br/></w:r>')
    return f'<w:p>{ppr}{"".join(pieces)}</w:p>'


def table_xml(font_xml, sz_half, n_lines):
    cell_p = cell_paragraph(font_xml, sz_half, n_lines)
    tbl = (
        '<w:tbl>'
        '<w:tblPr>'
        '<w:tblW w:w="4000" w:type="dxa"/>'
        '<w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>'
        '</w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
        '<w:tr>'
        '<w:tc>'
        '<w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
        f'{cell_p}'
        '</w:tc>'
        '</w:tr>'
        '</w:tbl>'
    )
    sentinel = f'<w:p>{ppr_xml(font_xml, sz_half)}{run_xml(font_xml, sz_half, "END")}</w:p>'
    return tbl + sentinel


def body_xml(font_xml, sz_half):
    ppr = ppr_xml(font_xml, sz_half)
    is_cjk = "明朝" in font_xml or "Mincho" in font_xml or "Meiryo" in font_xml
    text = "あ" if is_cjk else "Ag"
    p = lambda i: f'<w:p>{ppr}{run_xml(font_xml, sz_half, f"P{i}{text}")}</w:p>'
    return p(1) + p(2) + p(3)


def write_docx(path, inner):
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{inner}{SECT}</w:body></w:document>'
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", xml)


def open_and_measure(word, inner, measurer, idx):
    path = TMP_DIR / f"v5_{idx:04d}_{uuid.uuid4().hex[:8]}.docx"
    write_docx(path, inner)
    last_err = None
    for attempt in range(3):
        try:
            doc = word.Documents.Open(str(path.resolve()), ReadOnly=True)
            time.sleep(0.2)
            out = measurer(doc)
            doc.Close(SaveChanges=False)
            try:
                path.unlink()
            except Exception:
                pass
            return out
        except Exception as e:
            last_err = e
            time.sleep(0.5 + attempt * 0.5)
    try:
        path.unlink()
    except Exception:
        pass
    raise last_err


def measure_cell(doc):
    tbl = doc.Tables(1)
    tbl_top_y = tbl.Range.Information(6)
    after_rng = doc.Range(tbl.Range.End, tbl.Range.End)
    after_y = after_rng.Information(6)
    row_h = after_y - tbl_top_y
    line_ys = []
    cr = tbl.Cell(1, 1).Range
    for ci in range(cr.Start, cr.End):
        r = doc.Range(ci, ci + 1)
        y = r.Information(6)
        if not line_ys or abs(y - line_ys[-1]) > 0.5:
            line_ys.append(y)
    return {
        "tbl_top_y": round(tbl_top_y, 3),
        "after_y": round(after_y, 3),
        "row_h": round(row_h, 3),
        "line_ys": [round(y, 3) for y in line_ys],
    }


def measure_body(doc):
    y1 = doc.Paragraphs(1).Range.Information(6)
    y2 = doc.Paragraphs(2).Range.Information(6)
    y3 = doc.Paragraphs(3).Range.Information(6)
    return {"y1": round(y1, 3), "y2": round(y2, 3), "y3": round(y3, 3),
            "gap12": round(y2 - y1, 3), "gap23": round(y3 - y2, 3)}


def main():
    pythoncom.CoInitialize()
    # DispatchEx forces a NEW Word instance — doesn't attach to master's Word
    word = win32com.client.DispatchEx("Word.Application")
    time.sleep(2.0)
    word.Visible = False
    word.DisplayAlerts = False

    results = {"cells": [], "body": []}
    idx = 0
    try:
        total = len(FONTS) * len(SIZES) * len(NS)
        i = 0
        for font_xml, pretty in FONTS:
            for size in SIZES:
                sz_half = int(round(size * 2))
                for n in NS:
                    i += 1
                    idx += 1
                    try:
                        m = open_and_measure(word, table_xml(font_xml, sz_half, n), measure_cell, idx)
                        entry = {"font": pretty, "size": size, "n_lines": n, **m}
                        results["cells"].append(entry)
                        print(f"[{i:3d}/{total}] cell {pretty:<14} {size:>5.1f}pt n={n}: row_h={m['row_h']:>5.1f}")
                    except Exception as e:
                        results["cells"].append({"font": pretty, "size": size, "n_lines": n, "error": str(e)})
                        print(f"[{i:3d}/{total}] cell {pretty:<14} {size:>5.1f}pt n={n}: ERR {e}")
        for font_xml, pretty in FONTS:
            for size in SIZES:
                sz_half = int(round(size * 2))
                idx += 1
                try:
                    m = open_and_measure(word, body_xml(font_xml, sz_half), measure_body, idx)
                    entry = {"font": pretty, "size": size, **m}
                    results["body"].append(entry)
                    print(f"body {pretty:<14} {size:>5.1f}pt: gap12={m['gap12']:>5.1f} gap23={m['gap23']:>5.1f}")
                except Exception as e:
                    results["body"].append({"font": pretty, "size": size, "error": str(e)})
                    print(f"body {pretty:<14} {size:>5.1f}pt: ERR {e}")
    finally:
        try:
            word.Quit()
        except Exception:
            pass
        # Cleanup any leftover tmp files
        for f in TMP_DIR.glob("*.docx"):
            try:
                f.unlink()
            except Exception:
                pass

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved -> {OUT}")


if __name__ == "__main__":
    main()
