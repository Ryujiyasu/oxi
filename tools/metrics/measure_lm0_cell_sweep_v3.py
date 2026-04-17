"""
Ra: LM0 multi-line table cell height — v3 (raw-XML methodology).

v2 used Word COM Documents.Add() + ParagraphFormat.LineSpacingRule=0,
which may inject an explicit <w:spacing w:line="240"/> element and produce
line heights that don't match lineRule=auto default. v3 builds the docx
directly via raw XML with NO <w:spacing> element — same methodology as
the sibling worktree's pilot_lm0_lineauto.py that found 13.5pt for
MS Mincho 10.5pt body.

Output: tools/metrics/output/lm0_multiline_cell_v3.json
"""
import io
import json
import os
import sys
import time
import zipfile
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT = Path(__file__).with_name("output") / "lm0_multiline_cell_v3.json"
OUT.parent.mkdir(parents=True, exist_ok=True)
TMP = Path("pipeline_data") / "_lm0_cell_v3_tmp.docx"
TMP.parent.mkdir(parents=True, exist_ok=True)

CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'

SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992" w:gutter="0"/><w:cols w:space="425"/></w:sectPr>'

FONTS = [
    ("ＭＳ 明朝", "MS Mincho"),
    ("ＭＳ ゴシック", "MS Gothic"),
]
SIZES = [9.0, 10.0, 10.5, 11.0, 12.0, 13.0, 14.0, 16.0, 18.0]
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
    """A single paragraph with n_lines via <w:br/> soft line breaks."""
    ppr = ppr_xml(font_xml, sz_half)
    # Runs: L1 <br/> L2 <br/> ... Ln, all in a single <w:p>
    pieces = []
    for i in range(n_lines):
        pieces.append(run_xml(font_xml, sz_half, f"L{i+1}"))
        if i < n_lines - 1:
            pieces.append('<w:r><w:br/></w:r>')
    return f'<w:p>{ppr}{"".join(pieces)}</w:p>'


def table_xml(font_xml, sz_half, n_lines):
    """Single-row, single-cell table; zero cell margins."""
    cell_p = cell_paragraph(font_xml, sz_half, n_lines)
    # tblW 4000 twips = 200pt; tcW 4000 twips; zero cell margins via tcMar
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
    # Sentinel paragraph after the table for after_y measurement
    sentinel = f'<w:p>{ppr_xml(font_xml, sz_half)}{run_xml(font_xml, sz_half, "END")}</w:p>'
    return tbl + sentinel


def body_xml(font_xml, sz_half):
    """Three plain paragraphs for gap measurement."""
    ppr = ppr_xml(font_xml, sz_half)
    p = lambda i: f'<w:p>{ppr}{run_xml(font_xml, sz_half, f"P{i}")}</w:p>'
    return p(1) + p(2) + p(3)


def write_docx(inner):
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{inner}{SECT}</w:body></w:document>'
    )
    with zipfile.ZipFile(TMP, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", xml)


def open_and_measure(word, inner, measurer):
    write_docx(inner)
    last_err = None
    for attempt in range(5):
        try:
            doc = word.Documents.Open(str(TMP.resolve()), ReadOnly=True)
            time.sleep(0.3)
            out = measurer(doc)
            doc.Close(SaveChanges=False)
            return out
        except Exception as e:
            last_err = e
            time.sleep(0.8 + attempt * 0.5)
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
    word = win32com.client.Dispatch("Word.Application")
    time.sleep(1.0)
    word.Visible = False
    word.DisplayAlerts = False
    results = {"cells": [], "body": []}
    try:
        total = len(FONTS) * len(SIZES) * len(NS)
        i = 0
        for font_xml, pretty in FONTS:
            for size in SIZES:
                sz_half = int(round(size * 2))
                for n in NS:
                    i += 1
                    try:
                        m = open_and_measure(word, table_xml(font_xml, sz_half, n), measure_cell)
                        entry = {"font": pretty, "size": size, "n_lines": n, **m}
                        results["cells"].append(entry)
                        print(f"[{i:3d}/{total}] cell {pretty} {size}pt n={n}: row_h={m['row_h']}")
                    except Exception as e:
                        results["cells"].append({"font": pretty, "size": size, "n_lines": n, "error": str(e)})
                        print(f"[{i:3d}/{total}] cell {pretty} {size}pt n={n}: ERR {e}")
        for font_xml, pretty in FONTS:
            for size in SIZES:
                sz_half = int(round(size * 2))
                try:
                    m = open_and_measure(word, body_xml(font_xml, sz_half), measure_body)
                    entry = {"font": pretty, "size": size, **m}
                    results["body"].append(entry)
                    print(f"body {pretty} {size}pt: gap12={m['gap12']} gap23={m['gap23']}")
                except Exception as e:
                    results["body"].append({"font": pretty, "size": size, "error": str(e)})
                    print(f"body {pretty} {size}pt: ERR {e}")
    finally:
        try:
            word.Quit()
        except Exception:
            pass
        try:
            os.remove(TMP)
        except Exception:
            pass

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved -> {OUT}")

    # Derive last_alloc and gap
    print("\n=== last_alloc, gap per (font, size) ===")
    by = {}
    for c in results["cells"]:
        if c.get("error"): continue
        by.setdefault((c["font"], c["size"]), {})[c["n_lines"]] = c["row_h"]
    for k, rows in sorted(by.items()):
        if 1 in rows and 2 in rows:
            la = rows[1]; gap = rows[2] - rows[1]
            font, size = k
            predicted_34 = {n: la + (n - 1) * gap for n in (3, 4) if n in rows}
            ok = all(abs(predicted_34[n] - rows[n]) < 0.5 for n in predicted_34)
            print(f"  {font:>10} {size:>5.1f}pt: last_alloc={la:6.2f} gap={gap:6.2f}  ok={ok}  actual_3={rows.get(3)} actual_4={rows.get(4)}")


if __name__ == "__main__":
    main()
