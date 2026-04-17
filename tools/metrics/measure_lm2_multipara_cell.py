"""
Ra: multi-PARAGRAPH cell row height (b35 class).

b35's cells contain multiple <w:p> (not multi-line via chr(11)). The
earlier v3 sweep measured multi-line-via-break-within-paragraph, which is
a different layout case. This script measures Word's row height for
cells containing N separate paragraphs.

Covers docGrid=LM2 (linePitch=360, the b35 / d77a case) and no-docGrid=LM0
for comparison. Writes to tools/metrics/output/lm2_multipara_cell.json.
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

OUT = Path(__file__).with_name("output") / "lm2_multipara_cell_ext.json"
OUT.parent.mkdir(parents=True, exist_ok=True)
TMP = Path("pipeline_data") / "_lm2_multipara_tmp.docx"
TMP.parent.mkdir(parents=True, exist_ok=True)

CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'

FONTS = [("ＭＳ 明朝", "MS Mincho"), ("ＭＳ ゴシック", "MS Gothic")]
SIZES = [9.0, 10.0, 10.5, 11.0, 12.0, 13.0, 14.0, 16.0, 18.0]
NS = [1, 2, 3]
DOC_GRIDS = [
    ("none", ""),
    ("lm2_360", '<w:docGrid w:type="lines" w:linePitch="360"/>'),
    ("lm2_350", '<w:docGrid w:type="lines" w:linePitch="350"/>'),
    ("linesAndChars_350", '<w:docGrid w:type="linesAndChars" w:linePitch="350" w:charSpace="-2714"/>'),
]


def build(font_xml, sz_half, n_paras, docgrid_xml):
    rpr = f'<w:rPr><w:rFonts w:ascii="{font_xml}" w:eastAsia="{font_xml}" w:hAnsi="{font_xml}"/><w:sz w:val="{sz_half}"/><w:szCs w:val="{sz_half}"/></w:rPr>'
    ppr = f'<w:pPr>{rpr}</w:pPr>'
    one_p = lambda i: f'<w:p>{ppr}<w:r>{rpr}<w:t>P{i}</w:t></w:r></w:p>'
    cell_paras = "".join(one_p(i+1) for i in range(n_paras))
    tbl = (
        '<w:tbl>'
        '<w:tblPr>'
        '<w:tblW w:w="4000" w:type="dxa"/>'
        '<w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>'
        '</w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
        f'{cell_paras}'
        '</w:tc></w:tr>'
        '</w:tbl>'
    )
    sentinel = f'<w:p>{ppr}<w:r>{rpr}<w:t>END</w:t></w:r></w:p>'
    sect = f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992" w:gutter="0"/>{docgrid_xml}</w:sectPr>'
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{tbl}{sentinel}{sect}</w:body></w:document>'
    )
    with zipfile.ZipFile(TMP, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", xml)


def measure(word, font_xml, size, n_paras, grid_label, grid_xml):
    try: os.remove(TMP)
    except FileNotFoundError: pass
    build(font_xml, int(round(size * 2)), n_paras, grid_xml)
    last_err = None
    for attempt in range(4):
        try:
            doc = word.Documents.Open(str(TMP.resolve()), ReadOnly=True)
            time.sleep(0.3)
            tbl = doc.Tables(1)
            tbl_top_y = tbl.Range.Information(6)
            after_rng = doc.Range(tbl.Range.End, tbl.Range.End)
            after_y = after_rng.Information(6)
            row_h = after_y - tbl_top_y
            # Per-paragraph Y via sel
            sel = word.Selection
            para_ys = []
            for i in range(1, n_paras + 1):
                # navigate: paragraph i of the table's first cell = doc.Paragraphs[X]
                # Paragraphs in doc order: cell paras (1..n_paras), then sentinel
                p_y = doc.Paragraphs(i).Range.Information(6)
                para_ys.append(p_y)
            doc.Close(False)
            return {
                "font": font_xml, "size": size, "n_paras": n_paras,
                "grid": grid_label,
                "tbl_top_y": round(tbl_top_y, 2),
                "after_y": round(after_y, 2),
                "row_h": round(row_h, 2),
                "para_ys": [round(y, 2) for y in para_ys],
            }
        except Exception as e:
            last_err = e
            time.sleep(0.8 + attempt * 0.5)
    return {"font": font_xml, "size": size, "n_paras": n_paras, "grid": grid_label, "error": str(last_err)}


def main():
    word = win32com.client.Dispatch("Word.Application")
    time.sleep(1.0)
    word.Visible = False
    word.DisplayAlerts = False
    results = []
    total = len(FONTS) * len(SIZES) * len(NS) * len(DOC_GRIDS)
    i = 0
    try:
        for font_xml, pretty in FONTS:
            for size in SIZES:
                for n in NS:
                    for gl, gx in DOC_GRIDS:
                        i += 1
                        m = measure(word, font_xml, size, n, gl, gx)
                        m["font_pretty"] = pretty
                        results.append(m)
                        if "error" in m:
                            print(f"[{i:2d}/{total}] {pretty} {size}pt n={n} grid={gl}: ERR {m['error']}")
                        else:
                            print(f"[{i:2d}/{total}] {pretty} {size}pt n={n} grid={gl}: row_h={m['row_h']} para_ys={m['para_ys']}")
    finally:
        try: word.Quit()
        except Exception: pass
        try: os.remove(TMP)
        except Exception: pass
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved -> {OUT}")

    # Derive formula per (font, size, grid)
    print("\n=== row_h per (font, size, grid) ===")
    groups = {}
    for r in results:
        if "error" in r: continue
        k = (r["font_pretty"], r["size"], r["grid"])
        groups.setdefault(k, {})[r["n_paras"]] = r["row_h"]
    for k, rows in sorted(groups.items()):
        font, size, grid = k
        pieces = " ".join(f"n{n}={rows[n]}" for n in sorted(rows))
        if 1 in rows and 2 in rows:
            gap12 = rows[2] - rows[1]
            pieces += f"  | gap12={gap12}"
        if 2 in rows and 3 in rows:
            gap23 = rows[3] - rows[2]
            pieces += f" gap23={gap23}"
        print(f"  {font:>10} {size:>5.1f}pt grid={grid:<8}: {pieces}")


if __name__ == "__main__":
    main()
