"""Compare Oxi vs Word row heights on controlled minimal-repro cells.

Creates a docx with single-cell table (tcMar=0, no borders) and measures
Word row_h via COM, then renders with Oxi (--dump-layout) and computes
the Oxi row_h by inspecting the table border elements.

Purpose: localize the gap between Oxi's computed cell_h and Word's row_h.
"""
import io, json, os, sys, time, subprocess, zipfile
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OXI_RENDERER = Path(r"C:/Users/ryuji/oxi-4/tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe")
TMP_DOCX = Path("pipeline_data") / "_compare_tmp.docx"
TMP_JSON = Path("pipeline_data") / "_compare_tmp.json"

CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'

def para_xml(fs_half, text, font="ＭＳ 明朝"):
    rpr = f'<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/><w:sz w:val="{fs_half}"/><w:szCs w:val="{fs_half}"/></w:rPr>'
    ppr = f'<w:pPr>{rpr}</w:pPr>'
    content = f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>' if text else ''
    return f'<w:p>{ppr}{content}</w:p>'


def build(paras_spec, docgrid_xml):
    paras_xml = [para_xml(hp, txt) for (hp, txt) in paras_spec]
    tbl = (
        '<w:tbl>'
        '<w:tblPr>'
        '<w:tblW w:w="4000" w:type="dxa"/>'
        '<w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>'
        '</w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
        f'{"".join(paras_xml)}'
        '</w:tc></w:tr>'
        '</w:tbl>'
    )
    sentinel = para_xml(21, "END")
    sect = f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992" w:gutter="0"/>{docgrid_xml}</w:sectPr>'
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{tbl}{sentinel}{sect}</w:body></w:document>'
    )
    with zipfile.ZipFile(TMP_DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", xml)


def measure_word(word, paras_spec, docgrid_xml):
    build(paras_spec, docgrid_xml)
    for attempt in range(3):
        try:
            doc = word.Documents.Open(str(TMP_DOCX.resolve()), ReadOnly=True)
            time.sleep(0.3)
            tbl = doc.Tables(1)
            tbl_top_y = tbl.Range.Information(6)
            after_rng = doc.Range(tbl.Range.End, tbl.Range.End)
            after_y = after_rng.Information(6)
            row_h = after_y - tbl_top_y
            doc.Close(False)
            return round(row_h, 2)
        except Exception:
            time.sleep(0.5)
    return None


def measure_oxi(paras_spec, docgrid_xml):
    build(paras_spec, docgrid_xml)
    # Dump layout
    try: os.remove(TMP_JSON)
    except FileNotFoundError: pass
    trash_prefix = str(Path("pipeline_data") / "_compare_trash").replace("\\", "/")
    result = subprocess.run([str(OXI_RENDERER), str(TMP_DOCX.resolve()), trash_prefix, "150", f"--dump-layout={TMP_JSON.resolve()}"],
                   capture_output=True, timeout=30)
    if result.returncode != 0:
        print(f"  renderer err: {result.stderr.decode('utf-8', errors='replace')[:200]}")
    if not TMP_JSON.exists():
        return None
    d = json.load(open(TMP_JSON, encoding='utf-8'))
    p1 = d['pages'][0]['elements']
    # Find horizontal borders (h=0, w>0)
    h_bords = [e for e in p1 if e.get('type') == 'border' and e.get('h', 0) == 0.0 and e.get('w', 0) > 0]
    ys = sorted(set(b['y'] for b in h_bords))
    if len(ys) >= 2:
        return round(ys[1] - ys[0], 2)
    return None


def main():
    word = win32com.client.Dispatch("Word.Application")
    time.sleep(1.0)
    word.Visible = False
    word.DisplayAlerts = False

    grid = ("lm2_350", '<w:docGrid w:type="linesAndChars" w:linePitch="350" w:charSpace="-2714"/>')

    test_cases = [
        ("1 content fs=10.5", [(21, "AB")]),
        ("1 empty fs=10.5", [(21, "")]),
        ("2 content fs=10.5", [(21, "AB"), (21, "CD")]),
        ("3 content fs=10.5", [(21, "AB"), (21, "CD"), (21, "EF")]),
        ("1 content fs=12", [(24, "AB")]),
        ("1 content fs=9", [(18, "AB")]),
    ]

    print(f"{'case':<30} {'word':>8} {'oxi':>8} {'Δ(oxi-word)':>12}")
    print("-" * 65)
    try:
        for label, spec in test_cases:
            wh = measure_word(word, spec, grid[1])
            oh = measure_oxi(spec, grid[1])
            if wh is None or oh is None:
                print(f"{label:<30} ERR")
                continue
            delta = oh - wh
            mark = "*" if abs(delta) > 0.5 else " "
            print(f"{mark} {label:<30} {wh:>8.2f} {oh:>8.2f} {delta:>+8.2f}")
    finally:
        try: word.Quit()
        except Exception: pass
        try: os.remove(TMP_DOCX)
        except Exception: pass
        try: os.remove(TMP_JSON)
        except Exception: pass


if __name__ == "__main__":
    main()
