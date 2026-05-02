"""§13.1 round 29 — cellMar variation × Mech 2 cap.

Round 28 confirmed Mech 2 inside cell = body when cellMar=4.95pt
(default). Round 29 validates that §13.1's effective_width =
tcW - 2*cellMar formula holds for non-default cellMar values.

Test: cellMar ∈ {0pt, 4.95pt default, 10pt, asymmetric L=2,R=10pt}.
For each cellMar, set tcW such that inner content = exactly 282pt
(i.e., right at Mech 2 cap boundary at fs=12, cap=6pt, slack=6pt).

If §13.1 formula correct: all cellMar variants fit 24 chars with
mech2_comp=6pt (body equivalence). If wrong: behavior diverges.
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\mech2_cellmar_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\mech2_cellmar.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"


def make_cell_doc_xml(probe, sz_val, page_w_tw, margin_tw, cell_w_tw,
                      cellmar_l_tw, cellmar_r_tw):
    """1-row, 1-cell table with custom cellMar."""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body>'
            '<w:tbl>'
            '<w:tblPr>'
            f'<w:tblW w:w="{cell_w_tw}" w:type="dxa"/>'
            f'<w:tblCellMar>'
            f'<w:top w:w="0" w:type="dxa"/>'
            f'<w:left w:w="{cellmar_l_tw}" w:type="dxa"/>'
            f'<w:bottom w:w="0" w:type="dxa"/>'
            f'<w:right w:w="{cellmar_r_tw}" w:type="dxa"/>'
            f'</w:tblCellMar>'
            '</w:tblPr>'
            '<w:tblGrid>'
            f'<w:gridCol w:w="{cell_w_tw}"/>'
            '</w:tblGrid>'
            '<w:tr>'
            '<w:tc>'
            f'<w:tcPr><w:tcW w:w="{cell_w_tw}" w:type="dxa"/></w:tcPr>'
            '<w:p>'
            '<w:pPr><w:jc w:val="both"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/></w:rPr>'
            f'<w:t>{probe}</w:t></w:r>'
            '</w:p>'
            '</w:tc>'
            '</w:tr>'
            '</w:tbl>'
            f'<w:sectPr><w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1134" w:right="{margin_tw}" w:bottom="1134" w:left="{margin_tw}"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def make_styles_xml(sz_val):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults>'
            '<w:rPrDefault><w:rPr>'
            f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            '<w:kern w:val="2"/>'
            f'<w:sz w:val="{sz_val}"/>'
            '</w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr/></w:pPrDefault>'
            '</w:docDefaults></w:styles>')


def make_settings_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:characterSpacingControl w:val="compressPunctuation"/>'
            '</w:settings>')


def make_docx(label, doc_xml, sz_val):
    out_path = os.path.join(OUT_DIR, f"{label}.docx")
    styles_xml = make_styles_xml(sz_val)
    settings_xml = make_settings_xml()
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/word/settings.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
        '</Types>'
    )
    pkg_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
        ' Target="word/document.xml"/>'
        '</Relationships>'
    )
    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"'
        ' Target="styles.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"'
        ' Target="settings.xml"/>'
        '</Relationships>'
    )
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", pkg_rels)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/document.xml", doc_xml)
        z.writestr("word/styles.xml", styles_xml)
        z.writestr("word/settings.xml", settings_xml)
    return out_path


def kill_word():
    subprocess.run(['taskkill','/F','/IM','WINWORD.EXE'], capture_output=True)
    time.sleep(2)


def measure_one(path):
    word = w32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        d = word.Documents.Open(str(path), ReadOnly=True)
        time.sleep(0.2)
        try:
            chars = d.Range().Characters
            xs = []
            for ci in range(1, chars.Count + 1):
                try:
                    c = chars(ci)
                    t = c.Text
                    if not t or any(ord(ch) < 32 for ch in t): continue
                    xs.append((t, float(c.Information(5)), float(c.Information(6))))
                except: continue
        finally:
            try: d.Close(SaveChanges=False)
            except: pass
        if not xs: return {"error": "no chars"}
        y0 = xs[0][2]
        line1 = sorted([(t, x) for t, x, y in xs if abs(y - y0) < 0.5],
                       key=lambda v: v[1])
        n_line1 = len(line1)
        char_advances = []
        for i in range(len(line1) - 1):
            t = line1[i][0]
            adv = round(line1[i+1][1] - line1[i][1], 3)
            char_advances.append({"ch": t, "adv": adv})
        return {"n_chars_line1": n_line1, "char_advances": char_advances,
                "first_x": line1[0][1] if line1 else None,
                "last_x": line1[-1][1] if line1 else None}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    probe = "「漢」漢" * 6  # 24 chars, natural=288pt
    fs_pt = 12.0
    sz_val = 24
    margin_tw = 170 * 10

    target_inner_pt = 282.0  # = nat - 6pt (cap exactly)

    # cellMar variants: (label, left_pt, right_pt)
    cellmar_variants = [
        ("zero",        0.0,    0.0),
        ("default",     4.95,   4.95),
        ("ten",         10.0,   10.0),
        ("asym_2_10",   2.0,    10.0),
    ]

    # Also test 2 slacks for each: slack=6 (cap exact), slack=8 (wrap)
    slacks = [3.0, 6.0, 8.0]

    for slack in slacks:
        target_inner_pt = 288.0 - slack
        print(f"\n=== slack={slack} target_inner={target_inner_pt}pt ===")
        for v_label, l_pt, r_pt in cellmar_variants:
            cellmar_total = l_pt + r_pt
            cell_w_pt = target_inner_pt + cellmar_total
            cell_w_tw = int(round(cell_w_pt * 20))
            cellmar_l_tw = int(round(l_pt * 20))
            cellmar_r_tw = int(round(r_pt * 20))
            page_w_tw = int((cell_w_pt + 170) * 20)

            doc_xml = make_cell_doc_xml(probe, sz_val, page_w_tw, margin_tw,
                                         cell_w_tw, cellmar_l_tw, cellmar_r_tw)
            label = f"sl{slack}_{v_label}"
            try:
                p = make_docx(label, doc_xml, sz_val)
            except Exception as e:
                out[label] = {"build_error": str(e)}
                continue
            kill_word()
            try:
                r = measure_one(p)
            except Exception as e:
                r = {"measure_error": str(e)}
                kill_word()
            entry = {
                "slack": slack, "variant": v_label,
                "cellmar_l_pt": l_pt, "cellmar_r_pt": r_pt,
                "cell_w_pt": round(cell_w_pt, 2),
                "expected_inner_pt": round(target_inner_pt, 2),
                **r,
            }
            advs = entry.pop("char_advances", None)
            if advs:
                sum_actual = sum(c["adv"] for c in advs)
                n = len(advs) + 1
                sum_natural = n * fs_pt - fs_pt  # n-1 advances * fs
                # actually n_advances * fs (but last char advance unknown)
                entry["n_chars"] = entry.get("n_chars_line1")
                entry["mech2_comp"] = round(len(advs) * fs_pt - sum_actual, 2)
            out[label] = entry
            n_v = entry.get("n_chars_line1", "?")
            m2 = entry.get("mech2_comp", "?")
            fx = entry.get("first_x", "?")
            lx = entry.get("last_x", "?")
            print(f"  cellmar L={l_pt}/R={r_pt} cell_w={cell_w_pt:.1f} | "
                  f"n={n_v} m2={m2} first_x={fx} last_x={lx}")
            with open(RESULT, "w", encoding="utf-8") as f:
                json.dump(out, f, ensure_ascii=False, indent=2)

    print("\n========== SUMMARY ==========")
    print(f"  expected: §13.1 says effective_width = tcW - 2×cellMar")
    print(f"  if formula correct, m2 should equal slack for slack ≤ 6, n=24")
    print(f"  if formula correct, line wraps to n=23 for slack > 6")
    print()
    print(f"{'slack':>6} {'variant':>12} {'L/R':>10} {'cell_w':>8} {'n':>4} {'m2':>6}")
    for slack in slacks:
        for v_label, l_pt, r_pt in cellmar_variants:
            label = f"sl{slack}_{v_label}"
            info = out.get(label, {})
            n = info.get("n_chars_line1", "?")
            m2 = info.get("mech2_comp", "?")
            cw = info.get("cell_w_pt", "?")
            print(f"  {slack:>4} {v_label:>12} {f'{l_pt}/{r_pt}':>10} "
                  f"{cw:>8} {str(n):>4} {str(m2):>6}")


if __name__ == "__main__":
    main()
