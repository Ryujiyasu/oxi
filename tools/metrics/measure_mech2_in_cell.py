"""§4.7b round 28 — Mech 2 inside table cell vs body paragraph.

Round 16/22 closed Mech 2 cap formula at body-level (cap=fs/2 mixed,
0.75×fs pure-yak). Open question: does Mech 2 fire identically inside
table cells, or is the cap modified by cell context (cellMar, border,
jc inheritance)?

Test design:
  Body probe: paragraph with 24-char `「漢」漢×6` at cw=N pt, jc=both
  Cell probe: 1-cell table, cell_w = N + 2 × cellMar (= 4.95pt default)
              same 24-char content, cellMar=4.95pt (default)
              cell_inner_content_w = cell_w - 2 × cellMar = N pt

If Mech 2 cap inside cell = body cap, both should fit/wrap at same
slack values. If cell context modifies cap, results diverge.

Probe: 「漢」漢「漢」漢「漢」漢「漢」漢「漢」漢「漢」漢
       (24 chars, 12 mid-line yak, no Mech 1, cap_pred = 6.0 at fs=12)

Sweep slacks {3, 5, 6, 7, 8, 10} for body and cell at fs=12 MS Mincho.
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\mech2_cell_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\mech2_cell.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"
CELL_MAR_PT = 4.95  # default LeftPadding/RightPadding


def make_body_doc_xml(probe, sz_val, page_w_tw, margin_tw):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            '<w:pPr><w:jc w:val="both"/>'
            '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            '</w:pPr>'
            '<w:r><w:rPr>'
            f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/></w:rPr>'
            f'<w:t>{probe}</w:t></w:r></w:p>'
            f'<w:sectPr><w:pgSz w:w="{page_w_tw}" w:h="16838"/>'
            f'<w:pgMar w:top="1134" w:right="{margin_tw}" w:bottom="1134" w:left="{margin_tw}"'
            ' w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:linePitch="360"/>'
            '</w:sectPr></w:body></w:document>')


def make_cell_doc_xml(probe, sz_val, page_w_tw, margin_tw, cell_w_tw):
    """1-row, 1-cell table. Default cellMar = 99 twips (4.95pt)."""
    cell_mar_tw = int(round(CELL_MAR_PT * 20))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body>'
            '<w:tbl>'
            '<w:tblPr>'
            f'<w:tblW w:w="{cell_w_tw}" w:type="dxa"/>'
            f'<w:tblCellMar>'
            f'<w:top w:w="0" w:type="dxa"/>'
            f'<w:left w:w="{cell_mar_tw}" w:type="dxa"/>'
            f'<w:bottom w:w="0" w:type="dxa"/>'
            f'<w:right w:w="{cell_mar_tw}" w:type="dxa"/>'
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
                    # Filter all marker chars (single \r, \x07, or combinations like '\r\x07')
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
        return {"n_chars_line1": n_line1, "char_advances": char_advances}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    probe = "「漢」漢" * 6  # 24 chars: 12 mid-line yak (alternating「漢」漢)
    # No Mech 1 fires (yak adjacent to CJK). natural = 24×12 = 288pt
    # Round 16 cap_pred = fs/2 = 6.0pt at fs=12
    fs_pt = 12.0
    sz_val = 24
    margin_tw = 170 * 10

    slacks = [3.0, 5.0, 6.0, 7.0, 8.0, 10.0]
    nat = 288.0

    for slack in slacks:
        # Body: page content width = nat - slack
        body_cw_pt = nat - slack
        body_page_w_tw = int((body_cw_pt + 170) * 20)
        body_doc_xml = make_body_doc_xml(probe, sz_val, body_page_w_tw, margin_tw)

        # Cell: cell_w = inner content + 2 × cellMar
        # Page width = cell_w + 2 × page_margin (170pt each)
        cell_inner_pt = body_cw_pt
        cell_w_pt = cell_inner_pt + 2 * CELL_MAR_PT
        cell_w_tw = int(round(cell_w_pt * 20))
        cell_page_w_tw = int((cell_w_pt + 170) * 20)
        cell_doc_xml = make_cell_doc_xml(probe, sz_val, cell_page_w_tw, margin_tw, cell_w_tw)

        # Build + measure body
        body_label = f"body_sl{slack}"
        try:
            p_body = make_docx(body_label, body_doc_xml, sz_val)
        except Exception as e:
            out[f"slack_{slack}"] = {"build_error_body": str(e)}
            continue
        kill_word()
        try:
            r_body = measure_one(p_body)
        except Exception as e:
            r_body = {"measure_error": str(e)}
            kill_word()

        # Build + measure cell
        cell_label = f"cell_sl{slack}"
        try:
            p_cell = make_docx(cell_label, cell_doc_xml, sz_val)
        except Exception as e:
            out[f"slack_{slack}"]["build_error_cell"] = str(e)
            continue
        kill_word()
        try:
            r_cell = measure_one(p_cell)
        except Exception as e:
            r_cell = {"measure_error": str(e)}
            kill_word()

        # Compute mech2 compression for body and cell
        def comp_mech2(r):
            if "char_advances" not in r:
                return None, None
            advs = r["char_advances"]
            sum_actual = sum(c["adv"] for c in advs)
            sum_natural = sum(fs_pt for _ in advs)  # all CJK fs=12
            return r.get("n_chars_line1"), round(sum_natural - sum_actual, 2)

        body_n, body_comp = comp_mech2(r_body)
        cell_n, cell_comp = comp_mech2(r_cell)

        out[f"slack_{slack}"] = {
            "slack": slack,
            "body_cw_pt": round(body_cw_pt, 2),
            "cell_inner_pt": round(cell_inner_pt, 2),
            "cell_w_pt": round(cell_w_pt, 2),
            "body_n": body_n,
            "body_mech2_comp": body_comp,
            "cell_n": cell_n,
            "cell_mech2_comp": cell_comp,
        }
        print(f"slack={slack:>4.1f} cw={body_cw_pt:>6.1f} | body n={body_n} m2={body_comp} | cell n={cell_n} m2={cell_comp}")
        with open(RESULT, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)

    print("\n========== SUMMARY ==========")
    print(f"{'slack':>6} {'body n':>8} {'body m2':>9} {'cell n':>8} {'cell m2':>9} {'match':>8}")
    for slack in slacks:
        info = out.get(f"slack_{slack}", {})
        bn = info.get("body_n")
        cn = info.get("cell_n")
        bm = info.get("body_mech2_comp")
        cm = info.get("cell_mech2_comp")
        match = "✓" if (bn == cn and bm == cm) else "✗"
        print(f"  {slack:>4} {str(bn):>8} {str(bm):>9} {str(cn):>8} {str(cm):>9} {match:>8}")


if __name__ == "__main__":
    main()
