"""§4.x round 31 — overflowPunct line-end behavior.

ECMA-376 §17.3.1.32 — overflowPunct paragraph property:
  When true, line-ending punctuation may extend past the right margin
  (Japanese typography convention "ぶら下がり").

Test: a line that ends with 」 (closing yakumono Type B). Without
overflowPunct, 」 must fit within content width. With overflowPunct,
」 can extend beyond.

Probe construction:
  20 chars CJK + 」 at pos 21 (line-end candidate).
  Vary cw to test wrap behavior:
    cw = 20 × 12 + 12 (full fit)
    cw = 20 × 12 + 6  (force tight)
    cw = 20 × 12 + 0  (line ends just at 」 boundary)
    cw = 20 × 12 - 2  (force wrap test)

Settings:
  V_def: no overflowPunct override (defaults TBD)
  V_on:  pPr/overflowPunct val=1 explicit ON
  V_off: pPr/overflowPunct val=0 explicit OFF

Measure 」's x-position to determine if it overflows or wraps.
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\overflow_punct_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\overflow_punct.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"


def make_doc_xml(probe, sz_val, page_w_tw, margin_tw, overflow_punct=None):
    """overflow_punct: None (default), True (val=1), False (val=0)."""
    op_xml = ""
    if overflow_punct is True:
        op_xml = '<w:overflowPunct w:val="1"/>'
    elif overflow_punct is False:
        op_xml = '<w:overflowPunct w:val="0"/>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:body><w:p>'
            f'<w:pPr>{op_xml}<w:jc w:val="left"/>'
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


def make_styles_xml(sz_val):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults>'
            '<w:rPrDefault><w:rPr>'
            f'<w:rFonts w:ascii="{CJK_FONT}" w:hAnsi="{CJK_FONT}" w:eastAsia="{CJK_FONT}" w:hint="eastAsia"/>'
            f'<w:sz w:val="{sz_val}"/>'
            '</w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr/></w:pPrDefault>'
            '</w:docDefaults></w:styles>')


def make_settings_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
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
        # Determine number of lines
        ys = sorted(set(x[2] for x in xs))
        # Build per-line summary
        lines = []
        for y in ys:
            line_chars = sorted([(t, x) for t, x, yy in xs if abs(yy - y) < 0.5], key=lambda v: v[1])
            lines.append({"y": y, "n": len(line_chars), "first_x": line_chars[0][1] if line_chars else None,
                          "last_x": line_chars[-1][1] if line_chars else None,
                          "last_char": line_chars[-1][0] if line_chars else None})
        return {"n_total": len(xs), "n_lines": len(lines), "lines": lines}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    fs_pt = 12.0
    sz_val = 24
    margin_tw = 170 * 10  # 85pt margins each side

    # Probe: 20 CJK chars + 」 at pos 21
    # Natural width = 21 × 12 = 252pt
    probe = "漢" * 20 + "」"
    nat = 21 * 12.0  # 252pt
    print(f"Probe: 漢×20 + 」 (21 chars), natural={nat}pt\n")

    # cw values to test:
    # 252pt — exact fit, all 21 chars on line
    # 246pt — short by 6pt = the 」 width-half (might overflow with hangPunct)
    # 240pt — short by 12pt = exactly 1 char
    # 235pt — short by 17pt
    cw_pts = [252.0, 246.0, 240.0, 235.0]

    # Settings
    settings = [
        ("def", None,  "no overflowPunct override"),
        ("on",  True,  "overflowPunct val=1"),
        ("off", False, "overflowPunct val=0"),
    ]

    for cw_pt in cw_pts:
        page_w_tw = int((cw_pt + 170) * 20)
        for s_label, op_val, desc in settings:
            doc_xml = make_doc_xml(probe, sz_val, page_w_tw, margin_tw, op_val)
            label = f"cw{cw_pt}_{s_label}"
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
            entry = {"cw_pt": cw_pt, "settings": s_label, "desc": desc,
                     "op_val": op_val, **r}
            out[label] = entry
            n_lines = entry.get("n_lines", "?")
            lines = entry.get("lines", [])
            line1_n = lines[0]["n"] if lines else "?"
            line1_last_char = lines[0]["last_char"] if lines else "?"
            line1_last_x = lines[0]["last_x"] if lines else "?"
            print(f"  cw={cw_pt:>5} {s_label:>3} | n_lines={n_lines} line1: n={line1_n} last={line1_last_char!r} x={line1_last_x}")
            with open(RESULT, "w", encoding="utf-8") as f:
                json.dump(out, f, ensure_ascii=False, indent=2)

    print("\n========== SUMMARY ==========")
    print(f"  natural=252pt; right_margin_x = page_margin (85) + cw_pt = right edge")
    for cw_pt in cw_pts:
        right_edge_x = 85.0 + cw_pt
        print(f"\n  cw={cw_pt} (right edge x={right_edge_x}):")
        for s_label, _, _ in settings:
            label = f"cw{cw_pt}_{s_label}"
            info = out.get(label, {})
            lines = info.get("lines", [])
            n_lines = info.get("n_lines", "?")
            if lines:
                l1 = lines[0]
                last = l1.get("last_char")
                lx = l1.get("last_x")
                overflow = "(overflow!)" if isinstance(lx, (int, float)) and lx > right_edge_x - 12 else ""
                print(f"    {s_label:>3}: lines={n_lines} line1 n={l1['n']} last={last!r} last_x={lx} {overflow}")


if __name__ == "__main__":
    main()
