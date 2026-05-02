"""§4.x round 34 — topLinePunct line-start punctuation compression.

ECMA-376 §17.3.1.43 — paragraph property `<w:topLinePunct>`
"specifies whether punctuation shall be compressed at the start
of a line if needed to allow another character on that line".

Test: paragraph starting with 「 (Type A line-start opening yak).
With topLinePunct=on, 「 at line start may be compressed to half-width
to allow 1 more char on the line.

Probe: 「 + 漢×N (line starts with Type A yak).
Vary cw to observe how 「 is positioned with topLinePunct on/off.

If compression: 「 width = 6pt (half) instead of 12pt → 1 more char fits.
If no compression: 「 width = 12pt → fewer chars per line.
"""
import json, os, sys, time, zipfile, subprocess
import win32com.client as w32

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OUT_DIR = os.path.abspath(r"C:\Users\ryuji\oxi-1\tools\metrics\top_line_punct_repro")
RESULT = os.path.abspath(r"C:\Users\ryuji\oxi-1\pipeline_data\top_line_punct.json")
os.makedirs(OUT_DIR, exist_ok=True)

CJK_FONT = "ＭＳ 明朝"


def make_doc_xml(probe, sz_val, page_w_tw, margin_tw, top_line_punct=None):
    op_xml = ""
    if top_line_punct is True:
        op_xml = '<w:topLinePunct w:val="1"/>'
    elif top_line_punct is False:
        op_xml = '<w:topLinePunct w:val="0"/>'
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
        ys = sorted(set(x[2] for x in xs))
        lines = []
        for y in ys:
            line_chars = sorted([(t, x) for t, x, yy in xs if abs(yy - y) < 0.5], key=lambda v: v[1])
            advs = []
            for i in range(len(line_chars) - 1):
                advs.append(round(line_chars[i+1][1] - line_chars[i][1], 3))
            lines.append({
                "y": y,
                "n": len(line_chars),
                "first_x": line_chars[0][1] if line_chars else None,
                "first_char": line_chars[0][0] if line_chars else None,
                "first_char_adv": advs[0] if advs else None,
                "second_char": line_chars[1][0] if len(line_chars) > 1 else None,
                "chars": "".join(t for t, _ in line_chars),
            })
        return {"n_total": len(xs), "n_lines": len(lines), "lines": lines}
    finally:
        try: word.Quit()
        except: pass


def main():
    out = {}
    fs_pt = 12.0
    sz_val = 24
    margin_tw = 170 * 10

    # Probe: 「 at line start, then 漢×N
    # Vary cw to observe whether 「 is compressed at line start
    probe = "「" + "漢" * 20  # 21 chars, natural 21*12=252pt
    nat = 21 * 12.0

    cw_pts = [252.0, 248.0, 246.0, 240.0]  # range around natural
    settings = [
        ("def", None,  "no override"),
        ("on",  True,  "topLinePunct val=1"),
        ("off", False, "topLinePunct val=0"),
    ]

    print(f"Probe: 「 + 漢×20 (21 chars), natural={nat}pt\n")
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
            entry = {"cw_pt": cw_pt, "settings": s_label, "desc": desc, "tlp_val": op_val, **r}
            out[label] = entry
            n_lines = entry.get("n_lines", "?")
            lines = entry.get("lines", [])
            l1 = lines[0] if lines else {}
            print(f"  cw={cw_pt:>5} {s_label:>3} | n_lines={n_lines} line1: n={l1.get('n')} 「_x={l1.get('first_x')} 「_adv={l1.get('first_char_adv')}")
            with open(RESULT, "w", encoding="utf-8") as f:
                json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
