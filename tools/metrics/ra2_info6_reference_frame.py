"""Resolve Information(6) reference frame inconsistency.

§3.3 spec sweep: Info(6) = glyph top (= topMargin + centering offset)
§13.6 b35 verify: Info(6) = line-box top (= topMargin exactly)

Hypotheses to test (axes-isolated):
  1. Font: MS Gothic / MS Mincho / TNR / Calibri
  2. LineSpacing: not specified / Single (rule=0) / explicit val
  3. docGrid: lp=360 (18pt) / no docGrid

Build minimal repros via raw OOXML and measure first paragraph Info(6).
"""
import os, sys, time, json
import zipfile
import win32com.client

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

OUT = os.path.join(os.path.dirname(__file__), "output", "info6_reference_frame.json")
os.makedirs(os.path.dirname(OUT), exist_ok=True)
FIX_DIR = os.path.join(os.path.dirname(__file__), "output", "info6_ref_frame_fixtures")
os.makedirs(FIX_DIR, exist_ok=True)


def build_docx(out_path, font_ascii, font_eastasia, sz_halfpt, line_setting, with_docgrid):
    """Build minimal docx with single text paragraph.

    line_setting: None | ("auto", val) | ("exact", val) | ("single",) (i.e. rule=0)
    """
    rfonts = f'<w:rFonts w:ascii="{font_ascii}" w:eastAsia="{font_eastasia}" w:hAnsi="{font_ascii}"/>'
    sz = f'<w:sz w:val="{sz_halfpt}"/><w:szCs w:val="{sz_halfpt}"/>'

    if line_setting is None:
        spacing = ''
    elif line_setting[0] == "single":
        spacing = '<w:spacing w:line="240" w:lineRule="auto"/>'  # explicit Single
    elif line_setting[0] == "auto":
        spacing = f'<w:spacing w:line="{line_setting[1]}" w:lineRule="auto"/>'
    elif line_setting[0] == "exact":
        spacing = f'<w:spacing w:line="{line_setting[1]}" w:lineRule="exact"/>'
    else:
        spacing = ''

    docgrid = '<w:docGrid w:type="lines" w:linePitch="360"/>' if with_docgrid else ''

    doc_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>{spacing}<w:rPr>{rfonts}{sz}</w:rPr></w:pPr>
      <w:r><w:rPr>{rfonts}{sz}</w:rPr><w:t>Test</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      {docgrid}
    </w:sectPr>
  </w:body>
</w:document>'''

    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", doc_xml)


def restart_word():
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    time.sleep(3.0)
    return word


def measure(word, path):
    last_err = None
    for attempt in range(3):
        try:
            wdoc = word.Documents.Open(path)
            break
        except Exception as e:
            last_err = e
            time.sleep(2)
            try:
                while word.Documents.Count > 0:
                    word.Documents(1).Close(False)
            except Exception:
                pass
    else:
        raise last_err
    try:
        wdoc.Repaginate()
        time.sleep(0.1)
        p = wdoc.Paragraphs(1).Range
        y_para = round(p.Information(6), 4)
        try:
            y_char = round(p.Characters(1).Information(6), 4)
        except Exception:
            y_char = None
        return {"y_para": y_para, "y_char": y_char}
    finally:
        wdoc.Close(False)


def main():
    word = restart_word()
    results = []

    # 4 fonts × 4 sizes × 4 line settings × 2 docgrid = 128 cases
    # Trim to most informative subset
    fonts = [
        ("Times New Roman", "MS Mincho"),  # Latin + CJK fallback
        ("Calibri", "MS Mincho"),
        ("MS Gothic", "MS Gothic"),
        ("MS Mincho", "MS Mincho"),
    ]
    sz_pts = [10.5, 12.0, 14.0]  # spec example sizes
    line_settings = [
        None,  # No spacing element
        ("single",),  # Explicit Single
    ]
    docgrid_options = [True, False]

    try:
        for ascii_font, ea_font in fonts:
            for sz_pt in sz_pts:
                sz_hp = int(sz_pt * 2)
                for ls in line_settings:
                    for dg in docgrid_options:
                        ls_label = "none" if ls is None else "single"
                        dg_label = "grid" if dg else "noGrid"
                        label = f"{ascii_font.replace(' ','')}-{sz_pt}-{ls_label}-{dg_label}"
                        path = os.path.join(FIX_DIR, f"{label}.docx")
                        build_docx(path, ascii_font, ea_font, sz_hp, ls, dg)
                        try:
                            r = measure(word, path)
                            r.update({
                                "font_ascii": ascii_font, "font_ea": ea_font,
                                "sz_pt": sz_pt, "line": ls_label, "docgrid": dg_label,
                            })
                            results.append(r)
                            y = r["y_para"]
                            offset = y - 72.0  # topMargin = 72pt
                            print(f"  {label}: y={y} offset_from_topmargin={offset:+.2f}pt")
                        except Exception as e:
                            msg = str(e)
                            if "RPC" in msg or "拒否" in msg or "コール" in msg:
                                print(f"  {label}: COM failure → restart")
                                try: word.Quit()
                                except: pass
                                time.sleep(3)
                                word = restart_word()
                            else:
                                print(f"  {label}: ERR {msg[:60]}")
    finally:
        with open(OUT, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        print(f"\nSaved {len(results)} records to {OUT}")
        try: word.Quit()
        except: pass


if __name__ == "__main__":
    main()
