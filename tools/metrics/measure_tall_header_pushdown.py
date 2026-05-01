"""
Ra: Tall-header pushdown formula (spec §8.2 TBD).

Goal: derive body_y as a function of (header_lines, header_font, header_size,
header_distance, top_margin) when the header content overflows the headerDistance
and crosses topMargin.

Hypothesis (initial): body_y = max(top_margin, header_distance + n * header_lh + buffer)

Sweep:
  header_n_lines ∈ {1, 2, 3, 4, 5}
  header_font/size ∈ {(MS Gothic, 10.5), (Calibri, 11)}
  header_distance_tw ∈ {360, 720, 1080}      # 18, 36, 54 pt
  top_margin_tw ∈ {720, 1440, 2160}          # 36, 72, 108 pt
  noGrid only (no docGrid element)

Total: 5 * 2 * 3 * 3 = 90 records.

Output: tools/metrics/output/tall_header_pushdown.json
"""
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

OUT = Path(__file__).with_name("output") / "tall_header_pushdown.json"
TMP_DIR = Path("pipeline_data") / "_tall_header_tmp"
TMP_DIR.mkdir(parents=True, exist_ok=True)

CT = '''<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
</Types>'''

RELS_PKG = '''<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

RELS_DOC = '''<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>'''


def header_xml(font, sz_half, n_lines):
    pieces = []
    for i in range(n_lines):
        pieces.append(
            f'<w:r><w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
            f'<w:sz w:val="{sz_half}"/><w:szCs w:val="{sz_half}"/></w:rPr>'
            f'<w:t xml:space="preserve">H{i+1}.</w:t></w:r>'
        )
        if i < n_lines - 1:
            pieces.append('<w:r><w:br/></w:r>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:p><w:pPr><w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
        f'<w:sz w:val="{sz_half}"/><w:szCs w:val="{sz_half}"/></w:rPr></w:pPr>'
        f'{"".join(pieces)}</w:p></w:hdr>'
    )


def body_xml(font, sz_half, top_tw, header_tw):
    """Document body — 3 plain paragraphs, sectPr defines header link + margins."""
    p = lambda i: (
        f'<w:p><w:pPr><w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
        f'<w:sz w:val="{sz_half}"/><w:szCs w:val="{sz_half}"/></w:rPr></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
        f'<w:sz w:val="{sz_half}"/><w:szCs w:val="{sz_half}"/></w:rPr>'
        f'<w:t xml:space="preserve">B{i}.</w:t></w:r></w:p>'
    )
    section = (
        '<w:sectPr>'
        f'<w:headerReference w:type="default" r:id="rIdH1"/>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        f'<w:pgMar w:top="{top_tw}" w:right="851" w:bottom="1134" w:left="851" '
        f'w:header="{header_tw}" w:footer="992" w:gutter="0"/>'
        '<w:cols w:space="425"/>'
        '</w:sectPr>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<w:body>{p(1) + p(2) + p(3)}{section}</w:body></w:document>'
    )


def write_docx(path, font, sz_half, n_lines, top_tw, header_tw):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS_PKG)
        z.writestr("word/_rels/document.xml.rels", RELS_DOC)
        z.writestr("word/document.xml", body_xml(font, sz_half, top_tw, header_tw))
        z.writestr("word/header1.xml", header_xml(font, sz_half, n_lines))


def open_measure(word, path):
    last = None
    for attempt in range(3):
        try:
            doc = word.Documents.Open(str(path.resolve()), ReadOnly=True)
            time.sleep(0.2)
            ys = []
            for i in range(1, 4):
                try:
                    y = doc.Paragraphs(i).Range.Information(6)
                    ys.append(round(y, 3))
                except Exception:
                    break
            doc.Close(SaveChanges=False)
            return ys
        except Exception as e:
            last = e
            time.sleep(0.5 + attempt * 0.5)
    raise last


FONTS = [("Calibri", 11.0), ("MS Gothic", 10.5)]
HEADER_LINES = [1, 2, 3, 4, 5]
HEADER_DIST_TW = [360, 720, 1080]      # 18, 36, 54 pt
TOP_MARGIN_TW = [720, 1440, 2160]      # 36, 72, 108 pt


def main():
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    time.sleep(2.0)
    word.Visible = False
    word.DisplayAlerts = False
    results = []
    idx = 0
    try:
        total = len(FONTS) * len(HEADER_LINES) * len(HEADER_DIST_TW) * len(TOP_MARGIN_TW)
        i = 0
        for font, size in FONTS:
            sz_half = int(round(size * 2))
            for n in HEADER_LINES:
                for hd_tw in HEADER_DIST_TW:
                    for top_tw in TOP_MARGIN_TW:
                        i += 1; idx += 1
                        path = TMP_DIR / f"hd_{idx:04d}_{uuid.uuid4().hex[:8]}.docx"
                        rec = {
                            "font": font, "size": size, "n_lines": n,
                            "header_dist_tw": hd_tw, "top_margin_tw": top_tw,
                            "header_dist_pt": hd_tw / 20, "top_margin_pt": top_tw / 20,
                        }
                        try:
                            write_docx(path, font, sz_half, n, top_tw, hd_tw)
                            ys = open_measure(word, path)
                            rec.update({"body_ys": ys, "p1_y": ys[0] if ys else None})
                            print(f"[{i:3d}/{total}] {font:>9} {size:>4.1f}pt n={n} hd={hd_tw/20:>5.1f}pt top={top_tw/20:>5.1f}pt -> p1_y={ys[0] if ys else 'ERR'}")
                        except Exception as e:
                            rec["error"] = str(e)
                            print(f"[{i:3d}/{total}] ERR: {e}")
                        try:
                            path.unlink()
                        except Exception:
                            pass
                        results.append(rec)
    finally:
        try:
            word.Quit()
        except Exception:
            pass
        for f in TMP_DIR.glob("*.docx"):
            try: f.unlink()
            except: pass

    OUT.write_text(json.dumps(results, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nSaved -> {OUT}")


if __name__ == "__main__":
    main()
