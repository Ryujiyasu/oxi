"""
Ra: paragraph spacing (sa/sb) interaction with docGrid (LM≥1).

Spec §2.1 line 198 says "grid snap may also be applied to spacing" with one example
(sa=sb=10 → 9.75pt). This sweep pins the exact rule.

Hypotheses:
  A. sa is added directly: P2_y = P1_y + line_h + sa
  B. sa is grid-snapped: P2_y = P1_y + line_h + grid_snap(sa)
  C. cell-ceil: P2_y = P1_y + ceil((line_h + sa) / pitch) * pitch
  D. line_h includes sa: line_h_with_sa = ceil((natural_lh + sa) / pitch) * pitch

Sweep:
  font/size ∈ {(MS Mincho, 10.5), (Calibri, 11)}
  sa_tw ∈ {0, 80, 120, 160, 240, 360, 480}     # 0, 4, 6, 8, 12, 18, 24 pt
  pitch_tw ∈ {360, 480}                          # 18, 24 pt
  3 paragraphs (each 1 line) — measure y1, y2, y3

Output: tools/metrics/output/grid_paragraph_spacing.json
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

OUT = Path(__file__).with_name("output") / "grid_paragraph_spacing.json"
TMP_DIR = Path("pipeline_data") / "_grid_spacing_tmp"
TMP_DIR.mkdir(parents=True, exist_ok=True)

CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'

FONTS = [("ＭＳ 明朝", "MS Mincho", 10.5), ("Calibri", "Calibri", 11.0)]
# Include sub-integer-pt sa values to verify spec §2.1 "sa=10→9.75" claim:
# 200tw=10pt = 13.333px = round 13px = 9.75pt (if pixel-snap applies)
SA_TW = [0, 80, 100, 120, 150, 200, 210, 240, 300, 360, 400, 480]
# pt: 0, 4, 5, 6, 7.5, 10, 10.5, 12, 15, 18, 20, 24
PITCH_TW = [360, 480]                       # 18, 24 pt


def run_xml(font, sz_half, text):
    return (
        f'<w:r><w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
        f'<w:sz w:val="{sz_half}"/><w:szCs w:val="{sz_half}"/></w:rPr>'
        f'<w:t xml:space="preserve">{text}</w:t></w:r>'
    )


def ppr_xml(font, sz_half, sa_tw=None):
    spacing = f'<w:spacing w:after="{sa_tw}" w:before="0"/>' if sa_tw is not None else ""
    return (
        f'<w:pPr>{spacing}'
        f'<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
        f'<w:sz w:val="{sz_half}"/><w:szCs w:val="{sz_half}"/></w:rPr></w:pPr>'
    )


def paragraph(font, sz_half, label, sa_tw=None):
    return f'<w:p>{ppr_xml(font, sz_half, sa_tw)}{run_xml(font, sz_half, label)}</w:p>'


def doc_body(font, sz_half, sa_tw, pitch_tw):
    # 3 paragraphs; each has the spacing element to apply sa uniformly
    paras = "".join(paragraph(font, sz_half, f"P{i+1}あ", sa_tw) for i in range(3))
    section = (
        '<w:sectPr>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992" w:gutter="0"/>'
        f'<w:docGrid w:type="lines" w:linePitch="{pitch_tw}"/>'
        '</w:sectPr>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{paras}{section}</w:body></w:document>'
    )


def write_docx(path, font, sz_half, sa_tw, pitch_tw):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", doc_body(font, sz_half, sa_tw, pitch_tw))


def measure(word, path):
    last = None
    for attempt in range(3):
        try:
            doc = word.Documents.Open(str(path.resolve()), ReadOnly=True)
            time.sleep(0.2)
            ys = []
            for i in range(1, 4):
                y = doc.Paragraphs(i).Range.Information(6)
                ys.append(round(y, 3))
            doc.Close(SaveChanges=False)
            return ys
        except Exception as e:
            last = e
            time.sleep(0.5 + attempt * 0.5)
    raise last


def main():
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    time.sleep(2.0)
    word.Visible = False
    word.DisplayAlerts = False
    results = []
    idx = 0
    try:
        total = len(FONTS) * len(PITCH_TW) * len(SA_TW)
        i = 0
        for font_xml, pretty, size in FONTS:
            sz_half = int(round(size * 2))
            for pitch_tw in PITCH_TW:
                for sa_tw in SA_TW:
                    i += 1; idx += 1
                    path = TMP_DIR / f"gps_{idx:04d}_{uuid.uuid4().hex[:8]}.docx"
                    rec = {"font": pretty, "size": size, "pitch_pt": pitch_tw / 20,
                           "sa_pt": sa_tw / 20, "sa_tw": sa_tw, "pitch_tw": pitch_tw}
                    try:
                        write_docx(path, font_xml, sz_half, sa_tw, pitch_tw)
                        ys = measure(word, path)
                        rec["paragraph_ys"] = ys
                        gap12 = ys[1] - ys[0]
                        gap23 = ys[2] - ys[1]
                        rec["gap12"] = round(gap12, 3)
                        rec["gap23"] = round(gap23, 3)
                        print(f"[{i:3d}/{total}] {pretty:>10} {size:>4.1f}pt pitch={pitch_tw/20:>4.1f} sa={sa_tw/20:>5.2f}pt -> y={ys} gaps=({gap12},{gap23})")
                    except Exception as e:
                        rec["error"] = str(e)
                        print(f"[{i:3d}/{total}] ERR {e}")
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
