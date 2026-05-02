"""
Ra: per-line and per-paragraph Y position in LM≥1 (docGrid type=lines/linesAndChars).

Goal: derive the formula for line Y across:
  - Multiple lines within a single paragraph (via <w:br/>)
  - Multiple paragraphs (via <w:p>)

Hypothesis (initial):
  line_k_y = topMargin + (cumulative_grid_n_so_far) * pitch + (pitch - lh_natural) / 2

For multi-line single paragraph (n lines, all same font):
  line_k_y = topMargin + (k-1) * pitch + (pitch - lh)/2  if pitch > lh (single-cell snap)
           = topMargin + (k-1) * line_height + (line_height - lh)/2  if pitch < lh (multi-cell)
  where line_height = ceil(lh / pitch) * pitch

For multi-paragraph (3 paragraphs of 1 line each):
  P1_y = topMargin + (pitch - lh)/2
  P2_y = P1_y + pitch + sa_grid_extra ?
  P3_y = P2_y + pitch + sa_grid_extra ?

Sweep:
  font ∈ {MS Mincho 10.5pt, Calibri 11pt}
  pitch_tw ∈ {360, 400, 480, 560, 640} (= 18, 20, 24, 28, 32 pt)
  layout: 3 paragraphs × 3 lines each (via <w:br/>)
  topMargin fixed at 1440tw (72pt)

Output: tools/metrics/output/grid_per_line_y.json
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

OUT = Path(__file__).with_name("output") / "grid_per_line_y.json"
TMP_DIR = Path("pipeline_data") / "_grid_per_line_tmp"
TMP_DIR.mkdir(parents=True, exist_ok=True)

CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'

FONTS = [
    # Critical test: Yu Mincho / Meiryo body lh ≠ size × 83/64
    # If grid centering uses round(size×83/64) UNIVERSALLY: line_h matches Calibri
    # If grid centering uses font-specific lh: line_h is much larger
    ("Yu Mincho", "Yu Mincho", 11.0),   # body lh ≈ 18.5pt, but size×83/64 = 14.27
    ("Meiryo", "Meiryo", 11.0),         # body lh ≈ 21.5pt, but size×83/64 = 14.27
    ("ＭＳ 明朝", "MS Mincho", 10.5),   # baseline
    ("Calibri", "Calibri", 11.0),       # baseline
]
PITCH_TW = [240, 300, 360, 400, 480, 560, 640]  # 12, 15, 18, 20, 24, 28, 32 pt
N_PARAS = 3
N_LINES_PER_PARA = 3


def run_xml(font, sz_half, text):
    return (
        f'<w:r><w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
        f'<w:sz w:val="{sz_half}"/><w:szCs w:val="{sz_half}"/></w:rPr>'
        f'<w:t xml:space="preserve">{text}</w:t></w:r>'
    )


def ppr_xml(font, sz_half):
    return (
        f'<w:pPr><w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
        f'<w:sz w:val="{sz_half}"/><w:szCs w:val="{sz_half}"/></w:rPr></w:pPr>'
    )


def paragraph_xml(font, sz_half, n_lines, label):
    pieces = []
    for i in range(n_lines):
        pieces.append(run_xml(font, sz_half, f"{label}.{i+1}あ"))
        if i < n_lines - 1:
            pieces.append('<w:r><w:br/></w:r>')
    return f'<w:p>{ppr_xml(font, sz_half)}{"".join(pieces)}</w:p>'


def doc_body(font, sz_half, pitch_tw):
    paras = "".join(paragraph_xml(font, sz_half, N_LINES_PER_PARA, f"P{i+1}") for i in range(N_PARAS))
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


def write_docx(path, font, sz_half, pitch_tw):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", doc_body(font, sz_half, pitch_tw))


def measure(word, path):
    last = None
    for attempt in range(3):
        try:
            doc = word.Documents.Open(str(path.resolve()), ReadOnly=True)
            time.sleep(0.2)
            # Each paragraph: get y of each line via Range.Information(6) sweep
            paragraphs_y = []
            for pi in range(1, N_PARAS + 1):
                p = doc.Paragraphs(pi)
                rng = p.Range
                # Sweep characters; collect distinct y per line
                line_ys = []
                for ci in range(rng.Start, rng.End):
                    r = doc.Range(ci, ci + 1)
                    try:
                        y = r.Information(6)
                    except Exception:
                        continue
                    if not line_ys or abs(y - line_ys[-1]) > 0.5:
                        line_ys.append(round(y, 3))
                paragraphs_y.append(line_ys)
            doc.Close(SaveChanges=False)
            return paragraphs_y
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
        total = len(FONTS) * len(PITCH_TW)
        i = 0
        for font_xml, pretty, size in FONTS:
            sz_half = int(round(size * 2))
            for pitch_tw in PITCH_TW:
                i += 1; idx += 1
                path = TMP_DIR / f"gpl_{idx:04d}_{uuid.uuid4().hex[:8]}.docx"
                rec = {"font": pretty, "size": size, "pitch_tw": pitch_tw, "pitch_pt": pitch_tw / 20}
                try:
                    write_docx(path, font_xml, sz_half, pitch_tw)
                    paragraphs_y = measure(word, path)
                    rec["paragraphs_y"] = paragraphs_y
                    print(f"[{i:2d}/{total}] {pretty:>10} {size:>4.1f}pt pitch={pitch_tw/20:>4.1f}pt -> {paragraphs_y}")
                except Exception as e:
                    rec["error"] = str(e)
                    print(f"[{i:2d}/{total}] ERR: {e}")
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
