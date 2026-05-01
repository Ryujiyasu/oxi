"""
Ra: §1.7 Mixed Font Run × grid centering_lh formula validation.

Hypothesis (extending grid_centering_lh_2026_05_02.md):
  For a mixed-font line, centering_lh = round(max over all per-run values of
    {natural_lh(run.font, run.size), run.size × 83/64})

Test combinations (all in grid mode, pitch ∈ {18, 24}):
  (a) Calibri-11 alone               : centering=14, line_h=pitch (single)
  (b) Yu Mincho-11 alone              : centering=18, line_h=2×pitch at pitch=18
  (c) Calibri-11 + Yu Mincho-11 mixed : centering=18 (Yu dominates via natural_lh)
  (d) Calibri-18 alone                : centering=23, line_h=2×pitch at pitch≤21
  (e) Calibri-18 + Yu Mincho-11 mixed : centering=23 (Calibri-18 size×83/64 dominates)

Output: tools/metrics/output/grid_mixed_font_centering.json
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

OUT = Path(__file__).with_name("output") / "grid_mixed_font_centering.json"
TMP_DIR = Path("pipeline_data") / "_grid_mixed_tmp"
TMP_DIR.mkdir(parents=True, exist_ok=True)

CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'

# Combinations
COMBOS = [
    {"label": "Calibri-11_alone",         "runs": [("Calibri", 22, "Hello")]},
    {"label": "YuMincho-11_alone",        "runs": [("Yu Mincho", 22, "あ"), ("Yu Mincho", 22, "い")]},
    {"label": "Calibri-11+YuMincho-11",   "runs": [("Calibri", 22, "Hello"), ("Yu Mincho", 22, "あ")]},
    {"label": "Calibri-18_alone",         "runs": [("Calibri", 36, "Hello")]},
    {"label": "Calibri-18+YuMincho-11",   "runs": [("Calibri", 36, "Hello"), ("Yu Mincho", 22, "あ")]},
    {"label": "MSMincho-10.5_alone",      "runs": [("ＭＳ 明朝", 21, "あい")]},
    {"label": "MSMincho-10.5+Calibri-18", "runs": [("ＭＳ 明朝", 21, "あい"), ("Calibri", 36, "Hello")]},
    {"label": "MSMincho-18_alone",        "runs": [("ＭＳ 明朝", 36, "あ")]},
    {"label": "Meiryo-11_alone",          "runs": [("Meiryo", 22, "あい")]},
    {"label": "Meiryo-11+Calibri-11",     "runs": [("Meiryo", 22, "あ"), ("Calibri", 22, "Hello")]},
]
PITCH_TW = [360, 480]   # 18, 24 pt


def run_xml(font, sz_half, text):
    return (
        f'<w:r><w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
        f'<w:sz w:val="{sz_half}"/><w:szCs w:val="{sz_half}"/></w:rPr>'
        f'<w:t xml:space="preserve">{text}</w:t></w:r>'
    )


def paragraph(runs):
    rxml = "".join(run_xml(f, sz, t) for f, sz, t in runs)
    # Use first run's font/size in pPr for default
    f, sz, _ = runs[0]
    ppr = (f'<w:pPr><w:rPr><w:rFonts w:ascii="{f}" w:eastAsia="{f}" w:hAnsi="{f}"/>'
           f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr></w:pPr>')
    return f'<w:p>{ppr}{rxml}</w:p>'


def doc_body(combo, pitch_tw):
    # 3 paragraphs: P1 = sentinel (small font for stable y0), P2 = combo, P3 = sentinel
    sentinel = paragraph([("Calibri", 22, "X")])
    target = paragraph(combo["runs"])
    paras = sentinel + target + sentinel
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


def write_docx(path, combo, pitch_tw):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", doc_body(combo, pitch_tw))


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
        total = len(COMBOS) * len(PITCH_TW)
        i = 0
        for combo in COMBOS:
            for pitch_tw in PITCH_TW:
                i += 1; idx += 1
                path = TMP_DIR / f"gmf_{idx:04d}_{uuid.uuid4().hex[:8]}.docx"
                rec = {"label": combo["label"], "pitch_pt": pitch_tw / 20, "pitch_tw": pitch_tw}
                try:
                    write_docx(path, combo, pitch_tw)
                    ys = measure(word, path)
                    rec["paragraph_ys"] = ys
                    rec["gap12"] = round(ys[1] - ys[0], 3)
                    rec["gap23"] = round(ys[2] - ys[1], 3)
                    print(f"[{i:2d}/{total}] {combo['label']:>30} pitch={pitch_tw/20:>4.1f} -> y={ys} gap12={rec['gap12']} gap23={rec['gap23']}")
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
