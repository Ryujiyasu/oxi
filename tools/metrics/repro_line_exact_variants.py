"""Additive variant tests for line=exact boundary rule.

Each variant = fresh minimal docx from scratch, different parameter.
Confirms Word's rule: at a B→C boundary where B and C have different
line=X exact values, Word uses B's pitch for the B→C advance.

Variants:
  V1: baseline (A=260exact, B=260exact empty, C=300exact) — already confirmed
  V2: non-empty B (A=260exact, B=260exact "BBB", C=300exact "CCC")
  V3: different fonts (A=MS Mincho, B=MS Mincho empty, C=MS Gothic)
  V4: larger deltas (A=240exact, B=240exact, C=400exact)
  V5: DECREASING (A=400exact, B=400exact, C=240exact)
  V6: no line=exact (control — should not exhibit the bug)
"""
import os, subprocess, sys, time, zipfile, json
from pathlib import Path
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

TMP_DIR = Path("pipeline_data") / "_repro_variants"
TMP_DIR.mkdir(parents=True, exist_ok=True)

CT = '<?xml version="1.0"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'


def make_para(line_tw, text, font="ＭＳ 明朝", size_half=21, rule="exact"):
    rpr = f'<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/><w:sz w:val="{size_half}"/><w:szCs w:val="{size_half}"/></w:rPr>'
    if line_tw is not None:
        spacing = f'<w:spacing w:line="{line_tw}" w:lineRule="{rule}"/>'
    else:
        spacing = ''
    body = f'<w:r>{rpr}<w:t>{text}</w:t></w:r>' if text else ''
    return f'<w:p><w:pPr>{spacing}{rpr}</w:pPr>{body}</w:p>'


def build(path, paragraphs):
    sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>'
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{"".join(paragraphs)}{sect}</w:body></w:document>'
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", xml)


def measure_word(path):
    subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True, timeout=5)
    time.sleep(0.5)
    word = win32com.client.DispatchEx("Word.Application")
    try:
        try: word.Visible = False
        except: pass
        try: word.DisplayAlerts = False
        except: pass
        doc = word.Documents.Open(str(Path(path).resolve()), ReadOnly=True)
        time.sleep(0.3)
        doc.Repaginate()
        n = doc.Paragraphs.Count
        ys = []
        for i in range(1, n + 1):
            try:
                y = doc.Paragraphs(i).Range.Information(6)
                ys.append(round(y, 2))
            except Exception:
                ys.append(None)
        doc.Close(False)
        return ys
    finally:
        try: word.Quit()
        except: pass


VARIANTS = [
    ("V1_baseline", [
        make_para(260, "AAA"),
        make_para(260, ""),
        make_para(300, "CCC"),
    ]),
    ("V2_nonempty_B", [
        make_para(260, "AAA"),
        make_para(260, "BBB"),
        make_para(300, "CCC"),
    ]),
    ("V3_gothic_C", [
        make_para(260, "AAA"),
        make_para(260, ""),
        make_para(300, "CCC", font="ＭＳ ゴシック"),
    ]),
    ("V4_larger_delta", [
        make_para(240, "AAA"),
        make_para(240, ""),
        make_para(400, "CCC"),
    ]),
    ("V5_decreasing", [
        make_para(400, "AAA"),
        make_para(400, ""),
        make_para(240, "CCC"),
    ]),
    ("V6_no_exact", [
        make_para(None, "AAA"),
        make_para(None, ""),
        make_para(None, "CCC"),
    ]),
]


def main():
    for label, paras in VARIANTS:
        path = TMP_DIR / f"{label}.docx"
        build(path, paras)
        ys = measure_word(path)
        if ys and len(ys) >= 3 and all(y is not None for y in ys):
            ab = ys[1] - ys[0]
            bc = ys[2] - ys[1]
            ac = ys[2] - ys[0]
            print(f"[{label}] A={ys[0]} B={ys[1]} C={ys[2]} | A→B={ab:.2f} B→C={bc:.2f} A→C={ac:.2f}")
        else:
            print(f"[{label}] err: {ys}")


if __name__ == "__main__":
    main()
