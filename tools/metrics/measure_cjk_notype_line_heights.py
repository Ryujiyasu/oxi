# -*- coding: utf-8 -*-
"""Measure Word's NO-TYPE-docGrid / LM0 base (1.0x single) line height per
(CJK font, size) — the GDI-accurate natural CJK line height that Oxi's 83/64
device-snap only approximates (off by +-1px at scattered large sizes; the
gen2 title box-height residual, see memory/gen2_vertical_drift.md #3 RESIDUAL).

VERIFIED FACTS (this session, 2026-06-19):
  - The no-type docGrid line height == LM0 (no docGrid) line height, and is
    PITCH-INDEPENDENT (linePitch 288 == 360 == LM0). Typed grids differ.
  - 1.0x single spacing is the base; multiple-spacing factors multiply it.
  - Ground truth = Word COM Information(6) line-box-top gap between consecutive
    same-size lines (collapsed-start range per the R30 fix).

METHOD: one doc per font, no-type docGrid, paragraphs = [sizeS][sizeS][sizeS]
per size S (3 lines so the two inner gaps both equal the line height and we can
detect any anomaly), 1.0x line spacing. Open in Word, read Info(6) per para,
take the median consecutive-gap per size block.

Output: pipeline_data/cjk_notype_line_heights.json  {font: {size_str: lh_pt}}
Also prints the per-size delta vs Oxi's 83/64 natural (floor-device-snap) so the
correction is visible.

Usage:
  python tools/metrics/measure_cjk_notype_line_heights.py            # all fonts
  python tools/metrics/measure_cjk_notype_line_heights.py "MS Mincho" # one font
Run on a Windows machine with the fonts installed (uninstalled fonts are
detected via GDI GetTextFace and skipped).
"""
import json, os, sys, zipfile, math
import win32com.client, pythoncom
import ctypes
from ctypes import wintypes

REPO = r"c:\Users\ryuji\oxi-main"
OUTDIR = os.path.join(REPO, "tools", "golden-test", "repros", "gen2_lineheight")
OUTJSON = os.path.join(REPO, "pipeline_data", "cjk_notype_line_heights.json")

# (json-key family name, docx eastAsia name to embed). docx name uses the
# fullwidth Japanese facename Word resolves; json key is the canonical Oxi family.
FONTS = [
    ("MS Mincho",   "ＭＳ 明朝"),
    ("MS Gothic",   "ＭＳ ゴシック"),
    ("MS PMincho",  "ＭＳ Ｐ明朝"),
    ("MS PGothic",  "ＭＳ Ｐゴシック"),
    ("Meiryo",      "メイリオ"),
    ("Yu Mincho",   "游明朝"),
    ("Yu Gothic",   "游ゴシック"),
    ("HGMinchoE",   "HG明朝E"),
    ("HGSMinchoE",  "HGS明朝E"),
    ("HGGothicE",   "HGゴシックE"),
]

# Sizes: half-points 8..40 (covers body..title) + a few large title sizes.
SIZES = [round(8 + 0.5 * i, 1) for i in range(0, 65)] + [42.0, 48.0, 54.0]


def gdi_facename(docx_ea_name):
    """Resolve what GDI actually maps the requested facename to (detect fallback)."""
    g = ctypes.windll.gdi32; u = ctypes.windll.user32
    LF_FACESIZE = 32

    class LOGFONT(ctypes.Structure):
        _fields_ = [("lfHeight", wintypes.LONG), ("lfWidth", wintypes.LONG),
                    ("lfEscapement", wintypes.LONG), ("lfOrientation", wintypes.LONG),
                    ("lfWeight", wintypes.LONG), ("lfItalic", wintypes.BYTE),
                    ("lfUnderline", wintypes.BYTE), ("lfStrikeOut", wintypes.BYTE),
                    ("lfCharSet", wintypes.BYTE), ("lfOutPrecision", wintypes.BYTE),
                    ("lfClipPrecision", wintypes.BYTE), ("lfQuality", wintypes.BYTE),
                    ("lfPitchAndFamily", wintypes.BYTE), ("lfFaceName", ctypes.c_wchar * LF_FACESIZE)]
    hdc = u.GetDC(0)
    lf = LOGFONT(); lf.lfHeight = -20; lf.lfCharSet = 128; lf.lfFaceName = docx_ea_name
    hf = g.CreateFontIndirectW(ctypes.byref(lf)); old = g.SelectObject(hdc, hf)
    buf = ctypes.create_unicode_buffer(LF_FACESIZE)
    g.GetTextFaceW(hdc, LF_FACESIZE, buf)
    g.SelectObject(hdc, old); g.DeleteObject(hf); u.ReleaseDC(0, hdc)
    return buf.value


def build_doc(path, ea_name):
    CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>')
    RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    DOCRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
    STY = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr>'
           f'<w:rFonts w:ascii="Cambria" w:hAnsi="Cambria" w:eastAsia="{ea_name}"/><w:sz w:val="22"/><w:szCs w:val="22"/>'
           '<w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault>'
           '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
           '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')

    def para(hp):
        return ('<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
                f'<w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria" w:eastAsia="{ea_name}"/><w:sz w:val="{hp}"/></w:rPr></w:pPr>'
                f'<w:r><w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria" w:eastAsia="{ea_name}"/><w:sz w:val="{hp}"/></w:rPr>'
                '<w:t>漢</w:t></w:r></w:p>')
    blocks = []
    for s in SIZES:
        hp = int(round(s * 2))
        blocks += [para(hp), para(hp), para(hp)]   # 3 lines per size
    DOC = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
           + "".join(blocks) +
           '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
           '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
           '<w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOCRELS); z.writestr("word/styles.xml", STY)
        z.writestr("word/document.xml", DOC)


def measure_font(word, family, ea_name):
    resolved = gdi_facename(ea_name)
    # crude install check: GDI returns the requested face (or a close variant)
    path = os.path.join(OUTDIR, "cjklh_probe.docx")
    build_doc(path, ea_name)
    d = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
    ys = []
    n = d.Paragraphs.Count
    for i in range(1, n + 1):
        rng = d.Paragraphs(i).Range
        cs = d.Range(rng.Start, rng.Start)
        ys.append(float(cs.Information(6)))
    d.Close(False)
    out = {}
    for bi, s in enumerate(SIZES):
        i0 = bi * 3
        if i0 + 2 >= len(ys):
            break
        g1 = ys[i0 + 1] - ys[i0]; g2 = ys[i0 + 2] - ys[i0 + 1]
        # both gaps should equal the line height; a page break makes one huge -> drop it
        cands = [g for g in (g1, g2) if 0 < g < 200]
        if not cands:
            continue
        out[format_size(s)] = round(min(cands), 3)
    return resolved, out


def format_size(s):
    return str(int(s)) if float(s).is_integer() else str(s)


def oxi_8364_floor(fs):
    """Oxi's current no-type natural device-snap (83/64, win_sum=1 MS family)."""
    nat = fs * (83.0 / 64.0)
    return math.floor(nat / 0.75) * 0.75


def main():
    pythoncom.CoInitialize()
    sel = sys.argv[1] if len(sys.argv) > 1 else None
    fonts = [(k, e) for (k, e) in FONTS if sel is None or k == sel]
    w = win32com.client.DispatchEx("Word.Application"); w.Visible = False
    result = {}
    try:
        for family, ea in fonts:
            try:
                resolved, data = measure_font(w, family, ea)
            except Exception as e:
                print(f"!! {family}: measure error {e}"); continue
            installed = data and (resolved.lower().replace(" ", "")[:4] in family.lower().replace(" ", "")
                                  or family.lower().replace(" ", "")[:4] in resolved.lower().replace(" ", "")
                                  or len(data) > 20)
            note = "" if installed else f"  (GDI->{resolved!r}; possibly NOT installed -- verify)"
            print(f"\n=== {family}  (docx ea={ea!r}; GDI face={resolved!r}){note} ===")
            if not data:
                print("   no data"); continue
            result[family] = data
            # show a sampled comparison vs Oxi 83/64 for MS-family
            print("   sz   Word_lh  Oxi8364  delta")
            for s in [10.5, 11, 12, 13, 14, 16, 18, 20, 22, 24, 26, 28, 36]:
                k = format_size(s)
                if k in data:
                    o = oxi_8364_floor(s)
                    print(f"   {s:5} {data[k]:7.2f}  {o:7.2f}  {data[k]-o:+.2f}")
    finally:
        w.Quit()
    os.makedirs(os.path.dirname(OUTJSON), exist_ok=True)
    if os.path.exists(OUTJSON):
        prev = json.load(open(OUTJSON, encoding="utf-8"))
        prev.update(result); result = prev
    json.dump(result, open(OUTJSON, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"\nwrote {OUTJSON}: {list(result.keys())}")


if __name__ == "__main__":
    main()
