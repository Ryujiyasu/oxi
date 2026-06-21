# -*- coding: utf-8 -*-
"""Char-budget dataset generator.

Builds a minimal .docx whose body is the SAME 約物-rich CJK sentence repeated at
a swept right-indent (shrinking the available width in fine steps). Each width
shifts Word's line breaks, so across the sweep we capture many oikomi-vs-oidashi
decisions on identical content — the controlled input for deriving Word's
per-line 約物 compression rule (the char-budget wall).

Usage:
  python _cb_gen.py OUT.docx [--jc both|left] [--compat 15|14|11]
                    [--grid lines|linesAndChars|none] [--pitch 360]
                    [--font "MS Mincho"] [--sz 21]
                    [--ind0 0] [--ind1 60] [--step 0.5]   (right-indent twips? no: pt)
Indents are swept in POINTS, converted to twips (×20) in the docx.
"""
import sys, os, zipfile

# 約物-rich regulation-style base sentence: many 、 。 and varied trailing chars,
# long enough to wrap to several lines so each width yields multiple break points.
BASE1 = (
    "甲は、本契約に基づき、乙に対して、本件業務を、善良なる管理者の注意をもって、"
    "誠実かつ適切に遂行するものとし、これに関連する一切の責任を負う。"
    "なお、前項の規定にかかわらず、特別の事情がある場合には、別途協議のうえ、"
    "これを定めるものとする。"
)
# BASE2: LONG kanji runs separated by 約物, so a KANJI (not 約物) lands at the
# break boundary with a 約物 earlier on the line — the case-B test (does Word
# compress the mid-約物 to pull a trailing kanji onto the line = oikomi?).
BASE2 = (
    "本件業務遂行管理責任体制整備運用状況、定期報告書類作成提出義務履行確認、"
    "関係法令遵守状況点検記録保存、業務改善計画策定実施結果評価分析、"
    "年度末決算処理業務完了報告、次年度事業計画立案承認手続実施。"
)
# BASE3: many ADJACENT 約物 PAIRS (、「 」、 」。 「 etc.) to test 約物-pair kerning
# — the SEPARATE rule (adjacent 約物 kern by ~half-em) that real docs' "compression"
# (tokyoshugyo page-44 「、「→3.0」, S585c) actually is, vs general mid-約物 compress.
BASE3 = (
    "甲は、「本件業務」、「関連業務」、及び「付随業務」を、乙に対して、"
    "「別紙一」のとおり、委託するものとし、乙は、「これら」を、"
    "「善良なる管理者」の注意をもって、誠実に遂行する。"
)
def _pick(base):
    return {"1": BASE1, "2": BASE2, "3": BASE3}.get(str(base), BASE1)
BASE = BASE1

def esc(s): return s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def build(out, jc="both", compat="15", grid="lines", pitch="360",
          font="ＭＳ 明朝", sz="21", ind0=0.0, ind1=60.0, step=0.5, base="1", cpunct="1"):
    BASE = _pick(base)
    # docGrid element
    if grid == "none":
        docgrid = ""
    elif grid == "linesAndChars":
        docgrid = f'<w:docGrid w:type="linesAndChars" w:linePitch="{pitch}" w:charSpace="0"/>'
    else:  # lines
        docgrid = f'<w:docGrid w:type="lines" w:linePitch="{pitch}"/>'

    paras = []
    n = 0
    x = ind0
    while x <= ind1 + 1e-6:
        rt = int(round(x * 20))  # pt -> twips
        # tag each para with its index+indent so we can map data->config
        # No inline tag (it would add variable width and confound the break).
        # Each para = identical BASE; BASE starts with 甲 (unique to position 0),
        # so a PDF line starting with 甲 marks a paragraph start → the Nth such
        # paragraph maps to right-indent = ind0 + N*step (paragraphs are emitted
        # in sweep order).
        paras.append(
            f'<w:p><w:pPr><w:jc w:val="{jc}"/>'
            f'<w:ind w:right="{rt}"/>'
            f'<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
            f'<w:sz w:val="{sz}"/></w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
            f'<w:sz w:val="{sz}"/></w:rPr>'
            f'<w:t xml:space="preserve">{esc(BASE)}</w:t></w:r></w:p>'
        )
        n += 1
        x += step

    # A4 portrait, ~standard margins. Content width ~ 11906-1417*2 ... use a fixed
    # left/right margin; the per-para right-indent does the width sweep.
    sectpr = (
        f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        f'<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" '
        f'w:header="851" w:footer="992" w:gutter="0"/>'
        f'{docgrid}</w:sectPr>'
    )
    body = "".join(paras) + sectpr
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{body}</w:body></w:document>'
    )
    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:docDefaults><w:rPrDefault><w:rPr>'
        f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
        f'<w:sz w:val="{sz}"/></w:rPr></w:rPrDefault></w:docDefaults>'
        '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
        '<w:pPr><w:widowControl w:val="0"/></w:pPr></w:style>'
        '</w:styles>'
    )
    # compatibilityMode + kinsoku/justify-relevant settings left at Word defaults.
    # ★compressPunctuation: ALL the real FAIL docs (tokyoshugyo/nedocontract/
    # ikujidetail/kyotei) carry <w:characterSpacingControl w:val="compressPunctuation"/>.
    # WITHOUT it Word renders 約物 at full width (no compression/kern); WITH it
    # Word compresses 約物 — this is the setting that drives the char-budget wall.
    csc = "" if cpunct == "0" else \
        f'<w:characterSpacingControl w:val="compressPunctuation"/>'
    settings = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'{csc}'
        '<w:compat>'
        '<w:compatSetting w:name="compatibilityMode" '
        'w:uri="http://schemas.microsoft.com/office/word" '
        f'w:val="{compat}"/>'
        '</w:compat></w:settings>'
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
        '</Types>'
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>'
    )
    docrels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
        '</Relationships>'
    )
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", document)
        z.writestr("word/_rels/document.xml.rels", docrels)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", settings)
    print(f"wrote {out}: {n} paras, jc={jc} compat={compat} grid={grid} pitch={pitch} font={font} sz={sz} ind {ind0}-{ind1}/{step}pt")

if __name__ == "__main__":
    a = sys.argv
    def opt(name, dflt):
        return a[a.index(name)+1] if name in a else dflt
    out = a[1]
    build(out,
          jc=opt("--jc","both"), compat=opt("--compat","15"),
          grid=opt("--grid","lines"), pitch=opt("--pitch","360"),
          font=opt("--font","ＭＳ 明朝"), sz=opt("--sz","21"),
          ind0=float(opt("--ind0","0")), ind1=float(opt("--ind1","60")),
          step=float(opt("--step","0.5")), base=opt("--base","1"), cpunct=opt("--cpunct","1"))
