# -*- coding: utf-8 -*-
"""Char-budget CELL dataset generator (sibling of _cb_gen.py for the BODY).

tokyoshugyo's 条文/解説 boxes are 約物-rich CJK text in a SINGLE-CELL table,
jc=justify, compat=11 (legacy), docGrid type=lines, compressPunctuation. The CELL
wrapper is a SEPARATE code path from break_into_lines and does ZERO 約物 compression
where Word compresses — the documented cell char-budget wall.

This builds N single-cell tables, each holding the SAME 約物-rich sentence at a swept
cell WIDTH (gridCol/tcW), so across the sweep Word makes many oikomi-vs-oidashi cell
decisions on identical content → the controlled input to DERIVE Word's per-line CELL
約物 compression rule (then port to Oxi's cell wrapper, scoped via the S643 commentary
discriminator).

Each cell content starts with 甲 (unique marker) so a PDF line starting with 甲 is a
cell start; the Nth such cell maps to width = w0 + N*step (cells emitted in sweep order).

Usage:
  python _cb_gen_cell.py OUT.docx [--jc both|left] [--compat 11|14|15]
        [--pitch 360] [--font "ＭＳ 明朝"] [--sz 21]
        [--w0 6000] [--w1 8500] [--step 20]   (cell width in TWIPS)
        [--cellmar 108] [--base 1] [--cpunct 1]
"""
import sys, zipfile

# Same 約物-rich bases as _cb_gen.py (regulation style: many 、。 + varied trailers).
BASE1 = (
    "甲は、本契約に基づき、乙に対して、本件業務を、善良なる管理者の注意をもって、"
    "誠実かつ適切に遂行するものとし、これに関連する一切の責任を負う。"
    "なお、前項の規定にかかわらず、特別の事情がある場合には、別途協議のうえ、"
    "これを定めるものとする。"
)
BASE2 = (
    "甲本件業務遂行管理責任体制整備運用状況、定期報告書類作成提出義務履行確認、"
    "関係法令遵守状況点検記録保存、業務改善計画策定実施結果評価分析、"
    "年度末決算処理業務完了報告、次年度事業計画立案承認手続実施。"
)
BASE3 = (
    "甲は、「本件業務」、「関連業務」、及び「付随業務」を、乙に対して、"
    "「別紙一」のとおり、委託するものとし、乙は、「これら」を、"
    "「善良なる管理者」の注意をもって、誠実に遂行する。"
)
def _pick(b): return {"1": BASE1, "2": BASE2, "3": BASE3}.get(str(b), BASE1)

def esc(s): return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def build(out, jc="both", compat="11", pitch="360", font="ＭＳ 明朝", sz="21",
          w0=6000.0, w1=8500.0, step=20.0, cellmar="108", base="1", cpunct="1"):
    BASE = _pick(base)
    rpr = (f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
           f'<w:sz w:val="{sz}"/>')
    border = ('<w:tblBorders>'
              + ''.join(f'<w:{e} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                        for e in ("top", "left", "bottom", "right", "insideH", "insideV"))
              + '</w:tblBorders>')
    tcmar = (f'<w:tblCellMar>'
             f'<w:top w:w="0" w:type="dxa"/><w:left w:w="{cellmar}" w:type="dxa"/>'
             f'<w:bottom w:w="0" w:type="dxa"/><w:right w:w="{cellmar}" w:type="dxa"/>'
             f'</w:tblCellMar>')
    tbls = []
    n = 0
    w = w0
    while w <= w1 + 1e-6:
        cw = int(round(w))
        tbls.append(
            f'<w:tbl><w:tblPr><w:tblW w:w="{cw}" w:type="dxa"/>'
            f'<w:tblLayout w:type="fixed"/>{border}{tcmar}</w:tblPr>'
            f'<w:tblGrid><w:gridCol w:w="{cw}"/></w:tblGrid>'
            f'<w:tr><w:tc><w:tcPr><w:tcW w:w="{cw}" w:type="dxa"/></w:tcPr>'
            f'<w:p><w:pPr><w:jc w:val="{jc}"/><w:rPr>{rpr}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{esc(BASE)}</w:t></w:r>'
            f'</w:p></w:tc></w:tr></w:tbl>'
            # spacer empty para between tables so they don't merge
            f'<w:p><w:pPr><w:rPr>{rpr}</w:rPr></w:pPr></w:p>'
        )
        n += 1
        w += step

    docgrid = f'<w:docGrid w:type="lines" w:linePitch="{pitch}"/>'
    sectpr = (f'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
              f'<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" '
              f'w:header="851" w:footer="992" w:gutter="0"/>{docgrid}</w:sectPr>')
    document = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                f'<w:body>{"".join(tbls)}{sectpr}</w:body></w:document>')
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
              '<w:docDefaults><w:rPrDefault><w:rPr>'
              f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>'
              f'<w:sz w:val="{sz}"/></w:rPr></w:rPrDefault></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
              '<w:pPr><w:widowControl w:val="0"/></w:pPr></w:style></w:styles>')
    csc = "" if cpunct == "0" else '<w:characterSpacingControl w:val="compressPunctuation"/>'
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                f'{csc}<w:compat><w:compatSetting w:name="compatibilityMode" '
                'w:uri="http://schemas.microsoft.com/office/word" '
                f'w:val="{compat}"/></w:compat></w:settings>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    docrels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
               '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>')
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", document)
        z.writestr("word/_rels/document.xml.rels", docrels)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", settings)
    print(f"wrote {out}: {n} cells, jc={jc} compat={compat} pitch={pitch} font={font} "
          f"sz={sz} cellw {w0}-{w1}/{step}tw cellmar={cellmar} base={base}")


if __name__ == "__main__":
    a = sys.argv
    def opt(name, d): return a[a.index(name)+1] if name in a else d
    build(a[1], jc=opt("--jc", "both"), compat=opt("--compat", "11"),
          pitch=opt("--pitch", "360"), font=opt("--font", "ＭＳ 明朝"), sz=opt("--sz", "21"),
          w0=float(opt("--w0", "6000")), w1=float(opt("--w1", "8500")),
          step=float(opt("--step", "20")), cellmar=opt("--cellmar", "108"),
          base=opt("--base", "1"), cpunct=opt("--cpunct", "1"))
