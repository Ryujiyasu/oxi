"""Add one d77a style property back to S6 at a time — which brings back compression?"""
import os, re, zipfile

SRC = os.path.abspath(r"pipeline_data/d77a_S1_only_para28.docx")
OUT_DIR = os.path.abspath(r"pipeline_data")


def rewrite(src, dst, transforms):
    with zipfile.ZipFile(src, 'r') as zin:
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)
                if item in transforms:
                    data = transforms[item](data)
                zout.writestr(item, data)


# S6 minimal styles that loses compression
def make_styles(include_kern=False, include_jc_both=False, include_widow=False,
                include_lang=False, include_normal_rfonts=False,
                include_szCs=False):
    rpr_items = []
    if include_normal_rfonts:
        rpr_items.append('<w:rFonts w:ascii="Century" w:eastAsia="ＭＳ 明朝" w:hAnsi="Century" w:cs="Times New Roman"/>')
    if include_lang:
        rpr_items.append('<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>')
    doc_default_rpr = ''.join(rpr_items)

    normal_ppr_items = []
    if include_widow: normal_ppr_items.append('<w:widowControl w:val="0"/>')
    if include_jc_both: normal_ppr_items.append('<w:jc w:val="both"/>')
    normal_ppr = f'<w:pPr>{"".join(normal_ppr_items)}</w:pPr>' if normal_ppr_items else ''

    normal_rpr_items = []
    if include_kern: normal_rpr_items.append('<w:kern w:val="2"/>')
    normal_rpr_items.append('<w:sz w:val="21"/>')
    if include_szCs: normal_rpr_items.append('<w:szCs w:val="24"/>')
    normal_rpr = f'<w:rPr>{"".join(normal_rpr_items)}</w:rPr>'

    # Avoid empty rPr — Word dislikes it
    doc_default_block = f'<w:rPrDefault><w:rPr>{doc_default_rpr}</w:rPr></w:rPrDefault>' if doc_default_rpr else ''
    styles_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>{doc_default_block}<w:pPrDefault/></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>{normal_ppr}{normal_rpr}</w:style>
</w:styles>'''
    return styles_xml


VARIANTS = [
    # R1: baseline (same as S6, minimal)
    ("R1_baseline", {}),
    # R2: +kern=2
    ("R2_kern", {"include_kern": True}),
    # R3: +jc=both
    ("R3_jc_both", {"include_jc_both": True}),
    # R4: +widowControl=0
    ("R4_widow", {"include_widow": True}),
    # R5: +docDefaults lang
    ("R5_lang", {"include_lang": True}),
    # R6: +docDefaults rFonts (Century/MS Mincho)
    ("R6_docdefault_rfonts", {"include_normal_rfonts": True}),
    # R7: +szCs=24
    ("R7_szCs", {"include_szCs": True}),
    # R_ALL: all together
    ("R_ALL", {"include_kern": True, "include_jc_both": True, "include_widow": True,
              "include_lang": True, "include_normal_rfonts": True, "include_szCs": True}),
]


def main():
    for label, flags in VARIANTS:
        styles_xml = make_styles(**flags)
        out = os.path.join(OUT_DIR, f"d77a_{label}.docx")
        rewrite(SRC, out, {"word/styles.xml": lambda _, x=styles_xml: x.encode("utf-8")})
        print(f"[{label}] {out}")


if __name__ == "__main__":
    main()
