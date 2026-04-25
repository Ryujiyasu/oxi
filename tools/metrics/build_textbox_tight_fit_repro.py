"""Minimal repro: tight-fit single-line floating textbox.

A floating textbox where height = inset_t + line_height + inset_b exactly.
The OLD filter (`pe.y + pe.height > clip_bottom`) drops the only line.
The NEW filter (line-count-aware) keeps it.

Variants:
  TF_A: 11pt MS Mincho text, tight-fit (height 27.2pt = 3.6 + 20 + 3.6)
  TF_B: 11pt 様式１ exact replica (matches 459f05)
  TF_C: 14pt larger font tight-fit (height ≈ 30pt)
  TF_D: 2-line content in same height (overflow case — line 2 should drop)

Each is a single-page docx with anchor textbox at top-left.
"""
import os
import zipfile

OUT_DIR = os.path.abspath("tools/metrics/textbox_tight_fit_repro")
os.makedirs(OUT_DIR, exist_ok=True)

CT = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'


def doc_xml(text: str, height_emu: int, font: str, size_halfpt: int) -> str:
    """Build doc with floating textbox + body paragraph."""
    width_emu = 966158  # 76pt — reasonable for short text
    posh_emu = -450964  # -47.5pt left of column
    posv_emu = -417195  # -43.8pt above paragraph
    return f'''<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
<w:body>
<w:p>
<w:pPr><w:jc w:val="center"/></w:pPr>
<w:r>
<w:rPr><w:noProof/></w:rPr>
<w:drawing>
<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
<wp:simplePos x="0" y="0"/>
<wp:positionH relativeFrom="column"><wp:posOffset>{posh_emu}</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="paragraph"><wp:posOffset>{posv_emu}</wp:posOffset></wp:positionV>
<wp:extent cx="{width_emu}" cy="{height_emu}"/>
<wp:effectExtent l="0" t="0" r="0" b="0"/>
<wp:wrapNone/>
<wp:docPr id="15" name="Text Box"/>
<wp:cNvGraphicFramePr/>
<a:graphic>
<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
<wps:wsp>
<wps:cNvSpPr txBox="1"/>
<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{width_emu}" cy="{height_emu}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>
<wps:txbx>
<w:txbxContent>
<w:p>
<w:pPr><w:jc w:val="center"/></w:pPr>
<w:r>
<w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/><w:sz w:val="{size_halfpt}"/></w:rPr>
<w:t>{text}</w:t>
</w:r>
</w:p>
</w:txbxContent>
</wps:txbx>
<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0" upright="1"><a:spAutoFit/></wps:bodyPr>
</wps:wsp>
</a:graphicData>
</a:graphic>
</wp:anchor>
</w:drawing>
</w:r>
<w:r><w:rPr><w:rFonts w:ascii="ＭＳ 明朝" w:eastAsia="ＭＳ 明朝" w:hAnsi="ＭＳ 明朝"/><w:sz w:val="24"/></w:rPr><w:t>表題</w:t></w:r>
</w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
</w:body></w:document>'''


def build(label: str, text: str, height_emu: int, font='ＭＳ 明朝', size_halfpt=22):
    path = os.path.join(OUT_DIR, f'{label}.docx')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/document.xml', doc_xml(text, height_emu, font, size_halfpt))
    print(f"Built {path} text={text!r} height_pt={height_emu/12700:.2f} font={font} size={size_halfpt/2}pt")


# 11pt = 22 halfpt; 11pt × 1.15 ≈ 12.7pt; default lineHeight may be ~13-14pt
# Word's default line spacing for 11pt MS Mincho ≈ 17-20pt
# tight-fit: textbox height = inset_t (3.6) + line_height (~17-20pt) + inset_b (3.6) ≈ 24-27pt
# 27.2pt = 345440 EMU (samushiki1 case)
build('TF_A', '様式１', 345440, font='ＭＳ 明朝', size_halfpt=22)  # 11pt
build('TF_B', '見本', 345440, font='ＭＳ 明朝', size_halfpt=22)
build('TF_C', '注意', 380000, font='ＭＳ 明朝', size_halfpt=28)  # 14pt
build('TF_D', '2 line text overflow case', 345440, font='ＭＳ 明朝', size_halfpt=22)  # tight, but 2 lines wouldn't fit
