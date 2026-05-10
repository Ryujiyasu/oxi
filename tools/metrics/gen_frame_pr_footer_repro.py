"""Day 33 part 18 — Generate minimal repros for framePr footer height rule.

Hypothesis: Word excludes framePr-wrapped paragraphs from inline footer_h
(they are floating frames). Compare last-page-fitting paragraph y between:
  FP_A: footer with 1 empty paragraph
  FP_B: footer with framePr-page-number paragraph + 1 empty paragraph
  FP_C: footer with framePr-page-number paragraph only

If Word's max-content-fit y is identical across all three, hypothesis confirmed.
"""
from __future__ import annotations
import os, zipfile, shutil
from pathlib import Path

OUT = Path('tools/golden-test/repros/frame_pr_footer')
OUT.mkdir(parents=True, exist_ok=True)

NS = '''xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
 xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'''

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>'''

ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>'''

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="MS Mincho" w:hAnsi="MS Mincho" w:eastAsia="MS Mincho"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:styleId="a"><w:name w:val="Footer"/><w:rPr><w:sz w:val="18"/></w:rPr></w:style>
</w:styles>'''


def make_footer(variant):
    # Use BIG fs (sz=160 = 80pt) on framePr para so its inclusion would
    # push footer_h above bottom_margin floor (22.4pt threshold for our
    # docGrid 49.6pt footer_dist, 72pt bottom_margin). Distinguishes
    # "framePr counted" from "framePr skipped".
    # Variants:
    #   A: 1 empty para (baseline; footer_h ≈ 12pt < threshold → clamped)
    #   B: BIG framePr (80pt) + empty para (if framePr counted: footer_h
    #      ≈ 92pt > threshold → break shifts UP; if skipped: only empty
    #      ≈ 12pt → clamped, same as A)
    #   C: BIG framePr only (similar; if counted: 80pt; if skipped: 0pt)
    framepr_para = ('<w:p><w:pPr><w:pStyle w:val="a"/>'
                    '<w:framePr w:wrap="around" w:vAnchor="text" w:hAnchor="margin" w:xAlign="center" w:y="1"/>'
                    '<w:rPr><w:sz w:val="160"/></w:rPr>'
                    '</w:pPr><w:r><w:rPr><w:sz w:val="160"/></w:rPr><w:t>1</w:t></w:r></w:p>')
    empty_para = '<w:p><w:pPr><w:pStyle w:val="a"/></w:pPr></w:p>'
    if variant == 'A':
        body = empty_para
    elif variant == 'B':
        body = framepr_para + empty_para
    elif variant == 'C':
        body = framepr_para
    else:
        raise ValueError(variant)
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr {NS}>{body}</w:ftr>'


def make_document(n_paras):
    """Body: many short paragraphs to force end-of-page boundary detection."""
    paras = []
    for i in range(n_paras):
        paras.append(f'<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:t>P{i:03d}</w:t></w:r></w:p>')
    body_paras = ''.join(paras)
    sect = ('<w:sectPr>'
            '<w:footerReference w:type="default" r:id="rId2"/>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" w:header="851" w:footer="992" w:gutter="0"/>'
            '<w:docGrid w:type="lines" w:linePitch="312"/>'
            '</w:sectPr>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:document {NS}><w:body>{body_paras}{sect}</w:body></w:document>')


def build_docx(name, variant, n_paras):
    p = OUT / f'{name}.docx'
    if p.exists():
        p.unlink()
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/document.xml', make_document(n_paras))
        z.writestr('word/footer1.xml', make_footer(variant))
    print(f'  wrote {p}')


if __name__ == '__main__':
    # Use 50 short paragraphs at 15.5pt advance — total ~775pt body — exceeds page
    build_docx('FP_A_empty_only', 'A', 50)
    build_docx('FP_B_framepr_plus_empty', 'B', 50)
    build_docx('FP_C_framepr_only', 'C', 50)
    print('Done. Open each in Word to compare last-paragraph-on-page-1 y.')
