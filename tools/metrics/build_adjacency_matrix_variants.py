"""Adjacency matrix variant builder — disambiguate yakumono compression trigger.

Existing baseline (`adjacency_matrix_repro/`, measured 2026-04-25) shows Word
applies 50% closing-punct compression for Meiryo 10.5pt with:
  useFELayout=ON, cSC=doNotCompress, compat=14, kern=3

Oxi's gate (`mod.rs:4140` `yakumono_enabled = compress_punctuation`) requires
cSC=compressPunctuation; with cSC=doNotCompress it disables yakumono entirely.

2026-04-27 update: V_NOFE was REMOVED after `meiryo_linewidth_repro.json`
LW_30 (useFE=ON,kern=3) vs LW_31 (useFE=off,kern=off) per-char compare showed
identical 5.50pt compression for `、「` pair → useFELayout/kern are NOT the
gate. Remaining open variants:

  V_CP        : cSC=compressPunctuation (useFELayout=ON, compat=14)
                — does cP enable additional compression beyond baseline?
  V_COMPAT15  : compat=15              (useFELayout=ON, cSC=doNotCompress)
                — does compat=15 differ from compat=14 on yakumono behaviour?

Run order (after this build script):
  1. python tools/metrics/build_adjacency_matrix_variants.py
  2. python tools/metrics/measure_adjacency_matrix_variants.py
  3. compare per-variant prev-width matrices: any variant matching baseline
     means that knob is NOT the gate. Variant differing pins the knob.

If both V_CP and V_COMPAT15 match baseline, the gate is "always-on for
compat>=14 + cSC in {doNotCompress, compressPunctuation, ...}" — i.e., Word
applies the next-trigger rule unconditionally, and Oxi's `mod.rs:4140`
gate is over-restrictive. See RESEARCH_LOG.md 2026-04-27.
"""
import os
import zipfile

CT = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'
RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
DOC_RELS = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'


def settings_xml(use_fe_layout: bool, csc_value: str, compat_mode: int) -> str:
    fe_tag = '<w:useFELayout/>' if use_fe_layout else ''
    return f'''<?xml version="1.0"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:compat>{fe_tag}<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compat_mode}"/></w:compat>
  <w:characterSpacingControl w:val="{csc_value}"/>
</w:settings>'''


def doc_xml(text: str) -> str:
    rpr = '<w:rFonts w:ascii="メイリオ" w:eastAsia="メイリオ" w:hAnsi="メイリオ"/><w:sz w:val="21"/><w:szCs w:val="21"/><w:kern w:val="3"/>'
    esc = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    return f'''<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:pPr><w:rPr>{rpr}</w:rPr></w:pPr><w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{esc}</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
</w:body></w:document>'''


def build_one(out_dir: str, label: str, text: str, settings: str) -> None:
    path = os.path.join(out_dir, f'{label}.docx')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/settings.xml', settings)
        z.writestr('word/document.xml', doc_xml(text))


PUNCTS = {
    'CM':  '、',
    'PD':  '。',
    'LBK': '「',
    'RBK': '」',
    'LPN': '（',
    'RPN': '）',
    'FPD': '．',
    'FCM': '，',
}


VARIANTS = [
    # (variant_id, use_fe_layout, csc_value, compat_mode, description)
    # V_NOFE removed 2026-04-27 — falsified by meiryo_linewidth LW_30/LW_31
    ('V_CP',       True,  'compressPunctuation', 14, 'baseline plus cSC=compressPunctuation'),
    ('V_COMPAT15', True,  'doNotCompress',       15, 'baseline with compat=15'),
]


def main():
    for variant_id, use_fe, csc, compat, desc in VARIANTS:
        out_dir = os.path.abspath(f"tools/metrics/adjacency_matrix_repro_{variant_id}")
        os.makedirs(out_dir, exist_ok=True)
        settings = settings_xml(use_fe, csc, compat)
        count = 0
        for prev_label, prev_ch in PUNCTS.items():
            for next_label, next_ch in PUNCTS.items():
                if prev_label == next_label and prev_ch in '、。．，':
                    continue
                label = f'ADJ_{prev_label}_{next_label}'
                text = ('観' + prev_ch + next_ch + '測') * 10
                build_one(out_dir, label, text, settings)
                count += 1
        print(f'[{variant_id}] built {count} repros - {desc} -> {out_dir}')


if __name__ == "__main__":
    main()
