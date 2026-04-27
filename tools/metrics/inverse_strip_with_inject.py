"""Strategy B (corrected): inverse-strip 0e7af WITH fixture text injected.

Earlier inverse_strip_0e7af.py measured 、信 in para 34 — 信 is a CJK
ideograph, NOT a yakumono trigger, so no doc would compress that pair.
False-negative across all variants.

Corrected: take 0e7af, inject 観、「測 × 10 into a body paragraph (so we
have a measurable yakumono pair where the rule would be expected to
fire if it were enabled), THEN apply each docDefaults/styles strip
variant, then measure the injected 、 width.

If any variant flips from FULL (~11.5pt at inherited size) to
COMPRESSED (~half of that), the strip removed the discriminator.
"""
import os
import re
import sys
import zipfile

SRC = 'tools/golden-test/documents/docx/0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx'
OUT_DIR = 'tools/metrics/inverse_strip_inject_variants'

PATTERN = '観、「測' * 10  # 40 chars

VARIANTS = [
    'V0_baseline_inject',  # text inject only, no strip
    'V1_strip_lang',
    'V2_strip_rPrDefault',
    'V3_strip_pPrDefault',
    'V4_strip_docDefaults',
    'V5_minimal_styles',
]


def transform_styles(xml, variant):
    if variant == 'V1_strip_lang':
        return re.sub(r'<w:lang\s+[^/]*?/>', '', xml, count=1)
    if variant == 'V2_strip_rPrDefault':
        return re.sub(r'<w:rPrDefault>.*?</w:rPrDefault>', '', xml, flags=re.DOTALL)
    if variant == 'V3_strip_pPrDefault':
        return re.sub(r'<w:pPrDefault>.*?</w:pPrDefault>', '', xml, flags=re.DOTALL)
    if variant == 'V4_strip_docDefaults':
        return re.sub(r'<w:docDefaults>.*?</w:docDefaults>', '', xml, flags=re.DOTALL)
    if variant == 'V5_minimal_styles':
        return re.sub(
            r'(<w:styles\b[^>]*>).*?(</w:styles>)',
            r'\1<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>\2',
            xml,
            flags=re.DOTALL,
        )
    return xml


def inject_into_body_para(xml):
    """Replace text in the first plain body paragraph (no pStyle, no run sz)
    with PATTERN."""
    para_pattern = re.compile(r'<w:p\b[^>]*>.+?</w:p>', flags=re.DOTALL)
    out_parts = []
    last_end = 0
    replaced = False
    for m in para_pattern.finditer(xml):
        out_parts.append(xml[last_end:m.start()])
        body = m.group(0)
        if not replaced:
            ts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', body)
            joined = ''.join(ts)
            if joined.strip() and not re.search(r'<w:pStyle\b', body) and not re.search(r'<w:sz\s+w:val="(\d+)"', body):
                first_t_replaced = [False]
                def t_replacer(match):
                    if not first_t_replaced[0]:
                        first_t_replaced[0] = True
                        return f'<w:t xml:space="preserve">{PATTERN}</w:t>'
                    return '<w:t xml:space="preserve"></w:t>'
                body = re.sub(r'<w:t[^>]*>[^<]*</w:t>', t_replacer, body)
                replaced = True
        out_parts.append(body)
        last_end = m.end()
    out_parts.append(xml[last_end:])
    return ''.join(out_parts)


def make_variant(variant):
    out = os.path.join(OUT_DIR, f'0e7af_inject_{variant}.docx')
    with zipfile.ZipFile(SRC, 'r') as zin:
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'word/document.xml':
                    data = inject_into_body_para(data.decode('utf-8')).encode('utf-8')
                elif item.filename == 'word/styles.xml' and variant != 'V0_baseline_inject':
                    data = transform_styles(data.decode('utf-8'), variant).encode('utf-8')
                zout.writestr(item, data)
    return out


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    for v in VARIANTS:
        out = make_variant(v)
        print(f'wrote {out}', flush=True)


if __name__ == '__main__':
    sys.exit(main())
