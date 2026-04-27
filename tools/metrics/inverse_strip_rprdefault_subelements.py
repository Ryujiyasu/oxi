"""Drill down: which element WITHIN rPrDefault is the discriminator?

V1 (strip lang only) FAIL — gate not in lang.
V2 (strip whole rPrDefault) PASS — gate is in {rFonts, sz, szCs}.

Build 3 sub-variants stripping each element individually:
  V2a: strip rFonts (keep sz, szCs, lang)
  V2b: strip sz + szCs (keep rFonts, lang)
  V2c: strip rFonts + sz + szCs (keep only lang) — equivalent to "lang-only" rPrDefault

Whichever flips to COMPRESSED identifies the discriminator element.
"""
import os
import re
import sys
import zipfile

SRC = 'tools/golden-test/documents/docx/0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx'
OUT_DIR = 'tools/metrics/rprdefault_sub_variants'

PATTERN = '観、「測' * 10

VARIANTS = [
    'V2a_strip_rFonts',
    'V2b_strip_sz',
    'V2c_strip_rFonts_and_sz',
]


def transform_styles(xml, variant):
    if variant == 'V2a_strip_rFonts':
        return re.sub(r'<w:rFonts\s+[^/]*?/>', '', xml, count=1)
    if variant == 'V2b_strip_sz':
        # Strip sz and szCs (the first occurrences which are in rPrDefault)
        xml = re.sub(r'<w:sz\s+w:val="\d+"\s*/>', '', xml, count=1)
        xml = re.sub(r'<w:szCs\s+w:val="\d+"\s*/>', '', xml, count=1)
        return xml
    if variant == 'V2c_strip_rFonts_and_sz':
        xml = re.sub(r'<w:rFonts\s+[^/]*?/>', '', xml, count=1)
        xml = re.sub(r'<w:sz\s+w:val="\d+"\s*/>', '', xml, count=1)
        xml = re.sub(r'<w:szCs\s+w:val="\d+"\s*/>', '', xml, count=1)
        return xml
    return xml


def inject_into_body_para(xml):
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
                elif item.filename == 'word/styles.xml':
                    data = transform_styles(data.decode('utf-8'), variant).encode('utf-8')
                zout.writestr(item, data)
    return out


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    for v in VARIANTS:
        out = make_variant(v)
        print(f'wrote {out}', flush=True)
    # Sanity
    print('\n== sanity (rPrDefault content per variant) ==')
    for v in VARIANTS:
        path = os.path.join(OUT_DIR, f'0e7af_inject_{v}.docx')
        with zipfile.ZipFile(path) as z:
            s = z.read('word/styles.xml').decode('utf-8')
        m = re.search(r'<w:rPrDefault>.*?</w:rPrDefault>', s, flags=re.DOTALL)
        print(f'  {v}: rPrDefault = {m.group(0)[:200] if m else "(none)"}')


if __name__ == '__main__':
    sys.exit(main())
