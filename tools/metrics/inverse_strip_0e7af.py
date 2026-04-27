"""Strategy B: inverse-strip 0e7af.docx to find the support file / element
that suppresses yakumono compression.

After strategies A1-A2 (single-axis tests, jc=both, jc=both+wrap) all
REFUTED the discriminator candidates, the remaining suspect is in the
docDefaults / styles / fontTable / theme support files of 0e7af.

Approach: produce 5 variants of 0e7af with progressively-stripped
docDefaults and support files. For each variant, measure ONE existing
yakumono pair in the body (using a 。 followed by ） position from the
probe data) to see if compression starts.

Variants:
  V1: strip rPrDefault lang attribute (eastAsia=ja-JP suspicion)
  V2: strip rPrDefault entirely (font + size + lang baseline)
  V3: strip pPrDefault entirely (jc=both + widowControl)
  V4: strip docDefaults entirely (whole block gone)
  V5: replace styles.xml with minimal content (one Normal style only)

Whichever V_n first shows compression on the existing 。） pair = that
strip removed the discriminator → element pinned.
"""
import os
import re
import shutil
import sys
import zipfile

SRC = 'tools/golden-test/documents/docx/0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx'
OUT_DIR = 'tools/metrics/inverse_strip_variants'

VARIANTS = [
    'V1_strip_lang',
    'V2_strip_rPrDefault',
    'V3_strip_pPrDefault',
    'V4_strip_docDefaults',
    'V5_minimal_styles',
]


def transform_styles(xml, variant):
    if variant == 'V1_strip_lang':
        # Remove just <w:lang .../> inside rPrDefault
        return re.sub(r'<w:lang\s+[^/]*?/>', '', xml, count=1)
    if variant == 'V2_strip_rPrDefault':
        return re.sub(r'<w:rPrDefault>.*?</w:rPrDefault>', '', xml, flags=re.DOTALL)
    if variant == 'V3_strip_pPrDefault':
        return re.sub(r'<w:pPrDefault>.*?</w:pPrDefault>', '', xml, flags=re.DOTALL)
    if variant == 'V4_strip_docDefaults':
        return re.sub(r'<w:docDefaults>.*?</w:docDefaults>', '', xml, flags=re.DOTALL)
    if variant == 'V5_minimal_styles':
        # Replace whole <w:styles>...</w:styles> body with minimal content
        return re.sub(
            r'(<w:styles\b[^>]*>).*?(</w:styles>)',
            r'\1<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>\2',
            xml,
            flags=re.DOTALL,
        )
    return xml


def make_variant(variant):
    out = os.path.join(OUT_DIR, f'0e7af_{variant}.docx')
    with zipfile.ZipFile(SRC, 'r') as zin:
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'word/styles.xml':
                    data = transform_styles(data.decode('utf-8'), variant).encode('utf-8')
                zout.writestr(item, data)
    return out


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    for v in VARIANTS:
        out = make_variant(v)
        print(f'wrote {out}', flush=True)
    # Sanity: show docDefaults state for each
    print('\n== sanity ==')
    for v in VARIANTS:
        path = os.path.join(OUT_DIR, f'0e7af_{v}.docx')
        with zipfile.ZipFile(path) as z:
            s = z.read('word/styles.xml').decode('utf-8')
        has_dd = '<w:docDefaults>' in s
        has_pdd = '<w:pPrDefault>' in s
        has_rdd = '<w:rPrDefault>' in s
        has_lang = '<w:lang' in s
        print(f'  {v}: docDefaults={has_dd}, pPrDefault={has_pdd}, rPrDefault={has_rdd}, lang={has_lang}')


if __name__ == '__main__':
    sys.exit(main())
