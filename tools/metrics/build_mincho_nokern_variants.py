"""Build MS Mincho variants without kerning to test the kerning-gate hypothesis.

Following 2026-04-27 yakumono always-on falsification + 0e7af/683ff COM
probe finding: MS Mincho fixtures with `<w:kern>` compress yakumono pairs,
but 0e7af/683ff (MS Mincho, NO kerning) do NOT compress.

Hypothesis: For MS Mincho, `<w:kern>` is the gate that enables yakumono
compression. (Meiryo compresses regardless per LW_30/LW_31 finding —
font-specific behavior.)

This script takes MC_A_mincho.docx and produces 2 variants:
- MC_A_mincho_NOKERN: identical to MC_A_mincho but with all <w:kern>
  elements stripped from runs (paragraphs)
- MC_A_mincho_NOKERN_COMPAT15: NOKERN + compatibilityMode=15 (matches 0e7af)

Run: python tools/metrics/build_mincho_nokern_variants.py
"""
import os
import re
import shutil
import sys
import zipfile

SRC = 'tools/metrics/mincho_adjacency_repro/MC_A_mincho.docx'
OUT_DIR = 'tools/metrics/mincho_kern_variants'


def write_variant(name, modify_doc_xml=lambda x: x, modify_settings_xml=lambda x: x):
    out_path = os.path.join(OUT_DIR, f'{name}.docx')
    with zipfile.ZipFile(SRC, 'r') as zin:
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'word/document.xml':
                    data = modify_doc_xml(data.decode('utf-8')).encode('utf-8')
                elif item.filename == 'word/settings.xml':
                    data = modify_settings_xml(data.decode('utf-8')).encode('utf-8')
                zout.writestr(item, data)
    print(f'wrote {out_path}', flush=True)


def strip_kern(xml):
    return re.sub(r'<w:kern\s+w:val="\d+"\s*/>', '', xml)


def set_compat15(xml):
    return re.sub(
        r'(<w:compatSetting\s+w:name="compatibilityMode"[^>]*?w:val=")14(")',
        r'\g<1>15\g<2>',
        xml,
    )


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    print(f'src: {SRC}', flush=True)

    # Variant 1: NOKERN (kern stripped, everything else same as MC_A_mincho)
    write_variant('MC_A_mincho_NOKERN',
                  modify_doc_xml=strip_kern)

    # Variant 2: NOKERN + COMPAT15 (matches 0e7af compat setting)
    write_variant('MC_A_mincho_NOKERN_COMPAT15',
                  modify_doc_xml=strip_kern,
                  modify_settings_xml=set_compat15)

    # Sanity-check: read back kern + compat from both variants
    print('\n== sanity ==', flush=True)
    for name in ['MC_A_mincho_NOKERN', 'MC_A_mincho_NOKERN_COMPAT15']:
        path = os.path.join(OUT_DIR, f'{name}.docx')
        with zipfile.ZipFile(path) as z:
            d = z.read('word/document.xml').decode('utf-8')
            s = z.read('word/settings.xml').decode('utf-8')
        kern_count = len(re.findall(r'<w:kern\s+w:val="\d+"\s*/>', d))
        compat = re.search(r'compatibilityMode"[^>]*?w:val="(\d+)"', s)
        print(f'  {name}: kern_count={kern_count}, compat={compat.group(1) if compat else "?"}')


if __name__ == '__main__':
    sys.exit(main())
