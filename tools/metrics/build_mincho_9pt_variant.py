"""Build MS Mincho 9pt variant of MC_A_mincho to test size-as-discriminator.

Hypothesis (after kerning REFUTED): Word's yakumono compression depends
on font size. 0e7af = 9pt MS Mincho (no compression observed); MC_A_mincho
fixture = 10.5pt MS Mincho (compression observed). Build 9pt variant to
test if the threshold is at/below 9pt.
"""
import os
import re
import sys
import zipfile

SRC = 'tools/metrics/mincho_adjacency_repro/MC_A_mincho.docx'
OUT_DIR = 'tools/metrics/mincho_size_variants'


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    # Variant: 9pt = sz val "18" (half-points)
    # Original has sz val "21" = 10.5pt
    out_path = os.path.join(OUT_DIR, 'MC_A_mincho_9pt.docx')
    with zipfile.ZipFile(SRC, 'r') as zin:
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'word/document.xml':
                    xml = data.decode('utf-8')
                    # Replace all sz val="21" with sz val="18"
                    new_xml = re.sub(r'(<w:sz(?:Cs)?\s+w:val=")21(")', r'\g<1>18\g<2>', xml)
                    data = new_xml.encode('utf-8')
                zout.writestr(item, data)
    print(f'wrote {out_path}', flush=True)
    # Sanity
    with zipfile.ZipFile(out_path) as z:
        d = z.read('word/document.xml').decode('utf-8')
    sizes = re.findall(r'<w:sz\s+w:val="(\d+)"', d)
    print(f'sizes in variant: {sorted(set(sizes))}')


if __name__ == '__main__':
    sys.exit(main())
