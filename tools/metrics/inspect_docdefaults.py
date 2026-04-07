"""Inspect what font Oxi resolves for docDefaults eastAsia in regressed docs.

Builds a tiny Rust program output via cargo example, or alternatively reads
parser output via existing oxi binaries. For now, just compute parser
expectation manually from XML.
"""
import zipfile, re, sys, glob

sys.stdout.reconfigure(encoding='utf-8')


def get_doc_defaults_ea(docx_path):
    """Return (raw_attr, resolved_via_theme) for docDefaults eastAsia."""
    with zipfile.ZipFile(docx_path) as z:
        styles_xml = z.read('word/styles.xml').decode('utf-8', 'ignore')
        try:
            theme_xml = z.read('word/theme/theme1.xml').decode('utf-8', 'ignore')
        except KeyError:
            theme_xml = ''

    # Find rPrDefault rFonts
    m = re.search(r'<w:rPrDefault>.*?</w:rPrDefault>', styles_xml, re.DOTALL)
    if not m:
        return None, None
    rfonts = re.search(r'<w:rFonts[^/>]*/?>', m.group(0))
    if not rfonts:
        return None, None
    attrs = dict(re.findall(r'w:(\w+)="([^"]*)"', rfonts.group(0)))
    raw_ea = attrs.get('eastAsia')
    raw_theme = attrs.get('eastAsiaTheme')

    resolved = None
    if raw_ea:
        resolved = raw_ea
    elif raw_theme:
        # Resolve via theme minorFont script="Jpan"
        section = 'minorFont' if 'minor' in raw_theme else 'majorFont'
        sm = re.search(rf'<a:{section}>.*?</a:{section}>', theme_xml, re.DOTALL)
        if sm:
            jp = re.search(r'<a:font script="Jpan"[^/>]*?typeface="([^"]+)"', sm.group(0))
            if jp:
                resolved = jp.group(1)
            else:
                lat = re.search(r'<a:latin[^/>]*?typeface="([^"]+)"', sm.group(0))
                if lat:
                    resolved = lat.group(1)

    return (raw_ea or raw_theme), resolved


def main():
    docx_dir = 'tools/golden-test/documents/docx'
    files = sorted(glob.glob(f'{docx_dir}/*.docx'))
    counts = {}  # resolved -> count
    by_doc = {}
    for f in files:
        try:
            raw, resolved = get_doc_defaults_ea(f)
        except Exception as e:
            continue
        key = resolved or '(none)'
        counts[key] = counts.get(key, 0) + 1
        by_doc[f.split('/')[-1].split('\\')[-1]] = (raw, resolved)

    print('=== docDefaults eastAsia distribution ===')
    for k, v in sorted(counts.items(), key=lambda x: -x[1]):
        print(f'  {v:4d}  {k}')

    print(f'\nTotal docs: {len(files)}')
    print(f'\nSpecific regressed docs:')
    for d in ['04b88e7e0b25_index-19.docx', '4a36b62555f2_kyodokenkyuyoushiki10.docx']:
        if d in by_doc:
            print(f'  {d}: raw={by_doc[d][0]} → resolved={by_doc[d][1]}')


if __name__ == '__main__':
    main()
