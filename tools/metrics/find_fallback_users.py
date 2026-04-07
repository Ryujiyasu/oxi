"""Find 177-doc files where the eastAsia docDefaults fallback actually fires.

Criteria: docDefaults rPrDefault has eastAsia/eastAsiaTheme set, AND no run in
the document overrides it (no w:rFonts with w:eastAsia or w:asciiTheme/eastAsiaTheme
referencing East Asian content).
"""
import zipfile, re, glob, sys

sys.stdout.reconfigure(encoding='utf-8')


def doc_uses_fallback(docx_path):
    """Return True if some run/para fragments would rely on docDefaults eastAsia."""
    with zipfile.ZipFile(docx_path) as z:
        styles = z.read('word/styles.xml').decode('utf-8', 'ignore')
        doc = z.read('word/document.xml').decode('utf-8', 'ignore')

    # Has docDefaults eastAsia?
    rprd = re.search(r'<w:rPrDefault>.*?</w:rPrDefault>', styles, re.DOTALL)
    if not rprd:
        return False
    has_dd_ea = 'eastAsia=' in rprd.group(0) or 'eastAsiaTheme' in rprd.group(0)
    if not has_dd_ea:
        return False

    # Look at all run rFonts elements in document.xml.
    # Real-effect criteria: run has NO eastAsia/eastAsiaTheme AND run's ascii
    # (or asciiTheme) does NOT already point to a Japanese font. If ascii uses
    # a CJK theme like asciiTheme="minorEastAsia", the ascii path already
    # delivers MS Mincho — fallback is a no-op visually.
    rfonts_in_doc = re.findall(r'<w:rFonts[^/>]*/?>', doc)
    runs_total = len(rfonts_in_doc)
    real_fallback_runs = 0
    for r in rfonts_in_doc:
        has_ea = 'w:eastAsia=' in r or 'w:eastAsiaTheme=' in r
        if has_ea:
            continue
        ascii_uses_ea_theme = 'asciiTheme="minorEastAsia"' in r or 'asciiTheme="majorEastAsia"' in r
        if ascii_uses_ea_theme:
            continue
        ascii_match = re.search(r'w:ascii="([^"]*)"', r)
        if ascii_match:
            ascii_val = ascii_match.group(1)
            # Skip if ascii is already a Japanese font
            if any(c >= '\u3000' for c in ascii_val):
                continue
        real_fallback_runs += 1
    return {
        'has_dd_ea': True,
        'runs_total': runs_total,
        'real_fallback_runs': real_fallback_runs,
        'fallback_pct': 100 * real_fallback_runs / max(runs_total, 1),
    }


def main():
    files = sorted(glob.glob('tools/golden-test/documents/docx/*.docx'))
    candidates = []
    for f in files:
        try:
            r = doc_uses_fallback(f)
        except Exception:
            continue
        if not r:
            continue
        if r['fallback_pct'] > 30 and r['runs_total'] > 5:
            candidates.append((f, r))

    candidates.sort(key=lambda x: -x[1]['fallback_pct'])
    print(f"Found {len(candidates)} docs where docDefaults fallback fires for >50% of runs:\n")
    for f, r in candidates[:30]:
        name = f.split('\\')[-1].split('/')[-1]
        print(f"  {r['fallback_pct']:5.1f}%  ({r['real_fallback_runs']:4d}/{r['runs_total']:4d})  {name}")


if __name__ == '__main__':
    main()
