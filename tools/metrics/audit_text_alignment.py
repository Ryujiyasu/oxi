"""Day 32 part 8 — pPrDefault textAlignment audit.

Day 32 part 7 hypothesis: Bug 2 conditional detector = pPrDefault
textAlignment="baseline"|"top" (mod.rs:5402-5404 already handles this
case → returns 0.0 offset).

This tool extracts textAlignment settings from styles.xml + settings.xml
+ first-paragraph pPr for each Class A doc + preserve-class samples to
test whether textAlignment discriminates the two classes.

If Class A docs all lack textAlignment="baseline"|"top" and preserve
docs all have it, the detector hypothesis is confirmed.
"""
from __future__ import annotations
import os, sys, zipfile, re
sys.stdout.reconfigure(encoding='utf-8')

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DOCX_DIR = os.path.join(REPO, 'tools', 'golden-test', 'documents', 'docx')


def find_docx(doc_id):
    for f in os.listdir(DOCX_DIR):
        if f.startswith(doc_id) and f.endswith('.docx'):
            return os.path.join(DOCX_DIR, f)
    return None


def get_xml(zf, name):
    try:
        return zf.read(name).decode('utf-8')
    except KeyError:
        return ''


def find_text_alignment(xml, label):
    """Find <w:textAlignment w:val="..."/>"""
    if not xml:
        return None
    matches = re.findall(r'<w:textAlignment[^>]*w:val="([^"]+)"', xml)
    return matches[0] if matches else None


def find_doc_default_alignment(styles_xml):
    """Find textAlignment in <w:docDefaults><w:pPrDefault>."""
    if not styles_xml:
        return None
    m = re.search(r'<w:docDefaults>.*?</w:docDefaults>', styles_xml, re.DOTALL)
    if not m:
        return None
    return find_text_alignment(m.group(0), 'docDefaults')


def find_normal_style_alignment(styles_xml):
    """Find textAlignment in Normal style."""
    if not styles_xml:
        return None
    m = re.search(r'<w:style[^>]+w:styleId="(?:Normal|a)"[^>]*>.*?</w:style>', styles_xml, re.DOTALL)
    if not m:
        return None
    return find_text_alignment(m.group(0), 'normal style')


def find_first_paragraph_alignment(document_xml):
    """Find textAlignment in first paragraph's pPr."""
    if not document_xml:
        return None
    m = re.search(r'<w:p[^>]*>.*?</w:p>', document_xml, re.DOTALL)
    if not m:
        return None
    return find_text_alignment(m.group(0), 'first para')


def find_compat_settings(settings_xml):
    """Find compatibility flags that may affect line positioning."""
    flags = {}
    if not settings_xml:
        return flags
    for tag in ['adjustLineHeightInTable', 'noLineBreaksAfter', 'noLineBreaksBefore',
                'compatSetting', 'doNotExpandShiftReturn', 'autoSpaceLikeWord95',
                'doNotUseHTMLParagraphAutoSpacing', 'doNotBreakWrappedTables']:
        if f'<w:{tag}' in settings_xml:
            flags[tag] = True
    # extract compatSetting values
    compat = re.findall(r'<w:compatSetting w:name="([^"]+)" w:uri="[^"]*" w:val="([^"]+)"/>', settings_xml)
    for name, val in compat:
        flags[f'compat:{name}'] = val
    return flags


def audit(doc_id, label):
    docx = find_docx(doc_id)
    if not docx:
        print(f'{label} {doc_id}: NOT FOUND')
        return None
    with zipfile.ZipFile(docx) as zf:
        styles = get_xml(zf, 'word/styles.xml')
        settings = get_xml(zf, 'word/settings.xml')
        document = get_xml(zf, 'word/document.xml')
    doc_default = find_doc_default_alignment(styles)
    normal = find_normal_style_alignment(styles)
    first_para = find_first_paragraph_alignment(document)
    flags = find_compat_settings(settings)

    print(f'{label:<10} {doc_id}:')
    print(f'  docDefault   textAlignment: {doc_default!r}')
    print(f'  Normal style textAlignment: {normal!r}')
    print(f'  First-para   textAlignment: {first_para!r}')
    print(f'  adjustLineHeightInTable: {flags.get("adjustLineHeightInTable", False)}')
    print(f'  compatSettings: {[(k, v) for k, v in flags.items() if k.startswith("compat:")][:3]}')
    return {'doc_id': doc_id, 'doc_default': doc_default, 'normal': normal,
            'first_para': first_para, 'flags': flags}


def main():
    print('=== Class A docs (Bug 2 fires wrongly) ===')
    class_a = ['bd90b00ab7a7', 'de6e32b5960b', 'db9ca18368cd', 'd77a58485f16']
    a_results = []
    for d in class_a:
        r = audit(d, 'Class A')
        if r:
            a_results.append(r)
        print()

    print('=== Preserve-class sample (Bug 2 correctly applied) ===')
    preserve = ['e3c545fac7a7', '0e7af1ae8f21', 'b5f706bda1f8', '6515f6a8d65b']
    p_results = []
    for d in preserve:
        r = audit(d, 'Preserve')
        if r:
            p_results.append(r)
        print()

    print('=== Discrimination summary ===')
    a_ta = set(r['doc_default'] for r in a_results)
    p_ta = set(r['doc_default'] for r in p_results)
    print(f'Class A docDefault textAlignment values: {a_ta}')
    print(f'Preserve docDefault textAlignment values: {p_ta}')
    if a_ta & p_ta:
        print('  → docDefault textAlignment does NOT discriminate (overlap)')
    else:
        print('  → docDefault textAlignment DOES discriminate!')


if __name__ == '__main__':
    main()
