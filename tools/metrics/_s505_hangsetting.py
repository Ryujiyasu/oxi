# -*- coding: utf-8 -*-
"""S505 hang-gate discriminator: compare b837 (Word OIDASHI, no hang) vs e3c545 (Word HANGS,
S492 regressed if hang off) settings.xml + sectPr for hanging-punctuation / kinsoku / compat
flags. cp932-safe. Find the clean discriminator to gate can_hang."""
import io, zipfile, re, glob, os
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))


def dx_for(pref):
    for d in ['tools/golden-test/documents/docx', 'pipeline_data/docx']:
        g = glob.glob(os.path.join(ROOT, d, pref + '*.docx'))
        if g:
            return g[0]
    return None


FLAGS = ['hangingPunctuation', 'overflowPunct', 'noLineWrap', 'compressPunctuation',
         'doNotExpandShiftReturn', 'autoSpaceDE', 'autoSpaceDN', 'kinsoku', 'wordWrap',
         'doNotUseHTMLParagraphAutoSpacing', 'compatibilityMode', 'useWord2013TrackBottomHyphenation']
L = []
for pref, label in [('b837808d0555', 'b837 (OIDASHI/no-hang)'), ('e3c545fac7a7', 'e3c545 (HANGS)'),
                    ('db9ca18368cd', 'db9ca'), ('15076df085f5', '15076')]:
    dx = dx_for(pref)
    if not dx:
        L.append('%s: MISSING' % label); continue
    z = zipfile.ZipFile(dx)
    L.append('\n=== %s  %s ===' % (label, os.path.basename(dx)))
    try:
        s = z.read('word/settings.xml').decode('utf-8')
    except Exception:
        s = ''
    for fl in FLAGS:
        m = re.search(r'<w:%s[ />][^>]*>?' % fl, s)
        if m:
            L.append('  settings %s: %s' % (fl, m.group(0)[:80]))
    # compatSetting compatibilityMode value
    cm = re.search(r'compatibilityMode"[^/]*w:val="(\d+)"', s)
    L.append('  compatibilityMode = %s' % (cm.group(1) if cm else '?(default 12)'))
    doc = z.read('word/document.xml').decode('utf-8')
    dg = re.search(r'<w:docGrid [^>]*>', doc); L.append('  docGrid: %s' % (dg.group(0) if dg else 'none'))
    # paragraph-level kinsoku / overflowPunct in sectPr or pPr
    for fl in ['overflowPunct', 'autoSpaceDE', 'kinsoku', 'wordWrap', 'topLinePunct']:
        if ('<w:%s' % fl) in doc:
            L.append('  document has <w:%s>' % fl)
with io.open('c:/tmp/_s505_hang_out.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(L) + '\n')
print('wrote c:/tmp/_s505_hang_out.txt')
