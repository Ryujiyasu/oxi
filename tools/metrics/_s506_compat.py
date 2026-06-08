# -*- coding: utf-8 -*-
"""S506 compat check: read the actual compatibilityMode from settings.xml for the hang/oidashi
docs (correct regex; the _s505 one stopped at the URI slash). cp932-safe."""
import io, zipfile, re, glob, os
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
docs = [('b837808d0555', 'OIDASHI'), ('e3c545fac7a7', 'HANG'), ('683ffcab86e2', 'oidashi?'),
        ('0e7af1ae8f21', 'oidashi?'), ('d77a58485f16', 'oidashi?'), ('db9ca18368cd', '?'),
        ('15076df085f5', '?'), ('1ec1091177b1', '?')]
L = []
for pref, exp in docs:
    g = (glob.glob(os.path.join(ROOT, 'tools/golden-test/documents/docx', pref + '*.docx')) or
         glob.glob(os.path.join(ROOT, 'pipeline_data/docx', pref + '*.docx')))
    if not g:
        L.append('%s : MISSING' % pref); continue
    try:
        s = zipfile.ZipFile(g[0]).read('word/settings.xml').decode('utf-8')
    except Exception as e:
        L.append('%s : settings err %s' % (pref, e)); continue
    m = re.search(r'w:name="compatibilityMode"[^>]*?w:val="(\d+)"', s)
    if not m:
        m = re.search(r'w:val="(\d+)"[^>]*?w:name="compatibilityMode"', s)
    cm = m.group(1) if m else 'ABSENT(default 12)'
    L.append('%-16s compat=%-18s (expected %s)' % (pref, cm, exp))
with io.open('c:/tmp/_s506_compat_out.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(L) + '\n')
print('\n'.join(L))
