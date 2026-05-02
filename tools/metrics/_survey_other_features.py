# -*- coding: utf-8 -*-
"""Survey baseline for other untouched OOXML features."""
import sys, os, glob, zipfile, re
from collections import Counter
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

DOCX_DIR = r'C:\Users\ryuji\oxi-4\tools\golden-test\documents\docx'

features = {
    'sym': 0,                  # <w:sym> symbol char
    'softHyphen': 0,           # <w:softHyphen/>
    'noBreakHyphen': 0,        # <w:noBreakHyphen/>
    'pageBreakBefore': 0,      # <w:pageBreakBefore/>
    'keepNext': 0,             # <w:keepNext/>
    'keepLines': 0,            # <w:keepLines/>
    'sdt': 0,                  # <w:sdt> content control
    'object': 0,               # <w:object> OLE
    'commentRange': 0,         # <w:commentRangeStart/End>
    'bookmarkStart': 0,        # <w:bookmarkStart>
    'tabLeader': 0,            # <w:tab w:leader=...>
    'pPrChange': 0,            # tracked changes
    'rPrChange': 0,
    'numPr': 0,                # numbered/bulleted lists
    'fldSimple_HYPERLINK': 0,  # hyperlink fields
    'fldChar_HYPERLINK': 0,
}
docs_with = {k: [] for k in features}

for path in sorted(glob.glob(os.path.join(DOCX_DIR, '*.docx'))):
    name = os.path.basename(path)
    try:
        with zipfile.ZipFile(path) as z:
            doc = z.read('word/document.xml').decode('utf-8', errors='replace')
            try: numbering = z.read('word/numbering.xml').decode('utf-8', errors='replace')
            except: numbering = ''
    except: continue

    counts = {
        'sym': len(re.findall(r'<w:sym\s', doc)),
        'softHyphen': doc.count('<w:softHyphen/>'),
        'noBreakHyphen': doc.count('<w:noBreakHyphen/>'),
        'pageBreakBefore': doc.count('<w:pageBreakBefore/>') + len(re.findall(r'<w:pageBreakBefore\s+w:val="(?:true|1)"', doc)),
        'keepNext': len(re.findall(r'<w:keepNext\s*/>|<w:keepNext\s+w:val="(?:true|1)"', doc)),
        'keepLines': len(re.findall(r'<w:keepLines\s*/>|<w:keepLines\s+w:val="(?:true|1)"', doc)),
        'sdt': doc.count('<w:sdt>'),
        'object': doc.count('<w:object'),
        'commentRange': doc.count('<w:commentRangeStart'),
        'bookmarkStart': doc.count('<w:bookmarkStart'),
        'tabLeader': len(re.findall(r'<w:tab\s+[^/>]*?w:leader=', doc)),
        'pPrChange': doc.count('<w:pPrChange'),
        'rPrChange': doc.count('<w:rPrChange'),
        'numPr': len(re.findall(r'<w:numPr>', doc)),
        'fldSimple_HYPERLINK': len(re.findall(r'<w:fldSimple[^>]*?w:instr="[^"]*HYPERLINK', doc)),
        'fldChar_HYPERLINK': len(re.findall(r'<w:instrText[^>]*?>\s*HYPERLINK', doc)),
    }
    for k, v in counts.items():
        if v > 0:
            features[k] += 1
            docs_with[k].append((name, v))

print('=== Feature usage in baseline (184 docs) ===')
for k, count in sorted(features.items(), key=lambda x: -x[1]):
    total_instances = sum(v for _, v in docs_with[k])
    print(f'  {k:<25} {count:>3} docs, {total_instances:>5} total instances')

print()
print('=== Top 5 docs per feature (only features with 5+ docs) ===')
for k in sorted(features, key=lambda f: -features[f]):
    if features[k] < 5: continue
    docs_with[k].sort(key=lambda x: -x[1])
    print(f'\n{k} ({features[k]} docs):')
    for n, v in docs_with[k][:5]:
        print(f'  {n[:48]:<48} {v}')
