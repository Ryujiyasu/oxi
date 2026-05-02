# -*- coding: utf-8 -*-
"""Survey baseline for field code usage and types."""
import sys, os, glob, zipfile, re
from collections import Counter
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

DOCX_DIR = r'C:\Users\ryuji\oxi-4\tools\golden-test\documents\docx'

# Field code patterns to look for
results = []
all_field_types = Counter()
for path in sorted(glob.glob(os.path.join(DOCX_DIR, '*.docx'))):
    name = os.path.basename(path)
    try:
        with zipfile.ZipFile(path) as z:
            parts = {}
            for n in z.namelist():
                if n.endswith('.xml') and ('document' in n or 'header' in n or 'footer' in n or 'footnote' in n or 'endnote' in n):
                    try: parts[n] = z.read(n).decode('utf-8', errors='replace')
                    except: pass
    except: continue

    counts = {
        'fldSimple': 0,         # <w:fldSimple instr=...>...</w:fldSimple>
        'fldChar_begin': 0,     # <w:fldChar w:fldCharType="begin"/>
        'fldChar_end': 0,
        'fldChar_separate': 0,
        'instrText': 0,         # <w:instrText> ... </w:instrText>
        'fields_in_header': 0,
        'fields_in_footer': 0,
    }
    field_types_in_doc = Counter()
    for partname, content in parts.items():
        counts['fldSimple'] += content.count('<w:fldSimple')
        counts['fldChar_begin'] += content.count('w:fldCharType="begin"')
        counts['fldChar_end'] += content.count('w:fldCharType="end"')
        counts['fldChar_separate'] += content.count('w:fldCharType="separate"')
        counts['instrText'] += content.count('<w:instrText')
        # Extract instr field types
        for m in re.finditer(r'<w:fldSimple[^>]*?w:instr="([^"]*)"', content):
            instr = m.group(1).strip()
            # Just first word (PAGE, DATE, TOC, etc.)
            field_type = instr.split()[0] if instr.split() else '?'
            field_types_in_doc[field_type] += 1
            all_field_types[field_type] += 1
        for m in re.finditer(r'<w:instrText[^>]*?>(.*?)</w:instrText>', content, re.DOTALL):
            instr = m.group(1).strip()
            field_type = instr.split()[0] if instr.split() else '?'
            field_types_in_doc[field_type] += 1
            all_field_types[field_type] += 1
        if 'header' in partname:
            counts['fields_in_header'] += content.count('<w:fldSimple') + content.count('w:fldCharType="begin"')
        elif 'footer' in partname:
            counts['fields_in_footer'] += content.count('<w:fldSimple') + content.count('w:fldCharType="begin"')

    if any(counts.values()):
        results.append({'name': name, 'counts': counts, 'types': dict(field_types_in_doc)})

results.sort(key=lambda r: -(r['counts']['fldSimple'] + r['counts']['fldChar_begin']))

print(f"Total docs with fields: {len(results)} / 184\n")
print(f"=== All field types in baseline ===")
for ft, cnt in sorted(all_field_types.items(), key=lambda x: -x[1]):
    print(f"  {ft:<20} {cnt}")

print(f"\n=== Top 20 docs by field count ===")
print(f"{'doc':<48} {'simple':>6} {'begin':>5} {'instr':>5} {'hdr':>4} {'ftr':>4} types")
for r in results[:20]:
    c = r['counts']
    types_str = ' '.join(f"{k}:{v}" for k, v in sorted(r['types'].items()))[:60]
    print(f"  {r['name'][:46]:<46} {c['fldSimple']:>6} {c['fldChar_begin']:>5} {c['instrText']:>5} {c['fields_in_header']:>4} {c['fields_in_footer']:>4} {types_str}")
