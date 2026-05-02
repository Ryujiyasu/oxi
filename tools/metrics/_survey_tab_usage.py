# -*- coding: utf-8 -*-
"""Survey baseline for <w:tab/> usage and tab stop definitions."""
import sys, os, glob, zipfile, re
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

DOCX_DIR = r'C:\Users\ryuji\oxi-4\tools\golden-test\documents\docx'

results = []
for path in sorted(glob.glob(os.path.join(DOCX_DIR, '*.docx'))):
    name = os.path.basename(path)
    try:
        with zipfile.ZipFile(path) as z:
            doc = z.read('word/document.xml').decode('utf-8', errors='replace')
            settings = z.read('word/settings.xml').decode('utf-8', errors='replace')
    except: continue

    # Count tab elements (within run content)
    tab_chars = doc.count('<w:tab/>') + doc.count('<w:tab ')
    # Tab definitions in pPr
    explicit_tabs = len(re.findall(r'<w:tabs>.*?</w:tabs>', doc, re.DOTALL))
    # Default tab stop from settings
    default_tab = re.search(r'<w:defaultTabStop[^/>]*?w:val="(\d+)"', settings)
    default_tab_val = int(default_tab.group(1)) if default_tab else 720
    # Tab leaders
    tab_leaders = len(re.findall(r'<w:tab[^/>]*?w:leader="', doc))
    # Different tab stop types
    tab_types = {}
    for m in re.finditer(r'<w:tab\s+w:val="([^"]+)"[^/>]*?w:pos="(-?\d+)"', doc):
        v = m.group(1)
        tab_types[v] = tab_types.get(v, 0) + 1

    if tab_chars or explicit_tabs:
        results.append({
            'name': name,
            'tab_chars': tab_chars,
            'explicit_tabs': explicit_tabs,
            'default_tab': default_tab_val,
            'tab_leaders': tab_leaders,
            'tab_types': tab_types,
        })

# Sort by tab_chars desc
results.sort(key=lambda r: -r['tab_chars'])

print(f"Total docs with tabs: {len(results)} / 184")
print(f"\nTop 30 by tab character count:")
print(f"{'doc':<40} {'tab_chars':>10} {'expl_tabs':>10} {'default':>8} {'leaders':>7} types")
for r in results[:30]:
    types_str = ' '.join(f"{k}:{v}" for k, v in sorted(r['tab_types'].items()))
    print(f"  {r['name'][:38]:<38} {r['tab_chars']:>10} {r['explicit_tabs']:>10} {r['default_tab']:>8} {r['tab_leaders']:>7} {types_str}")

# Summary statistics
total_tab_chars = sum(r['tab_chars'] for r in results)
total_explicit = sum(r['explicit_tabs'] for r in results)
docs_with_leaders = sum(1 for r in results if r['tab_leaders'] > 0)
print(f"\nTotal <w:tab/> across baseline: {total_tab_chars}")
print(f"Total explicit <w:tabs>: {total_explicit}")
print(f"Docs with tab leaders: {docs_with_leaders}")

# All tab types
all_types = {}
for r in results:
    for k, v in r['tab_types'].items():
        all_types[k] = all_types.get(k, 0) + v
print(f"\nAll tab stop types in baseline:")
for k, v in sorted(all_types.items()):
    print(f"  {k}: {v}")
