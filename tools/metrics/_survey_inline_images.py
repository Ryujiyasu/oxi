# -*- coding: utf-8 -*-
"""Survey baseline for wp:inline image usage."""
import sys, os, glob, zipfile, re
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

DOCX_DIR = r'C:\Users\ryuji\oxi-4\tools\golden-test\documents\docx'

results = []
for path in sorted(glob.glob(os.path.join(DOCX_DIR, '*.docx'))):
    name = os.path.basename(path)
    try:
        with zipfile.ZipFile(path) as z:
            doc = z.read('word/document.xml').decode('utf-8', errors='replace')
            # Also check media
            n_images = sum(1 for n in z.namelist() if n.startswith('word/media/'))
    except: continue

    inline_count = doc.count('<wp:inline')
    anchor_count = doc.count('<wp:anchor')
    extent_inline = re.findall(r'<wp:inline[^>]*?>.*?<wp:extent\s+cx="(\d+)"\s+cy="(\d+)"', doc, re.DOTALL)

    if inline_count or n_images > 0:
        results.append({
            'name': name,
            'n_images': n_images,
            'inline_count': inline_count,
            'anchor_count': anchor_count,
            'inline_extents': extent_inline[:5],  # first 5 sizes
        })

results.sort(key=lambda r: -r['inline_count'])

print(f"Total docs with images or inline drawings: {len(results)} / 184\n")
print(f"{'doc':<48} {'n_imgs':>6} {'inline':>6} {'anchor':>6} first inline extent (cx, cy in EMU)")
for r in results[:30]:
    extents_str = ' '.join(f"({int(cx)}, {int(cy)})" for cx, cy in r['inline_extents'][:2])
    print(f"  {r['name'][:46]:<46} {r['n_images']:>6} {r['inline_count']:>6} {r['anchor_count']:>6} {extents_str}")

total_inline = sum(r['inline_count'] for r in results)
total_anchor = sum(r['anchor_count'] for r in results)
print(f"\nTotal <wp:inline>: {total_inline}")
print(f"Total <wp:anchor>: {total_anchor}")
docs_with_inline = sum(1 for r in results if r['inline_count'] > 0)
print(f"Docs with inline images: {docs_with_inline}")
