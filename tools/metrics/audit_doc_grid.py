"""Day 32 part 8b — docGrid pitch audit.

Day 32 part 8a (textAlignment hypothesis) FALSIFIED — Class A and
preserve docs both have docDefault textAlignment=None.

Hypothesis 2: docGrid line-pitch discriminates. db9ca/d77a58 +7pt
drift = (28-14)/2 = 7pt suggests grid_pitch creates 28pt line_height
for 14pt fs. bd90b00 +2.5pt suggests grid_pitch ≈ fs creating
single-cell line_height.

This tool extracts <w:docGrid> values from sectPr and computes
expected Bug 2 offset for given font sizes.
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


def audit(doc_id, label):
    docx = find_docx(doc_id)
    if not docx:
        print(f'{label} {doc_id}: NOT FOUND')
        return None
    with zipfile.ZipFile(docx) as zf:
        try:
            xml = zf.read('word/document.xml').decode('utf-8')
        except KeyError:
            return None
    # docGrid: <w:docGrid w:type="lines" w:linePitch="..." w:charSpace="..."/>
    m = re.search(r'<w:docGrid[^/>]*w:linePitch="(\d+)"', xml)
    line_pitch = int(m.group(1)) if m else None
    pitch_pt = (line_pitch / 20.0) if line_pitch else None  # twips → pt
    m = re.search(r'<w:docGrid[^/>]*w:type="(\w+)"', xml)
    grid_type = m.group(1) if m else None
    # First few paragraphs' fontSize
    fs_pattern = re.findall(r'<w:sz w:val="(\d+)"/>', xml[:8000])
    first_fs = [int(x) / 2.0 for x in fs_pattern[:5]]
    print(f'{label:<10} {doc_id}:')
    print(f'  docGrid type: {grid_type!r} linePitch: {line_pitch}tw = {pitch_pt}pt')
    print(f'  First fontSizes: {first_fs}')
    if pitch_pt and first_fs:
        for fs in first_fs[:3]:
            n_cells = max(1, int((fs + pitch_pt - 0.001) / pitch_pt))
            lh = n_cells * pitch_pt
            offset = (lh - fs) / 2.0
            print(f'    fs={fs}pt: {n_cells} cells * {pitch_pt}pt = lh={lh}pt, expected Bug 2 offset = {offset}pt')
    return {'doc_id': doc_id, 'pitch_pt': pitch_pt, 'first_fs': first_fs}


def main():
    print('=== Class A docs ===')
    class_a = ['bd90b00ab7a7', 'de6e32b5960b', 'db9ca18368cd', 'd77a58485f16']
    for d in class_a:
        audit(d, 'Class A')
        print()

    print('=== Preserve-class sample ===')
    preserve = ['e3c545fac7a7', '0e7af1ae8f21']
    for d in preserve:
        audit(d, 'Preserve')
        print()


if __name__ == '__main__':
    main()
