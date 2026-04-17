"""Compare Oxi layout dump vs Word DML for any doc; find first drift per page.

Matches paragraphs by text content (robust to table empty-cell y skips).

Usage:
  python tools/metrics/diff_layout_vs_word.py <docname-prefix>

Requires:
  - pipeline_data/_<prefix>_layout.json (from oxi-gdi-renderer --dump-layout)
  - C:/Users/ryuji/oxi-main/pipeline_data/word_dml/<full-docname>.json

Reports per-page:
  - Matched Word paragraphs with Oxi y and drift
  - First paragraph where drift jumps > 5pt (likely a bug location)
"""
import sys, json, os, glob
from collections import defaultdict

if len(sys.argv) < 2:
    print("Usage: python diff_layout_vs_word.py <docname-prefix>")
    sys.exit(1)

prefix = sys.argv[1]

dml_files = glob.glob(f"C:/Users/ryuji/oxi-main/pipeline_data/word_dml/{prefix}*.json")
if not dml_files:
    print(f"No Word DML for prefix '{prefix}'")
    sys.exit(1)
dml_path = dml_files[0]
docname = os.path.splitext(os.path.basename(dml_path))[0]

layout_path = f"pipeline_data/_{prefix}_layout.json"
if not os.path.exists(layout_path):
    print(f"Oxi layout dump not found at {layout_path}.")
    print(f"Generate via: tools/oxi-gdi-renderer/target/release/oxi-gdi-renderer.exe \\")
    print(f"  tools/golden-test/documents/docx/{docname}.docx /tmp/_dummy 150 \\")
    print(f"  --dump-layout={layout_path}")
    sys.exit(1)

print(f"Word DML: {os.path.basename(dml_path)}")
print(f"Oxi dump: {os.path.basename(layout_path)}")

wd = json.load(open(dml_path, encoding='utf-8'))
od = json.load(open(layout_path, encoding='utf-8'), strict=False)
print(f"Word pages: {wd.get('pages')}, Oxi pages: {len(od['pages'])}")

# Build Oxi text index: text → list of (page, y, x)
# New dump format uses 'type' not 'kind'
oxi_texts = []
for pi, page in enumerate(od['pages']):
    for e in page.get('elements', []):
        t = e.get('type', e.get('kind', '?'))
        if t == 'text':
            txt = e.get('text', '')
            if txt.strip():
                oxi_texts.append({
                    'page': pi + 1,
                    'y': e['y'],
                    'x': e['x'],
                    'text': txt,
                    'fs': e.get('font_size', 0),
                    'para_idx': e.get('para_idx'),
                })

# Per-Word-page drift analysis via content matching
word_pages = defaultdict(list)
for p in wd['paragraphs']:
    word_pages[p['page']].append(p)

for page_num in sorted(word_pages.keys())[:3]:
    word_paras = sorted(word_pages[page_num], key=lambda p: p['index'])
    w_with_text = [p for p in word_paras if p.get('text', '').strip() and len(p.get('lines', [])) > 0]
    oxi_page_texts = [t for t in oxi_texts if t['page'] == page_num]

    print(f"\n=== page {page_num}: Word {len(word_paras)} paras ({len(w_with_text)} with text), Oxi {len(oxi_page_texts)} text elements ===")
    print(f"{'wi':>4} {'wy':>7} {'fs':>5} {'oy':>7} {'dy':>7} {'jump':>6} | text")

    prev_dy = None
    max_jump = 0
    max_jump_at = None
    n_match = 0

    for p in w_with_text[:30]:
        wy = p['y']
        fs = p['font_size']
        # Find first distinctive text segment
        full_text = p.get('text', '').replace('\r', '').strip()
        # Use first 10 chars of actual content
        search = full_text[:10]
        if len(search) < 3:
            continue
        # Find matching Oxi text element
        matches = [t for t in oxi_page_texts if search in t['text'] or t['text'].startswith(search[:5])]
        if not matches:
            # Try broader match — any Oxi element starting with first 3 chars
            matches = [t for t in oxi_page_texts if t['text'].startswith(search[:3])]
        if not matches:
            continue

        # Use the first match
        m = matches[0]
        oy = m['y']
        dy = oy - wy
        jump = "" if prev_dy is None else f"{dy - prev_dy:+.2f}"
        mark = ""
        if prev_dy is not None and abs(dy - prev_dy) > max_jump:
            max_jump = abs(dy - prev_dy)
            max_jump_at = (p['index'], wy, oy)
        if prev_dy is not None and abs(dy - prev_dy) > 5:
            mark = " ← BIG JUMP"
        txt_display = full_text[:20].encode('cp932', 'replace').decode('cp932')
        print(f"{p['index']:>4} {wy:>7.1f} {fs:>5.1f} {oy:>7.2f} {dy:>+7.2f} {jump:>6}{mark} | {txt_display!r}")
        prev_dy = dy
        n_match += 1

    if max_jump_at:
        wi, wy, oy = max_jump_at
        print(f"  Max jump: {max_jump:.2f}pt at Word idx {wi} (wy={wy:.1f} oy={oy:.2f})")
    print(f"  Matched {n_match} paragraphs")
