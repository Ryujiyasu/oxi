"""S182: Build doc-feature taxonomy for the 55-doc baseline.

For each docx, extract structural features that predict layout
behavior:
  - docGrid type (none / lines / linesAndChars)
  - n_tables, max_rows_per_table, has_back_to_back_tables
  - n_rows_with_trHeight, n_rows_no_trHeight
  - has_explicit_tcMar (any cell)
  - has_inside_h_border (any table)
  - has_outer_border (any table)
  - n_style_with_spacing (styles.xml <w:style> blocks with <w:spacing>)
  - n_inline_spacing (document.xml <w:spacing> in pPr)
  - n_paras_total, n_paras_empty
  - max_table_depth (for nested tables)

Cross-reference with current per-doc IoU + positional dy (med_dy_vis)
to identify which docs would be impacted by candidate fixes.

Output: pipeline_data/doc_feature_taxonomy.json
        stdout: summary table + candidate-fix impact predictions

Usage:
  python tools/metrics/doc_feature_taxonomy.py
"""
from __future__ import annotations
import os, sys, json, re, zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

REPO = Path(__file__).resolve().parent.parent.parent
DOCS_DIR = REPO / 'tools' / 'golden-test' / 'documents' / 'docx'
IOU_SUMMARY = REPO / 'pipeline_data' / 'element_iou_diff' / '_summary.json'
POS_SUMMARY = REPO / 'pipeline_data' / 'pagination_diff_positional' / '_summary.json'
OUT_JSON = REPO / 'pipeline_data' / 'doc_feature_taxonomy.json'


def extract_features(docx_path: Path) -> dict:
    f = {
        'doc_id': docx_path.stem.split('_')[0],
        'filename': docx_path.name,
    }
    try:
        with zipfile.ZipFile(docx_path) as z:
            doc_xml = z.read('word/document.xml').decode('utf-8', errors='replace')
            styles_xml = ''
            if 'word/styles.xml' in z.namelist():
                styles_xml = z.read('word/styles.xml').decode('utf-8', errors='replace')
    except Exception as e:
        f['error'] = str(e)
        return f

    # docGrid type
    g = re.search(r'<w:docGrid\s+([^/>]*)/?>', doc_xml)
    grid_type = 'none'
    if g:
        attrs = g.group(1)
        tm = re.search(r'w:type="([^"]+)"', attrs)
        grid_type = tm.group(1) if tm else 'lines'
    f['grid_type'] = grid_type

    # Tables — count and structure
    table_pattern = re.compile(r'<w:tbl>.*?</w:tbl>', re.DOTALL)
    tables = table_pattern.findall(doc_xml)
    f['n_tables'] = len(tables)

    # Build per-table info to detect back-to-back
    # Strategy: find all <w:tbl> AND <w:p> positions in document body;
    # check if any two consecutive blocks are both <w:tbl>
    body_match = re.search(r'<w:body>(.*?)</w:body>', doc_xml, re.DOTALL)
    body = body_match.group(1) if body_match else doc_xml
    block_positions = []
    # Walk top-level blocks (tbl, p, sectPr)
    depth = 0
    i = 0
    while i < len(body):
        # Look for next <w:tbl ... > or <w:p ... > at depth 0
        tag_m = re.match(r'<w:(tbl|p|sectPr)([\s>])', body[i:])
        if tag_m:
            tag = tag_m.group(1)
            if tag == 'sectPr':
                # skip
                end = body.find('</w:sectPr>', i)
                if end < 0: break
                i = end + len('</w:sectPr>')
                continue
            # find matching close
            close_tag = f'</w:{tag}>'
            # find next close at same depth; naive: count opens/closes
            j = i
            open_tag_prefix = f'<w:{tag}'
            d = 1
            j = i + tag_m.end()
            while j < len(body) and d > 0:
                next_open = body.find(open_tag_prefix, j)
                next_close = body.find(close_tag, j)
                if next_close < 0:
                    break
                # Self-closing detection
                if next_open >= 0 and next_open < next_close:
                    # Check if self-closing <w:p .../>
                    ge = body.find('>', next_open)
                    if ge > 0 and body[ge-1] == '/':
                        # self-closing — no nested
                        j = ge + 1
                        continue
                    d += 1
                    j = body.find('>', next_open) + 1
                else:
                    d -= 1
                    j = next_close + len(close_tag)
            block_positions.append((tag, i, j))
            i = j
        else:
            i += 1
    # block_positions: list of (tag, start, end)
    f['n_top_blocks'] = len(block_positions)
    # back-to-back tables
    n_b2b = 0
    for k in range(1, len(block_positions)):
        if block_positions[k][0] == 'tbl' and block_positions[k-1][0] == 'tbl':
            n_b2b += 1
    f['n_back_to_back_tables'] = n_b2b
    f['has_back_to_back_tables'] = n_b2b > 0

    # trHeight presence
    trh_pattern = re.compile(r'<w:trHeight\s+[^/]*?w:val="(\d+)"')
    trh_values = [int(m.group(1)) for m in trh_pattern.finditer(doc_xml)]
    f['n_trH_rows'] = len(trh_values)
    f['trH_values_sample'] = trh_values[:5]

    # tcMar explicit
    f['n_tcMar_explicit'] = len(re.findall(r'<w:tcMar>', doc_xml))

    # Border features
    f['has_outer_border'] = '<w:tblBorders>' in doc_xml or '<w:tcBorders>' in doc_xml
    f['has_inside_h'] = '<w:insideH' in doc_xml

    # Inline spacing (pPr level) — count <w:spacing> within pPr
    # rough proxy: count <w:spacing> with before/after/beforeLines/afterLines
    inline_spacing = re.findall(r'<w:spacing\s+[^/]*?(before|after|beforeLines|afterLines)', doc_xml)
    f['n_inline_spacing_with_sb_sa'] = len(inline_spacing)

    # Style-defined spacing
    style_with_spacing = 0
    style_pattern = re.compile(r'<w:style\s[^>]*>.*?</w:style>', re.DOTALL)
    for sm in style_pattern.finditer(styles_xml):
        if re.search(r'<w:pPr>.*?<w:spacing\s+[^/]*?(before|after|beforeLines|afterLines)[^/]*?/>', sm.group(), re.DOTALL):
            style_with_spacing += 1
    f['n_style_with_spacing'] = style_with_spacing

    # Paragraph counts
    p_pattern = re.compile(r'<w:p[\s>]', re.DOTALL)
    f['n_paras_total'] = len(p_pattern.findall(doc_xml))

    # nested tables
    f['n_nested_tables'] = doc_xml.count('<w:tbl>') - len(tables)  # tables not at top level

    return f


def load_iou_signals() -> dict:
    """Load per-doc IoU + positional med_dy signals."""
    out = {}
    if IOU_SUMMARY.exists():
        with open(IOU_SUMMARY, encoding='utf-8') as f:
            d = json.load(f)
        for x in d.get('docs', []):
            out[x['doc_id']] = {'iou': x.get('mean_iou')}
    if POS_SUMMARY.exists():
        with open(POS_SUMMARY, encoding='utf-8') as f:
            d = json.load(f)
        for x in d.get('docs', []):
            did = x['doc_id']
            out.setdefault(did, {})
            out[did]['pos_conf'] = x.get('alignment_confidence')
            out[did]['med_dy'] = x.get('y_diff_visual_median')
            out[did]['n_match'] = x.get('n_match')
    return out


def main():
    if not DOCS_DIR.exists():
        print(f'docs dir not found: {DOCS_DIR}')
        return
    docs = sorted(DOCS_DIR.glob('*.docx'))
    signals = load_iou_signals()

    rows = []
    for d in docs:
        if d.name.startswith('~$'):
            continue
        feats = extract_features(d)
        did = feats['doc_id']
        if did in signals:
            feats['iou'] = signals[did].get('iou')
            feats['pos_conf'] = signals[did].get('pos_conf')
            feats['med_dy'] = signals[did].get('med_dy')
            feats['n_match'] = signals[did].get('n_match')
        rows.append(feats)

    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT_JSON, 'w', encoding='utf-8') as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)
    print(f'Wrote {len(rows)} entries to {OUT_JSON}\n')

    # Summary table
    print(f'{"doc_id":<14} {"grid":<14} {"tbls":>4} {"b2b":>3} {"trH":>4} {"tcM":>3} {"iH":<3} {"iou":>6} {"med_dy":>7}')
    for r in rows:
        iou_s = f'{r["iou"]:.4f}' if r.get('iou') is not None else '-'
        med_s = f'{r["med_dy"]:+.2f}' if r.get('med_dy') is not None else '-'
        print(f'  {r["doc_id"]:<14} {r["grid_type"]:<14} {r["n_tables"]:>4} {r["n_back_to_back_tables"]:>3} {r["n_trH_rows"]:>4} {r["n_tcMar_explicit"]:>3} {"Y" if r["has_inside_h"] else "N":<3} {iou_s:>6} {med_s:>7}')

    # Candidate fix impact predictions
    print('\n=== Candidate fix impact predictions ===')

    # S181 fix (BugA back-to-back, charGrid-less): affected docs
    s181_candidates = [r for r in rows if r.get('has_back_to_back_tables') and r.get('grid_type') in ('none', 'lines') and r.get('has_outer_border', False)]
    print(f'\n1. S181 BugA back-to-back fix would affect {len(s181_candidates)} docs:')
    for r in s181_candidates:
        print(f'     {r["doc_id"]}: tbls={r["n_tables"]} b2b={r["n_back_to_back_tables"]} grid={r["grid_type"]} iou={r.get("iou","-")}')

    # S180 candidate (implicit_pad_t row 0 of all tables): all docs with tcMar=0+border
    s180_candidates = [r for r in rows if r['n_tables'] > 0 and r['has_outer_border'] and r['n_tcMar_explicit'] == 0]
    print(f'\n2. S180 implicit_pad_t row_idx==0 fix would affect {len(s180_candidates)} docs (tables × outer border × no explicit tcMar)')

    # New candidate: trH=N tables without charGrid (RBM C/F dy=-1.85 bug)
    trh_no_grid = [r for r in rows if r['n_trH_rows'] > 0 and r['grid_type'] in ('none', 'lines')]
    print(f'\n3. trH + no-charGrid (RBM C/F bug class) would affect {len(trh_no_grid)} docs:')
    for r in trh_no_grid[:10]:
        print(f'     {r["doc_id"]}: trH_rows={r["n_trH_rows"]} grid={r["grid_type"]} iou={r.get("iou","-")} med_dy={r.get("med_dy","-")}')


if __name__ == '__main__':
    main()
