"""COM measurement: Word's actual fitText rendering per character.

Day 37+ multi-week refactor Phase A.

Goal: capture per-character x positions for every fitText paragraph in ed025c
and minimal repros, to reverse-engineer Word's exact fitText algorithm.

Method:
- Open doc via COM
- For each Word paragraph (1..N), check if its text contains characters that
  match an OOXML paragraph with a fitText run (by text-prefix match)
- For matched paragraphs, measure per-character x positions
- Capture OOXML metadata (fitText target, pre-computed cs, kern)
- Compute the actual per-char advances Word renders
- Output JSON for offline analysis

Output: pipeline_data/fittext_widths_<doc_id>.json
"""
from __future__ import annotations
import os, sys, json, traceback, zipfile, re
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc

DOCS = [
    ('ed025c', r'c:\Users\ryuji\oxi-main\tools\golden-test\documents\docx\ed025cbecffb_index-23.docx'),
]
OUT_DIR = os.path.abspath('pipeline_data')


def parse_paragraphs_with_fittext(docx_path):
    """Returns list of dicts: ooxml paragraph index + text + run info, only for paragraphs containing fitText."""
    with zipfile.ZipFile(docx_path) as z:
        xml = z.read('word/document.xml').decode('utf-8')
    paragraphs = re.findall(r'<w:p\b[^>]*>.*?</w:p>', xml, re.DOTALL)
    result = []
    for pi, p in enumerate(paragraphs):
        if '<w:fitText' not in p:
            continue
        runs = re.findall(r'<w:r\b[^>]*>.*?</w:r>', p, re.DOTALL)
        run_infos = []
        full_text = ''
        for r in runs:
            rpr_match = re.search(r'<w:rPr\b[^>]*>(.*?)</w:rPr>', r, re.DOTALL)
            rpr = rpr_match.group(1) if rpr_match else ''
            text_parts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', r)
            text = ''.join(text_parts)
            full_text += text
            ft = re.search(r'<w:fitText w:val="(\d+)"(?: w:id="([^"]+)")?', rpr)
            cs = re.search(r'<w:spacing w:val="(-?\d+)"', rpr)
            kern = re.search(r'<w:kern w:val="(\d+)"', rpr)
            sz = re.search(r'<w:sz w:val="(\d+)"', rpr)
            run_infos.append({
                'text': text,
                'fit_text_tw': int(ft.group(1)) if ft else None,
                'fit_text_id': ft.group(2) if ft and ft.group(2) else None,
                'cs_tw': int(cs.group(1)) if cs else None,
                'kern': int(kern.group(1)) if kern else None,
                'sz_half_pt': int(sz.group(1)) if sz else None,
            })
        result.append({
            'ooxml_pi': pi,
            'text': full_text,
            'runs': run_infos,
        })
    return result


def main():
    for doc_label, docx_path in DOCS:
        print(f'\n[+] Processing {doc_label}')
        ooxml_paras = parse_paragraphs_with_fittext(docx_path)
        print(f'  OOXML fitText paragraphs: {len(ooxml_paras)}')

        # Build text-prefix lookup. Use first 8 chars as key (most paragraphs are unique that way)
        text_to_ooxml = {}
        for entry in ooxml_paras:
            key = entry['text'][:20]
            if key not in text_to_ooxml:
                text_to_ooxml[key] = entry

        word = wc.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        word.ScreenUpdating = False
        doc = None
        try:
            doc = word.Documents.Open(docx_path, ReadOnly=True)
            doc.Repaginate()
            n_paras = doc.Paragraphs.Count

            output_entries = []
            unmatched = 0
            for wpi in range(1, n_paras + 1):
                para = doc.Paragraphs(wpi)
                full_text = (para.Range.Text or '').rstrip('\r\n\x07')
                # Look up in OOXML index by text prefix
                key = full_text[:20]
                if key not in text_to_ooxml:
                    continue
                ooxml_entry = text_to_ooxml[key]
                # Measure per-char positions
                start = para.Range.Start
                chars_measured = []
                for i, ch in enumerate(full_text):
                    try:
                        x = float(doc.Range(start + i, start + i).Information(5))
                        y = float(doc.Range(start + i, start + i).Information(6))
                    except Exception:
                        x, y = None, None
                    chars_measured.append({'i': i, 'ch': ch,
                                           'x': round(x, 3) if x is not None else None,
                                           'y': round(y, 3) if y is not None else None})
                if not chars_measured:
                    continue
                # advances
                advances = []
                for j in range(len(chars_measured) - 1):
                    cur = chars_measured[j]
                    nxt = chars_measured[j+1]
                    if cur['x'] is not None and nxt['x'] is not None and cur['y'] == nxt['y']:
                        advances.append(round(nxt['x'] - cur['x'], 3))
                    else:
                        advances.append(None)
                xs = [c['x'] for c in chars_measured if c['x'] is not None]
                rendered_w = (max(xs) - min(xs)) if xs else 0
                output_entries.append({
                    'word_para_idx': wpi,
                    'ooxml_para_idx': ooxml_entry['ooxml_pi'],
                    'text': full_text,
                    'ooxml_runs': ooxml_entry['runs'],
                    'measured_chars': chars_measured,
                    'measured_advances': advances,
                    'rendered_width_total_pt': round(rendered_w, 3),
                    'n_lines': len(set(c['y'] for c in chars_measured if c['y'] is not None)),
                })
                # Remove from lookup to prevent re-matching different Word paragraphs
                del text_to_ooxml[key]

            out_path = os.path.join(OUT_DIR, f'fittext_widths_{doc_label}.json')
            with open(out_path, 'w', encoding='utf-8') as f:
                json.dump({'doc_label': doc_label,
                           'n_ooxml_fittext_paras': len(ooxml_paras),
                           'n_measured': len(output_entries),
                           'paragraphs': output_entries}, f, ensure_ascii=False, indent=2)
            print(f'  [+] wrote {out_path} ({len(output_entries)}/{len(ooxml_paras)} matched)')
        except Exception:
            traceback.print_exc()
        finally:
            if doc is not None:
                doc.Close(SaveChanges=0)
            word.Quit()


if __name__ == '__main__':
    main()
