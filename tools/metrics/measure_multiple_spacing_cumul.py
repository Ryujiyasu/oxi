"""Measure Multiple spacing cumulative round pattern across multiple documents.
Goal: determine if j carries across paragraphs, how headings affect j,
and whether CEIL or ROUND is used at each position."""
import win32com.client
import os
import glob
import math

def measure_doc(word, docx_path):
    """Measure all Multiple spacing paragraphs in a document."""
    doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
    results = []

    try:
        total = doc.Paragraphs.Count
        prev_y = None
        prev_fs = None
        prev_sa = None
        prev_sb = None

        for i in range(1, min(total + 1, 80)):  # limit to 80 paras
            p = doc.Paragraphs(i)
            rng = p.Range
            page = rng.Information(3)
            if page > 1:
                break  # only page 1

            y = rng.Information(6)
            ls = p.Format.LineSpacing
            rule = p.Format.LineSpacingRule
            fs = rng.Font.Size
            sa = p.Format.SpaceAfter
            sb = p.Format.SpaceBefore

            # Count lines
            end_rng = doc.Range(rng.End - 1, rng.End)
            end_y = end_rng.Information(6)
            n_lines = max(1, round((end_y - y) / ls) + 1) if ls > 0 and end_y > y + 2 else 1

            gap = y - prev_y if prev_y is not None else 0

            results.append({
                'i': i, 'y': y, 'gap': gap, 'fs': fs, 'ls': ls, 'rule': rule,
                'sa': sa, 'sb': sb, 'n_lines': n_lines,
                'prev_fs': prev_fs, 'prev_sa': prev_sa,
            })

            prev_y = end_y if n_lines > 1 else y
            prev_fs = fs
            prev_sa = sa
            prev_sb = sb
    finally:
        doc.Close(False)

    return results


def analyze_advances(results):
    """Extract body text advances and find cumulative round pattern."""
    # Find the most common font size (= body text)
    fs_counts = {}
    for r in results:
        if r['rule'] == 5:  # Multiple spacing only
            fs_counts[r['fs']] = fs_counts.get(r['fs'], 0) + 1

    if not fs_counts:
        return None

    body_fs = max(fs_counts, key=fs_counts.get)

    # Extract consecutive body-to-body advances
    advances = []
    for k in range(1, len(results)):
        r = results[k]
        prev = results[k - 1]

        if r['rule'] != 5:
            continue

        gap = r['gap']
        if gap <= 0:
            continue

        # Determine spacing added
        collapsed = max(prev.get('prev_sa', 0) or 0, r['sb'])

        # Check if spacing was suppressed (contextual)
        # If gap < expected_advance + collapsed, spacing was likely suppressed
        advance = gap - collapsed if collapsed > 0 and gap > collapsed else gap

        advances.append({
            'from': prev['i'], 'to': r['i'],
            'gap': gap, 'advance': advance,
            'from_fs': prev['fs'], 'to_fs': r['fs'],
            'is_body': r['fs'] == body_fs and prev['fs'] == body_fs,
        })

    return {
        'body_fs': body_fs,
        'advances': advances,
    }


word = win32com.client.Dispatch("Word.Application")
word.Visible = False

try:
    # Find gen2 documents with Multiple spacing
    docs_dir = os.path.abspath("tools/golden-test/documents/docx")
    gen2_docs = sorted(glob.glob(os.path.join(docs_dir, "gen2_*.docx")))

    print(f"Scanning {len(gen2_docs)} gen2 documents...")
    print()

    for docx in gen2_docs[:20]:  # first 20
        name = os.path.basename(docx)[:40]
        results = measure_doc(word, docx)

        # Check if it has Multiple spacing
        mult_count = sum(1 for r in results if r['rule'] == 5)
        if mult_count < 5:
            continue

        analysis = analyze_advances(results)
        if not analysis:
            continue

        body_fs = analysis['body_fs']
        body_advances = [a for a in analysis['advances'] if a['is_body']]

        if len(body_advances) < 3:
            continue

        # Check advance pattern
        adv_values = [a['advance'] for a in body_advances]
        unique_advs = sorted(set(round(a, 1) for a in adv_values))

        print(f"=== {name} (body fs={body_fs}pt) ===")
        print(f"  Body advances: {[round(a,1) for a in adv_values]}")
        print(f"  Unique: {unique_advs}")

        # Check if alternating pattern exists
        if len(unique_advs) == 2:
            diff = unique_advs[1] - unique_advs[0]
            if abs(diff - 0.5) < 0.1:
                # Find positions of the smaller advance
                positions = [k for k, a in enumerate(adv_values) if round(a, 1) == unique_advs[0]]
                print(f"  Pattern: {unique_advs[0]} at positions {positions}")

                # Try to match with ROUND cumulative at different j offsets
                # CJK 83/64: floor(fs * 83/64 * 8) / 8
                cjk_h = math.floor(body_fs * 83/64 * 8) / 8
                raw_tw = cjk_h * 1.15 * 20

                for j_start in range(20):
                    predicted = []
                    j = j_start
                    for _ in range(len(adv_values)):
                        cn = round((j+1) * raw_tw / 10) * 10
                        cc = round(j * raw_tw / 10) * 10
                        predicted.append((cn - cc) / 20)
                        j += 1

                    if [round(p, 1) for p in predicted] == [round(a, 1) for a in adv_values]:
                        print(f"  MATCH: ROUND with j_start={j_start}, raw_tw={raw_tw:.2f}")
                        break
                else:
                    # Try CEIL
                    for j_start in range(20):
                        predicted = []
                        j = j_start
                        for _ in range(len(adv_values)):
                            cn = math.ceil((j+1) * raw_tw / 10) * 10
                            cc = math.ceil(j * raw_tw / 10) * 10
                            predicted.append((cn - cc) / 20)
                            j += 1

                        if [round(p, 1) for p in predicted] == [round(a, 1) for a in adv_values]:
                            print(f"  MATCH: CEIL with j_start={j_start}, raw_tw={raw_tw:.2f}")
                            break
                    else:
                        print(f"  NO MATCH for raw_tw={raw_tw:.2f}")
        print()

    # Also check 0e7a (Single spacing with Multiple P1)
    print("=== 0e7a P1 (Multiple 1.15x, MS Mincho 10.5pt) ===")
    oxi_doc = os.path.join(docs_dir, "0e7af1ae8f21_20230331_resources_open_data_contract_sample_00.docx")
    results = measure_doc(word, oxi_doc)
    p1 = results[0]
    p2 = results[1]
    print(f"  P1: y={p1['y']}, fs={p1['fs']}, ls={p1['ls']}, rule={p1['rule']}")
    print(f"  P2: y={p2['y']}, gap={p2['gap']}")
    cjk_h = math.floor(10.5 * 83/64 * 8) / 8
    raw_tw = cjk_h * 1.15 * 20
    ceil_adv = math.ceil(raw_tw / 10) * 10 / 20
    round_adv = round(raw_tw / 10) * 10 / 20
    print(f"  CJK 83/64 base={cjk_h}, raw_tw={raw_tw}")
    print(f"  CEIL advance={ceil_adv}, ROUND advance={round_adv}")
    print(f"  Word advance={p2['gap']} -> {'CEIL' if abs(p2['gap'] - ceil_adv) < 0.1 else 'ROUND' if abs(p2['gap'] - round_adv) < 0.1 else 'UNKNOWN'}")

finally:
    word.Quit()
