"""依頼 D: Cross-doc Mech 2 audit (selective vs proportional).

For each candidate doc with jc=both + kern:
  - Find multi-line body paragraphs (must wrap, must justify)
  - COM-measure per-char advance for each char
  - Classify each yakumono per Mech 1/2/uncompressed:
    * Mech 1 hit: advance ≈ font_size/2 (~half-width, e.g., 5.5pt for 10.5pt MS Mincho)
    * Mech 2 partial: advance between Mech 1 minimum (fontSize/3) and font_size
    * Uncompressed: advance ≈ font_size (full width)

Counts per non-final line:
  - n_mech1_hits, n_mech2_partial, n_uncompressed
  - Classification: SELECTIVE (mech1 + uncompressed coexist) vs PROPORTIONAL
    (all yakumono compressed equally by Mech 2) vs MIXED

Selected docs (jc=both + kern + known Mech 2 activity):
  1. 3a4f9fbe1a83 (Normal kern, R31 winner doc with known Mech 2 lines)
  2. 7f272a2dfd3b (docDefaults kern, R17 big_loser, jc=both)
  3. ed025cbecffb (docDefaults kern, R17 big_loser, jc=both)
  4. d77a58485f16 (Normal kern, R17 big_winner)
  5. b35123fe8efc (docDefaults kern, table-heavy)

Note: Word COM only exposes per-character Information(5) for X position; advances
are derived as x[i+1] - x[i].
"""
import json
import sys
import time
from pathlib import Path
import pythoncom
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX_DIR = Path("tools/golden-test/documents/docx")
OUT = Path("pipeline_data/mech2_cross_doc_audit_part2.json")

# Type A/B/C from spec §4.7
TYPE_A_OPEN = set("（「『【〔｛〈《［" + "“‘")
TYPE_B_CLOSE = set("）」』】〕｝〉》］、。，．" + "”’" + "—")

CANDIDATES = [
    {"doc": "ed025cbecffb_index-23.docx",                              "id": "ed025"},
    {"doc": "d77a58485f16_20240705_resources_data_outline_08.docx",    "id": "d77a"},
    {"doc": "b35123fe8efc_tokumei_08_01.docx",                         "id": "b35123"},
]

MAX_PARAS_TO_SCAN = 30  # cap per doc for time budget


def measure_paragraph_chars(doc, p_idx):
    """Return list of (char, advance_pt, x_pt) for paragraph p_idx."""
    try:
        p = doc.Paragraphs(p_idx)
        rng = p.Range
        text = rng.Text or ""
        # Skip table cell paragraphs (contain \x07)
        if "\x07" in text:
            return None
        # Skip very short
        n_chars = len(text.rstrip())
        if n_chars < 20:
            return None
        # Get per-char x position
        chars = []
        for ci in range(rng.Start, min(rng.Start + n_chars, rng.End)):
            r = doc.Range(ci, ci + 1)
            x = r.Information(5)  # horizontal position
            y = r.Information(6)
            ch = (r.Text or "")[:1]
            chars.append({"ch": ch, "x": x, "y": y})
        # Compute advance from x deltas (group by line via y change)
        if not chars:
            return None
        # Detect line transitions (y changes)
        lines = []
        cur_line = []
        cur_y = None
        for c in chars:
            if cur_y is None or abs(c["y"] - cur_y) < 1.0:
                cur_line.append(c)
                cur_y = c["y"] if cur_y is None else cur_y
            else:
                if cur_line:
                    lines.append(cur_line)
                cur_line = [c]
                cur_y = c["y"]
        if cur_line:
            lines.append(cur_line)
        # For each line, compute advance per char (x[i+1] - x[i])
        for line in lines:
            for i, c in enumerate(line):
                if i + 1 < len(line):
                    c["adv"] = round(line[i+1]["x"] - c["x"], 2)
                else:
                    c["adv"] = None  # last char of line
        # Determine if paragraph wraps (>1 line)
        return {
            "p_idx": p_idx,
            "n_lines": len(lines),
            "lines": lines,
            "first_y": chars[0]["y"],
            "last_y": chars[-1]["y"],
        }
    except Exception:
        return None


def classify_line(line):
    """Classify yakumono in this line as mech1_hit / mech2_partial / uncompressed."""
    n_yakumono = 0
    n_mech1 = 0
    n_mech2 = 0
    n_uncompressed = 0
    yakumono_details = []
    for c in line:
        ch = c["ch"]
        adv = c.get("adv")
        if not adv or not ch:
            continue
        if ch in TYPE_A_OPEN or ch in TYPE_B_CLOSE:
            n_yakumono += 1
            # Determine font_size from advance pattern; assume 10.5pt (most common JP)
            # Mech 1: ~5.0-6.0pt for 10.5pt font (half-width)
            # Mech 2 partial: 7.0-9.0pt
            # Uncompressed: ~10.5pt (full width)
            if adv < 6.5:
                n_mech1 += 1
                cat = "M1"
            elif adv < 10.0:
                n_mech2 += 1
                cat = "M2"
            else:
                n_uncompressed += 1
                cat = "FULL"
            yakumono_details.append({"ch": ch, "adv": adv, "cat": cat})
    return {"n_yakumono": n_yakumono, "n_mech1": n_mech1, "n_mech2": n_mech2,
            "n_uncompressed": n_uncompressed, "yakumono": yakumono_details}


def audit_doc(word, candidate):
    docx_path = DOCX_DIR / candidate["doc"]
    if not docx_path.exists():
        return {"id": candidate["id"], "error": "doc not found"}
    last = None
    for attempt in range(3):
        try:
            doc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
            time.sleep(0.5)
            n_paras = doc.Paragraphs.Count
            results = {"id": candidate["id"], "n_paras": n_paras, "lines_audit": []}
            paras_scanned = 0
            for p_idx in range(1, min(n_paras + 1, 200)):
                if paras_scanned >= MAX_PARAS_TO_SCAN:
                    break
                meas = measure_paragraph_chars(doc, p_idx)
                if meas is None or meas["n_lines"] < 2:
                    continue
                paras_scanned += 1
                # Audit each non-last line
                for line_i, line in enumerate(meas["lines"][:-1]):  # skip last line
                    cls = classify_line(line)
                    if cls["n_yakumono"] == 0:
                        continue
                    cls["p_idx"] = p_idx
                    cls["line_i"] = line_i
                    cls["line_y"] = line[0]["y"]
                    results["lines_audit"].append(cls)
            doc.Close(False)
            return results
        except Exception as e:
            last = e
            time.sleep(0.8 + attempt * 0.5)
    return {"id": candidate["id"], "error": str(last)}


def main():
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    time.sleep(2.0)
    word.Visible = False
    word.DisplayAlerts = False
    all_results = []
    try:
        for cand in CANDIDATES:
            print(f"\nAuditing {cand['id']}...")
            r = audit_doc(word, cand)
            all_results.append(r)
            if "error" in r:
                print(f"  ERR: {r['error']}")
                continue
            n_lines = len(r["lines_audit"])
            n_yak = sum(l["n_yakumono"] for l in r["lines_audit"])
            n_m1 = sum(l["n_mech1"] for l in r["lines_audit"])
            n_m2 = sum(l["n_mech2"] for l in r["lines_audit"])
            n_full = sum(l["n_uncompressed"] for l in r["lines_audit"])
            print(f"  {n_lines} lines audited, {n_yak} total yakumono")
            print(f"  Mech 1: {n_m1}, Mech 2 partial: {n_m2}, uncompressed (full): {n_full}")
            # Classification: lines with both Mech 1 hit AND uncompressed yakumono = SELECTIVE
            n_selective = sum(1 for l in r["lines_audit"]
                              if l["n_mech1"] > 0 and l["n_uncompressed"] > 0)
            n_pure_m1 = sum(1 for l in r["lines_audit"]
                            if l["n_mech1"] > 0 and l["n_uncompressed"] == 0 and l["n_mech2"] == 0)
            n_pure_full = sum(1 for l in r["lines_audit"]
                              if l["n_mech1"] == 0 and l["n_uncompressed"] > 0 and l["n_mech2"] == 0)
            n_proportional = sum(1 for l in r["lines_audit"]
                                 if l["n_mech2"] > 0 and l["n_uncompressed"] == 0)
            print(f"  Lines classification: selective={n_selective}, "
                  f"pure_M1={n_pure_m1}, pure_full={n_pure_full}, proportional_M2={n_proportional}")
    finally:
        try: word.Quit()
        except: pass

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(all_results, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nSaved -> {OUT}")


if __name__ == "__main__":
    main()
