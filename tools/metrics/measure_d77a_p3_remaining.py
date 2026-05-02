"""依頼 13: d77a p3 残り 6 paragraph (P27, P28, P30, P34, P35, P41).

Strategy: 1 Word session per paragraph (single-batch). Slow but reliable.
"""
import win32com.client
import os
import time
import json
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DOCX = os.path.abspath(
    "tools/golden-test/documents/docx/"
    "d77a58485f16_20240705_resources_data_outline_08.docx")
EXISTING_PATH = os.path.abspath(
    "pipeline_data/d77a_p1_p3_per_char_2026-05-02.json")
RESULT_PATH = EXISTING_PATH  # extend in-place

YAKUMONO_A = set("（「『【〔｛〈《［")
YAKUMONO_B = set("）」』】〕｝〉》］、。，．—")
YAKUMONO_C = set("・：；！？ー―／＼")

# Missing paragraphs from prior run
TARGET_INDICES = [27, 28, 30, 34, 35, 41]


def classify(ch):
    if ch in YAKUMONO_A:
        return "A"
    if ch in YAKUMONO_B:
        return "B"
    if ch in YAKUMONO_C:
        return "C"
    return None


def measure_one_paragraph(pi):
    """Open Word, measure single paragraph by index, close Word."""
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        d = word.Documents.Open(DOCX, ReadOnly=True)
        time.sleep(1.0)
        p = d.Paragraphs(pi)
        rng = p.Range
        text = rng.Text.strip()
        page = int(rng.Information(3))
        chars = rng.Characters
        per_char = []
        for ci in range(1, chars.Count + 1):
            try:
                c = chars(ci)
                t = c.Text
                if t in ("\r", "\x07"):
                    continue
                per_char.append({
                    "i": ci,
                    "ch": t,
                    "x": round(float(c.Information(5)), 4),
                    "y": round(float(c.Information(6)), 4),
                    "size": c.Font.Size,
                    "font": c.Font.Name,
                })
            except Exception:
                continue
        try:
            d.Close(SaveChanges=False)
        except Exception:
            pass
        if not per_char:
            return {"para_idx": pi, "page": page, "text": text,
                    "n_lines": 0, "lines": []}
        # Group by line
        lines = {}
        for r in per_char:
            lines.setdefault(round(r["y"], 1), []).append(r)
        line_data = []
        for y in sorted(lines.keys()):
            sorted_chars = sorted(lines[y], key=lambda r: r["x"])
            advs = []
            for i in range(len(sorted_chars) - 1):
                ch = sorted_chars[i]["ch"]
                adv = round(sorted_chars[i + 1]["x"]
                             - sorted_chars[i]["x"], 4)
                sz = sorted_chars[i]["size"]
                ratio = round(adv / sz, 3) if sz else None
                yclass = classify(ch)
                next_ch = sorted_chars[i + 1]["ch"]
                prev_ch = sorted_chars[i - 1]["ch"] if i > 0 else None
                rule_match = "none"
                if yclass == "A":
                    pc = classify(prev_ch) if prev_ch else None
                    if pc == "A":
                        rule_match = "A_after_A"
                elif yclass == "B":
                    nc = classify(next_ch) if next_ch else None
                    if nc in ("A", "B"):
                        rule_match = f"B_before_{nc}"
                advs.append({
                    "i": sorted_chars[i]["i"],
                    "ch": ch,
                    "prev_ch": prev_ch,
                    "next_ch": next_ch,
                    "adv": adv,
                    "size": sz,
                    "font": sorted_chars[i].get("font"),
                    "ratio": ratio,
                    "yakumono_class": yclass,
                    "rule_match": rule_match,
                    "compressed_word": (ratio is not None and ratio < 0.85
                                         and yclass is not None),
                })
            line_data.append({
                "y": y,
                "n_chars": len(sorted_chars),
                "advances": advs,
            })
        return {"para_idx": pi, "page": page, "text": text,
                "n_lines": len(line_data), "lines": line_data}
    except Exception as e:
        return {"para_idx": pi, "error": str(e)}
    finally:
        try:
            word.Quit()
        except Exception:
            pass
        time.sleep(2.0)


def load_existing():
    if os.path.exists(EXISTING_PATH):
        with open(EXISTING_PATH, encoding="utf-8") as f:
            return json.load(f)
    return {"page_1": [], "page_3": [], "errors": []}


def save(out):
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)


def main():
    out = load_existing()
    # Drop existing errors entries we're going to retry
    out["errors"] = [e for e in out.get("errors", [])
                      if e.get("para_idx") not in TARGET_INDICES]
    for pi in TARGET_INDICES:
        # Skip if already have it
        existing = [r for r in out["page_3"] if r.get("para_idx") == pi]
        if existing:
            print(f"  P{pi}: already have, skip", flush=True)
            continue
        print(f"\n=== Measuring P{pi} ===", flush=True)
        res = measure_one_paragraph(pi)
        if "error" in res:
            print(f"  P{pi} ERR: {res['error']}", flush=True)
            out.setdefault("errors", []).append(res)
        else:
            yak_compressed = []
            yak_not_compressed = []
            for ln in res["lines"]:
                for a in ln["advances"]:
                    if a["yakumono_class"]:
                        if a["compressed_word"]:
                            yak_compressed.append(a)
                        else:
                            yak_not_compressed.append(a)
            rule_NOT = [a for a in yak_not_compressed
                         if a["rule_match"] != "none"]
            rule_YES = [a for a in yak_compressed
                         if a["rule_match"] != "none"]
            print(f"  text: {res['text'][:60]!r}")
            print(f"  yak={len(yak_compressed) + len(yak_not_compressed)} "
                  f"compressed={len(yak_compressed)} "
                  f"rule_match_compressed={len(rule_YES)} "
                  f"rule_match_NOT_compressed={len(rule_NOT)}")
            for a in rule_YES[:5]:
                print(f"    MATCH [{a['i']}] {a['ch']!r} "
                      f"prev={a['prev_ch']!r} next={a['next_ch']!r} "
                      f"rule={a['rule_match']} adv={a['adv']} r={a['ratio']}")
            for a in rule_NOT[:5]:
                print(f"    ANOMALY [{a['i']}] {a['ch']!r} "
                      f"prev={a['prev_ch']!r} next={a['next_ch']!r} "
                      f"rule={a['rule_match']} adv={a['adv']}")
            out["page_3"].append(res)
        save(out)  # persist after each paragraph
    print(f"\nFinal: p3 = {len(out['page_3'])} paragraphs, "
          f"errors = {len(out.get('errors', []))}")


if __name__ == "__main__":
    main()
