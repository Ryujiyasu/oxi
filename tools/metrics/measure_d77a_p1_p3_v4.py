"""依頼 9 v4: d77a p1 + p3 — direct paragraph index access (no iteration).

p1 paragraphs (with text): idx 2, 4, 5, 8, 10, 12, 13, 15
p3 paragraphs (with text): idx 26, 27, 28, 30, 32, 33, 34, 35, 38, 39, 40, 41

Strategy: open Word, access each paragraph by its index, measure per-char,
restart Word every N paragraphs to avoid RPC death.
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
RESULT_PATH = os.path.abspath(
    "pipeline_data/d77a_p1_p3_per_char_2026-05-02.json")

YAKUMONO_A = set("（「『【〔｛〈《［")
YAKUMONO_B = set("）」』】〕｝〉》］、。，．—")
YAKUMONO_C = set("・：；！？ー―／＼")

P1_INDICES = [2, 4, 5, 8, 10, 12, 13, 15]
P3_INDICES = [26, 27, 28, 30, 32, 33, 34, 35, 38, 39, 40, 41]


def classify(ch):
    if ch in YAKUMONO_A:
        return "A"
    if ch in YAKUMONO_B:
        return "B"
    if ch in YAKUMONO_C:
        return "C"
    return None


def measure_paragraph_idx(word, doc, pi):
    p = doc.Paragraphs(pi)
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
    if not per_char:
        return {"para_idx": pi, "page": page, "text": text,
                "n_lines": 0, "lines": []}
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


def save(out):
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)


def measure_indices_with_restart(indices, batch_size=4):
    """Restart Word every batch_size paragraphs to avoid RPC death."""
    out = []
    for i in range(0, len(indices), batch_size):
        batch = indices[i:i + batch_size]
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        try:
            d = word.Documents.Open(DOCX, ReadOnly=True)
            time.sleep(1.0)
            for pi in batch:
                try:
                    res = measure_paragraph_idx(word, d, pi)
                    out.append(res)
                    yak = sum(1 for ln in res["lines"]
                              for a in ln["advances"]
                              if a["yakumono_class"])
                    yak_compr = sum(1 for ln in res["lines"]
                                    for a in ln["advances"]
                                    if a["compressed_word"])
                    rule_NOT = [a for ln in res["lines"]
                                for a in ln["advances"]
                                if (a["rule_match"] != "none"
                                    and not a["compressed_word"])]
                    rule_YES = [a for ln in res["lines"]
                                for a in ln["advances"]
                                if (a["rule_match"] != "none"
                                    and a["compressed_word"])]
                    print(f"\n  [P{pi} page{res['page']}] "
                          f"text={res['text'][:50]!r}", flush=True)
                    print(f"    yak={yak} compressed={yak_compr} "
                          f"rule_match_compressed={len(rule_YES)} "
                          f"rule_match_NOT_compressed={len(rule_NOT)}",
                          flush=True)
                    if rule_NOT:
                        for a in rule_NOT[:5]:
                            print(f"      ANOMALY [{a['i']}] "
                                  f"{a['ch']!r} "
                                  f"prev={a['prev_ch']!r} "
                                  f"next={a['next_ch']!r} "
                                  f"rule={a['rule_match']} "
                                  f"adv={a['adv']} font={a['font']!r}",
                                  flush=True)
                    if rule_YES:
                        for a in rule_YES[:5]:
                            print(f"      MATCH [{a['i']}] "
                                  f"{a['ch']!r} "
                                  f"prev={a['prev_ch']!r} "
                                  f"next={a['next_ch']!r} "
                                  f"rule={a['rule_match']} "
                                  f"adv={a['adv']} r={a['ratio']}",
                                  flush=True)
                except Exception as e:
                    print(f"  P{pi} ERR: {e}", flush=True)
                    out.append({"para_idx": pi, "error": str(e)})
                    save({"page_1": [r for r in out
                                       if r.get("page") == 1
                                       or (r.get("para_idx") in P1_INDICES)],
                          "page_3": [r for r in out
                                       if r.get("page") == 3
                                       or (r.get("para_idx") in P3_INDICES)]})
            try:
                d.Close(SaveChanges=False)
            except Exception:
                pass
        finally:
            try:
                word.Quit()
            except Exception:
                pass
            time.sleep(2.0)
    return out


def main():
    all_indices = P1_INDICES + P3_INDICES
    results = measure_indices_with_restart(all_indices, batch_size=4)
    out = {
        "page_1": [r for r in results if r.get("page") == 1],
        "page_3": [r for r in results if r.get("page") == 3],
        "errors": [r for r in results if "error" in r],
    }
    save(out)
    print(f"\nWrote {RESULT_PATH}: "
          f"p1={len(out['page_1'])} p3={len(out['page_3'])} "
          f"errors={len(out['errors'])}", flush=True)


if __name__ == "__main__":
    main()
